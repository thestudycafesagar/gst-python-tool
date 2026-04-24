import sys
import os
import time
import threading
import pandas as pd
from io import BytesIO
from PIL import Image, ImageTk
import customtkinter as ctk  # Modern UI Library
import tkinter as tk # Standard TK for rich text support
from tkinter import filedialog, messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Shared Stealth Driver Import
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _ROOT not in sys.path: sys.path.insert(0, _ROOT)
from stealth_driver import create_chrome_driver, build_chrome_options

# --- CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue") 

# --- WORKER THREAD ---
class GSTWorker:
    def __init__(self, app, file_path, manual_gstins=None):
        self.app = app
        self.file_path = file_path
        self.manual_gstins = manual_gstins or []
        self.keep_running = True
        self.driver = None

    def log(self, message, tag=None):
        self.app.update_log_safe(message, tag)

    def run(self):
        self.log("🚀 Initializing Browser Engine...", "info")
        
        try:
            self.driver = create_chrome_driver(build_chrome_options())
            driver = self.driver
        except Exception as e:
            self.log(f"❌ Failed to start browser: {e}", "error")
            return

        try:
            if self.manual_gstins:
                gstin_list = [str(x).strip() for x in self.manual_gstins if str(x).strip()]
            else:
                # Load Excel
                df = pd.read_excel(self.file_path)

                # --- FIX: NORMALIZE COLUMN NAMES (Case-Insensitive) ---
                # Remove spaces and convert to UPPERCASE
                df.columns = df.columns.str.strip().str.upper()

                if 'GSTIN' not in df.columns:
                    self.log("❌ Error: Excel must have a column named 'GSTIN' (case-insensitive)", "error")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    self.driver = None
                    return

                gstin_list = df['GSTIN'].astype(str).unique().tolist()

            gstin_list = list(dict.fromkeys(gstin_list))
            results = []
            total = len(gstin_list)

            self.log(f"📂 Found {total} unique GSTINs. Starting Batch Process...", "info")

            for index, gstin in enumerate(gstin_list):
                if not self.keep_running: break
                
                self.log(f"\n🔍 Processing ({index+1}/{total}): {gstin}", "normal")
                self.app.update_progress((index / total))
                self.app.reset_status_label() 

                try:
                    driver.get("https://services.gst.gov.in/services/searchtp")
                    
                    # 1. Enter GSTIN
                    try:
                        input_box = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.ID, "for_gstin"))
                        )
                        input_box.clear()
                        input_box.send_keys(gstin)
                    except TimeoutException:
                        self.log("⚠️ Site timeout. Retrying...", "warning")
                        continue

                    # --- MANUAL CAPTCHA WAIT LOOP ---
                    self.log("👉 Please enter the captcha in the browser and click Search.", "info")
                    
                    search_completed = False
                    while self.keep_running and not search_completed:
                        try:
                            # Check for Success/Error markers
                            page_source = driver.page_source
                            if "Legal Name" in page_source:
                                search_completed = True
                                self.log("✅ Result detected!", "success")
                            elif "The GSTIN/UIN that you have entered is invalid" in page_source:
                                search_completed = True
                            elif "Enter valid characters" in page_source or "characters shown" in page_source:
                                # Still on page, wait more
                                pass
                            
                            if not search_completed:
                                time.sleep(1)
                        except Exception:
                            time.sleep(1)
                    

                    if not self.keep_running: break

                    # Verify Results
                    self.log("⏳ Verifying...", "normal")

                    # We'll retry verification a few times before deciding it's a wrong captcha
                    verify_attempt = 0
                    max_verify_attempts = 3
                    skip_this_gstin = False

                    while True:
                        try:
                            WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH,
                                    "//strong[contains(text(),'Legal Name')] | " +
                                    "//p[contains(text(),'Legal Name')] | " +
                                    "//div[contains(text(),'Enter valid characters')] | " +
                                    "//span[contains(text(),'Invalid')]"))
                            )
                            break
                        except:
                            self.log("⚠️ Response slow. Checking page...", "warning")
                            page_source = driver.page_source

                            # Check if GST is invalid (explicit error element)
                            is_invalid_gst = False
                            try:
                                err_msg = driver.find_element(By.XPATH, "//span[contains(text(), 'The GSTIN/UIN that you have entered is invalid')]")
                                if err_msg.is_displayed():
                                    is_invalid_gst = True
                            except:
                                pass

                            if is_invalid_gst:
                                self.log(f"🚫 INVALID GSTIN DETECTED: {gstin}", "fatal")
                                # Save debug snapshot to help diagnose false negatives
                                try:
                                    self.save_debug_snapshot(driver, gstin, tag="invalid")
                                except Exception:
                                    pass
                                self.app.show_invalid_gst_alert()
                                results.append({"GSTIN": gstin, "Status": "Invalid GSTIN"})
                                time.sleep(2)
                                skip_this_gstin = True
                                break

                            # If page still shows captcha prompt text, wait a few times before giving up
                            if "Enter valid characters" in page_source or "characters shown" in page_source:
                                verify_attempt += 1
                                if verify_attempt <= max_verify_attempts:
                                    self.log(f"⏳ Waiting for manual captcha submission... ({verify_attempt}/{max_verify_attempts})", "info")
                                    time.sleep(1.5)
                                    continue
                                else:
                                    self.log("❌ WRONG CAPTCHA! Retrying...", "error")
                                    # capture debug snapshot on captcha timeouts
                                    try:
                                        self.save_debug_snapshot(driver, gstin, tag=f"captcha_timeout_{verify_attempt}")
                                    except Exception:
                                        pass
                                    time.sleep(1.5)
                                    skip_this_gstin = True
                                    break

                            # If result content appears despite timeout, proceed to extract
                            if "Legal Name" in page_source:
                                break

                            # Unknown state: retry a few times then skip
                            verify_attempt += 1
                            if verify_attempt <= max_verify_attempts:
                                self.log("⚠️ Unknown state. Retrying...", "warning")
                                time.sleep(1.5)
                                continue
                            else:
                                self.log("⚠️ Unable to determine page state. Skipping this GSTIN.", "warning")
                                skip_this_gstin = True
                                break

                    if skip_this_gstin:
                        continue

                    # Re-evaluate page and extract if result is present
                    page_source = driver.page_source

                    # Robust visible-element checks to avoid false negatives from template text
                    is_invalid_visible = False
                    try:
                        invalid_elems = driver.find_elements(By.XPATH, "//span[contains(text(), 'The GSTIN/UIN that you have entered is invalid')]")
                        for el in invalid_elems:
                            try:
                                if el.is_displayed():
                                    is_invalid_visible = True
                                    break
                            except:
                                continue
                    except:
                        is_invalid_visible = False

                    # Check for obvious success markers being visible on page
                    is_success_visible = False
                    try:
                        success_elems = driver.find_elements(By.XPATH, "//*[contains(text(),'Search Result based on GSTIN')]|//strong[contains(text(),'Legal Name')]|//p[contains(text(),'Legal Name of Business')]")
                        for el in success_elems:
                            try:
                                if el.is_displayed() and el.text.strip():
                                    is_success_visible = True
                                    break
                            except:
                                continue
                    except:
                        is_success_visible = False

                    if is_success_visible:
                        # proceed to extract
                        self.app.reset_status_label()
                        self.log("✅ Captcha Correct! Extracting Data...", "success")

                        row_data = {"GSTIN": gstin}

                        def get_text(label):
                            xpaths = [
                                f"//p[contains(text(),'{label}')]/following-sibling::p",
                                f"//strong[contains(text(),'{label}')]/../following-sibling::p",
                                f"//*[contains(text(),'{label}')]/following::p[1]"
                            ]
                            for xpath in xpaths:
                                try:
                                    el = driver.find_element(By.XPATH, xpath)
                                    if el.text.strip(): return el.text.strip()
                                except: continue
                            return "N/A"

                        row_data["Legal Name"] = get_text("Legal Name of Business")
                        row_data["Trade Name"] = get_text("Trade Name")
                        row_data["Effective Date"] = get_text("Effective Date of registration")
                        row_data["Constitution"] = get_text("Constitution of Business")
                        row_data["Taxpayer Type"] = get_text("Taxpayer Type")
                        row_data["Address"] = get_text("Principal Place of Business")

                        # Status + Suspension Fix
                        try:
                            siblings = driver.find_elements(By.XPATH, "//strong[contains(text(),'GSTIN / UIN')]/parent::p/following-sibling::p")
                            if siblings:
                                status_text = siblings[0].text.strip()
                                if len(siblings) > 1:
                                    date_text = siblings[1].text.strip()
                                    if "Effective" in date_text:
                                        status_text += f" {date_text}"
                                row_data["Status"] = status_text
                            else:
                                row_data["Status"] = get_text("GSTIN / UIN")
                        except:
                            row_data["Status"] = get_text("GSTIN / UIN")

                        def get_list_text(xpath):
                            try:
                                items = driver.find_elements(By.XPATH, xpath)
                                return ", ".join([x.text.strip() for x in items if x.text.strip()])
                            except: return "N/A"

                        row_data["Admin Office"] = get_list_text("//strong[contains(text(),'Administrative Office')]/parent::p/following-sibling::ul//li")
                        row_data["Other Office"] = get_list_text("//strong[contains(text(),'Other Office')]/parent::p/following-sibling::ul//li")
                        row_data["Core Business"] = get_text("Nature Of Core Business Activity")
                        row_data["Business Activities"] = get_list_text("//p[contains(text(),'Nature of Business Activities')]/ancestor::div[@class='panel-heading']/following-sibling::div//li")

                        try:
                            hsn_rows = []
                            rows = driver.find_elements(By.XPATH, "//div[contains(@class,'table-responsive')]//table[contains(@class,'tbl')]//tbody//tr")
                            for row in rows:
                                cols = row.find_elements(By.TAG_NAME, "td")
                                if len(cols) >= 2:
                                    t1, t2 = cols[0].text.strip(), cols[1].text.strip()
                                    if t1 or t2: hsn_rows.append(f"{t1}-{t2}")
                                if len(cols) >= 4:
                                    t3, t4 = cols[2].text.strip(), cols[3].text.strip()
                                    if t3 or t4: hsn_rows.append(f"{t3}-{t4}")
                            row_data["Goods & Services"] = " | ".join(hsn_rows)
                        except:
                            row_data["Goods & Services"] = "N/A"

                        self.log("📥 Extracting Filing Data...", "normal")
                        try:
                            show_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "filingTable")))
                            driver.execute_script("arguments[0].click();", show_btn)

                            search_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary.srchbtn")))
                            driver.execute_script("arguments[0].click();", search_btn)
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//h4[contains(text(),'GSTR3B')]")))
                            time.sleep(1)

                            def get_filing_history(header_text):
                                history = []
                                try:
                                    xpath = f"//h4[contains(text(),'{header_text}')]/ancestor::div[@class='table-responsive']//tbody/tr"
                                    rows = driver.find_elements(By.XPATH, xpath)
                                    for row in rows[:5]:
                                        cols = row.find_elements(By.TAG_NAME, "td")
                                        if len(cols) >= 4:
                                            history.append(f"[{cols[1].text}-{cols[0].text}: {cols[3].text} on {cols[2].text}]")
                                    return " | ".join(history)
                                except: return "Not Found"

                            row_data["GSTR-3B History"] = get_filing_history("GSTR3B")
                            row_data["GSTR-1 History"] = get_filing_history("GSTR-1")
                        except:
                            row_data["GSTR-3B History"] = "Hidden/Error"
                            row_data["GSTR-1 History"] = "Hidden/Error"

                        self.log("📥 Extracting Return Frequency...", "normal")
                        try:
                            freq_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "profileTable")))
                            driver.execute_script("arguments[0].click();", freq_btn)
                            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'exp')]")))
                            time.sleep(1)
                            freq_data = []
                            freq_rows = driver.find_elements(By.XPATH, "//table[contains(@class,'exp')]//tbody//tr")
                            for row in freq_rows:
                                cols = row.find_elements(By.TAG_NAME, "td")
                                if len(cols) >= 9:
                                    yr = cols[0].text.strip()
                                    q1, f1 = cols[1].text.strip(), cols[2].text.strip()
                                    q2, f2 = cols[3].text.strip(), cols[4].text.strip()
                                    q3, f3 = cols[5].text.strip(), cols[6].text.strip()
                                    q4, f4 = cols[7].text.strip(), cols[8].text.strip()
                                    freq_data.append(f"[{yr}: {q1}({f1}), {q2}({f2}), {q3}({f3}), {q4}({f4})]")
                            row_data["Return Frequency"] = " | ".join(freq_data)
                        except:
                            row_data["Return Frequency"] = "Not Available"

                        self.log(f"✅ Success: {row_data['Legal Name']}", "success")
                        results.append(row_data)
                        continue
                    else:
                        # Not clearly successful; if invalid is visible mark invalid, else attempt a gentle extraction
                        if is_invalid_visible:
                            self.log(f"🚫 INVALID GSTIN DETECTED: {gstin}", "fatal")
                            try:
                                self.save_debug_snapshot(driver, gstin, tag="invalid_final")
                            except Exception:
                                pass
                            self.app.show_invalid_gst_alert()
                            results.append({"GSTIN": gstin, "Status": "Invalid GSTIN"})
                            time.sleep(2)
                            continue
                        else:
                            # Try extracting Legal Name directly as a fallback
                            try:
                                candidate = "N/A"
                                try:
                                    el = driver.find_element(By.XPATH, "//strong[contains(text(),'Legal Name')]/../following-sibling::p")
                                    if el and el.text.strip():
                                        candidate = el.text.strip()
                                except:
                                    pass

                                if candidate != "N/A":
                                    # minimal extraction succeeded
                                    row_data = {"GSTIN": gstin, "Legal Name": candidate, "Status": "Partial"}
                                    results.append(row_data)
                                    self.log(f"✅ Partial extraction succeeded for {gstin}", "success")
                                    continue
                                else:
                                    self.log("⚠️ Unable to determine page state. Marking as Invalid.", "warning")
                                    try:
                                        self.save_debug_snapshot(driver, gstin, tag="unknown_state")
                                    except Exception:
                                        pass
                                    results.append({"GSTIN": gstin, "Status": "Unknown/Skipped"})
                                    continue
                            except Exception:
                                results.append({"GSTIN": gstin, "Status": "Error"})
                                continue

                except Exception as e:
                    self.log(f"❌ Error processing {gstin}: {str(e)}", "error")
                    results.append({"GSTIN": gstin, "Status": "Error"})

            if not self.keep_running:
                self.log("🛑 Process stopped by user.", "warning")
                if results:
                    timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
                    base_dir = os.path.join(os.getcwd(), "GST Downloaded", "GST Verifier", "reports")
                    os.makedirs(base_dir, exist_ok=True)
                    output_file = os.path.join(base_dir, f"GST_Report_{timestamp}.xlsx")
                    self.log(f"📊 Partial results count: {len(results)}. Saving to: {output_file}", "info")
                    try:
                        pd.DataFrame(results).to_excel(output_file, index=False)
                        self.app.process_finished(f"Stopped by user. Partial report saved in GST Verifier reports folder")
                    except Exception as e:
                        # Try CSV fallback
                        try:
                            csv_file = output_file.replace('.xlsx', '.csv')
                            pd.DataFrame(results).to_csv(csv_file, index=False)
                            self.log(f"⚠️ Excel save failed ({e}). CSV saved at {csv_file}", "warning")
                            self.app.process_finished(f"Stopped by user. Partial CSV saved in GST Verifier reports folder")
                        except Exception as e2:
                            self.app.log_message(f"CRITICAL ERROR: Could not save report. {e}; {e2}", "fatal")
                            self.app.process_finished("Stopped by user.")
                else:
                    self.app.process_finished("Stopped by user.")
                return
            
            # Export
            self.app.update_progress(1.0)
            timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
            base_dir = os.path.join(os.getcwd(), "GST Downloaded", "GST Verifier", "reports")
            os.makedirs(base_dir, exist_ok=True)
            output_file = os.path.join(base_dir, f"GST_Report_{timestamp}.xlsx")
            self.log(f"📊 Final results count: {len(results)}. Saving to: {output_file}", "info")

            try:
                pd.DataFrame(results).to_excel(output_file, index=False)
                self.app.process_finished(f"Completed! Saved in GST Verifier reports folder")
            except Exception as e:
                # Try CSV fallback
                try:
                    csv_file = output_file.replace('.xlsx', '.csv')
                    pd.DataFrame(results).to_csv(csv_file, index=False)
                    self.log(f"⚠️ Excel save failed ({e}). CSV saved at {csv_file}", "warning")
                    self.app.process_finished(f"Completed! Saved CSV in GST Verifier reports folder")
                except Exception as e2:
                    self.app.log_message(f"CRITICAL ERROR: Could not save report. {e}; {e2}", "fatal")
                    self.app.process_finished("Failed to save")
            
        except Exception as e:
            self.log(f"CRITICAL ERROR: {e}", "fatal")
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None

    def receive_captcha_input(self, text):
        self.user_captcha_response = text

    def save_debug_snapshot(self, driver, gstin, tag="debug"):
        """Save a screenshot and page HTML to GST Downloaded/GST Verifier/debug for analysis."""
        try:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            debug_dir = os.path.join(os.getcwd(), "GST Downloaded", "GST Verifier", "debug")
            os.makedirs(debug_dir, exist_ok=True)
            safe_gstin = str(gstin).replace('/', '_').replace('\\', '_')
            img_path = os.path.join(debug_dir, f"{tag}_{safe_gstin}_{timestamp}.png")
            html_path = os.path.join(debug_dir, f"{tag}_{safe_gstin}_{timestamp}.html")
            try:
                # prefer save_screenshot (widely supported)
                driver.save_screenshot(img_path)
            except Exception:
                try:
                    open(img_path, 'wb').write(driver.get_screenshot_as_png())
                except Exception:
                    pass
            try:
                with open(html_path, 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
            except Exception:
                pass
            # Log saved paths
            self.log(f"🔍 Debug snapshot saved: {img_path}", "info")
            self.log(f"🔍 Debug HTML saved: {html_path}", "info")
        except Exception as e:
            try:
                self.log(f"⚠️ Failed to save debug snapshot: {e}", "warning")
            except Exception:
                pass

    def stop(self):
        self.keep_running = False
        self.user_captcha_response = ""
        self.is_waiting_for_captcha = False
        try:
            if self.driver:
                self.driver.quit()
                self.driver = None
        except Exception:
            pass


# --- MODERN GUI CLASS (CustomTkinter + Rich Text) ---
class GSTApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("GST Bulk Verification Pro")
        self.geometry("700x850")
        self.worker = None
        self.manual_credentials = []
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        # HEADER
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        # CONTENT AREA (SCROLLABLE)
        self.scroll_container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll_container.grid_columnconfigure(0, weight=1)
        
        self.lbl_title = ctk.CTkLabel(self.header_frame, text="GST VERIFICATION TOOL", 
                                      font=("Segoe UI", 24, "bold"))
        self.lbl_title.pack(side="left")
        
        self.lbl_subtitle = ctk.CTkLabel(self.header_frame, text="v3.1 | Pro Edition", 
                                         font=("Segoe UI", 12), text_color="gray")
        self.lbl_subtitle.pack(side="left", padx=10, pady=(10, 0))

        # --- 1. FILE UPLOAD ---
        self.frame_file = ctk.CTkFrame(self.scroll_container)
        self.frame_file.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.lbl_step1 = ctk.CTkLabel(self.frame_file, text="STEP 1: Upload Data", font=("Segoe UI", 14, "bold"))
        self.lbl_step1.pack(anchor="w", padx=15, pady=(10, 5))
        
        self.file_entry = ctk.CTkEntry(self.frame_file, placeholder_text="Select Excel file or add GSTIN manually...", width=400)
        self.file_entry.pack(side="left", padx=15, pady=(0, 15), expand=True, fill="x")
        
        self.btn_browse = ctk.CTkButton(self.frame_file, text="Browse", command=self.browse_excel, 
                                        width=90, height=28, font=("Segoe UI", 12, "bold"))
        self.btn_browse.pack(side="right", padx=(0, 15), pady=(0, 15))

        # --- Button Row Tools (Internal Actions) ---
        self.btn_row_actions = ctk.CTkFrame(self.scroll_container, fg_color="transparent")
        self.btn_row_actions.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")

        self.btn_add_id = ctk.CTkButton(self.btn_row_actions, text="➕ Add GSTIN", command=self.add_id_password,
                         fg_color="#059669", hover_color="#047857", height=28, font=("Segoe UI", 12, "bold"), width=120)
        self.btn_add_id.pack(side="left", padx=(0, 10))

        self.btn_view_id = ctk.CTkButton(self.btn_row_actions, text="👁 View", command=self.view_saved_user,
                         fg_color="#475569", hover_color="#334155", height=28, font=("Segoe UI", 11, "bold"), width=80)
        self.btn_view_id.pack(side="left", padx=(0, 10))

        self.btn_delete_id = ctk.CTkButton(self.btn_row_actions, text="🗑 Delete", command=self.delete_saved_user,
                           fg_color="#7C3AED", hover_color="#6D28D9", height=28, font=("Segoe UI", 11, "bold"), width=80)
        self.btn_delete_id.pack(side="left", padx=(0, 10))

        self.btn_sample = ctk.CTkButton(self.btn_row_actions, text="📥 Download Sample", command=self.download_sample,
                           fg_color="#2563EB", hover_color="#1D4ED8", height=28, font=("Segoe UI", 12, "bold"), width=160)
        self.btn_sample.pack(side="left", padx=(0, 10))

        self.btn_demo = ctk.CTkButton(self.btn_row_actions, text="▶ View Demo", command=self.open_demo_link, 
                          fg_color="#DC2626", hover_color="#B91C1C", height=28, font=("Segoe UI", 12, "bold"), width=120)
        self.btn_demo.pack(side="left")
        
        self.btn_view_id.configure(state="disabled")
        self.btn_delete_id.configure(state="disabled")

        # --- 2. RICH LOG & PROGRESS ---
        self.frame_log = ctk.CTkFrame(self.scroll_container)
        self.frame_log.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        
        self.lbl_step2 = ctk.CTkLabel(self.frame_log, text="Process Log", font=("Segoe UI", 14, "bold"))
        self.lbl_step2.pack(anchor="w", padx=15, pady=(10, 5))
        
        # RICH TEXT BOX (Using Standard TK for Colors)
        self.log_text = tk.Text(self.frame_log, height=15, bg="#0F172A", fg="#CBD5E1", 
                                font=("Consolas", 11), borderwidth=0, highlightthickness=0)
        self.log_text.pack(fill="both", expand=True, padx=15, pady=5)
        
        # Define Tags for Colors
        self.log_text.tag_config("normal", foreground="#CBD5E1")
        self.log_text.tag_config("info", foreground="#3B82F6") # Blue
        self.log_text.tag_config("success", foreground="#10B981") # Green
        self.log_text.tag_config("warning", foreground="#F59E0B") # Orange
        self.log_text.tag_config("error", foreground="#EF4444") # Red
        self.log_text.tag_config("fatal", foreground="#DC2626", font=("Consolas", 12, "bold")) # Big Red
        
        self.progress_bar = ctk.CTkProgressBar(self.frame_log)
        self.progress_bar.pack(fill="x", padx=15, pady=15)
        self.progress_bar.set(0)

        # --- 3. ACTION BUTTON (FIXED AT BOTTOM) ---
        btn_footer = ctk.CTkFrame(self, fg_color="transparent")
        btn_footer.grid(row=2, column=0, padx=10, pady=(10, 20), sticky="ew")
        btn_footer.grid_columnconfigure(0, weight=1)
        
        self.btn_start = ctk.CTkButton(btn_footer, text="START AUTOMATION", font=("Segoe UI", 16, "bold"), height=50, command=self.start_process, fg_color="#2563EB", hover_color="#1D4ED8")
        self.btn_start.grid(row=0, column=0, sticky="ew")
        
        self.btn_stop = ctk.CTkButton(btn_footer, text="⏹ STOP", font=("Segoe UI", 16, "bold"), height=50, command=self.stop_process, fg_color="#475569", hover_color="#334155")
        self.btn_stop.grid(row=0, column=1, padx=(10, 0))
        self.btn_stop.grid_remove()

        self.btn_open_folder = ctk.CTkButton(btn_footer, text="📂 OPEN OUTPUT FOLDER", font=("Segoe UI", 16, "bold"), height=50, command=self.open_output_folder, fg_color="#64748B", hover_color="#475569")
        self.btn_open_folder.grid(row=0, column=2, padx=(10, 0))
        self.btn_open_folder.grid_remove()
    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://youtu.be/RAwvIz1RU-w")

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filename)
            self.manual_credentials = []
            self._refresh_manual_controls()
            self.log_message(f"✅ File loaded: {filename.split('/')[-1]}", "info")

    def download_sample(self):
        try:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile="GST_Verifier_Sample.xlsx",
                title="Save Sample Template"
            )
            if not save_path:
                return

            df = pd.DataFrame({"GSTIN": ["27ABCDE1234F1Z5", "07AAACR1234A1Z1"]})
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Sample file saved successfully at:\n{save_path}")
            self.log_message(f"📥 Sample file downloaded: {os.path.basename(save_path)}", "success")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save sample file: {e}")
            self.log_message(f"❌ Failed to download sample: {e}", "error")

    def open_output_folder(self):
        target = os.path.join(os.getcwd(), "GST Downloaded", "GST Verifier")
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showinfo("Info", "Output folder not found.")

    def _get_saved_user_id(self):
        if not self.manual_credentials:
            return ""
        return str(self.manual_credentials[0].get("Username", "")).strip()

    def _refresh_manual_controls(self):
        has_manual = bool(self.manual_credentials)
        self.btn_view_id.configure(state="normal" if has_manual else "disabled")
        self.btn_delete_id.configure(state="normal" if has_manual else "disabled")
        if has_manual:
            user_id = self._get_saved_user_id()
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, f"Selected GSTIN: {user_id}")

    def view_saved_user(self):
        user_id = self._get_saved_user_id()
        if not user_id:
            messagebox.showinfo("Info", "No saved GSTIN found.")
            return
        messagebox.showinfo("Saved GSTIN", f"Current GSTIN: {user_id}")

    def delete_saved_user(self):
        user_id = self._get_saved_user_id()
        if not user_id:
            messagebox.showinfo("Info", "No saved GSTIN found.")
            return
        if not messagebox.askyesno("Delete GSTIN", f"Delete saved GSTIN {user_id}?"):
            return
        self.manual_credentials = []
        self.file_entry.delete(0, "end")
        self._refresh_manual_controls()
        messagebox.showinfo("Deleted", "Saved GSTIN deleted successfully.")

    def add_id_password(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add GSTIN")
        dialog.geometry("420x180")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        card = ctk.CTkFrame(dialog, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(card, text="GSTIN").pack(anchor="w")
        ent_user = ctk.CTkEntry(card, placeholder_text="Enter GSTIN (e.g., 27ABCDE1234F1Z5)")
        ent_user.pack(fill="x", pady=(4, 14))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x")

        def _save():
            username = (ent_user.get() or "").strip().upper()
            if not username:
                messagebox.showerror("Missing Data", "Please enter GSTIN", parent=dialog)
                return

            if len(username) != 15 or not username.isalnum():
                messagebox.showerror(
                    "Invalid GSTIN",
                    "Please enter a valid 15-character GSTIN.\nExample: 27ABCDE1234F1Z5",
                    parent=dialog,
                )
                return

            existing_user = self._get_saved_user_id()
            if existing_user and not messagebox.askyesno(
                "Overwrite GSTIN",
                "Your previous GSTIN will be overwritten with this.",
                parent=dialog
            ):
                return

            self.manual_credentials = [{"Username": username}]
            self._refresh_manual_controls()
            messagebox.showinfo("Added", f"GSTIN saved: {username}", parent=dialog)
            dialog.destroy()

        ctk.CTkButton(btn_row, text="Cancel", width=110, command=dialog.destroy).pack(side="right")
        ctk.CTkButton(btn_row, text="Add", width=110, command=_save).pack(side="right", padx=(0, 8))

        ent_user.focus_set()
        dialog.bind("<Return>", lambda _e: _save())

    def update_log_safe(self, message, tag="normal"):
        self.after(0, lambda: self.log_message(message, tag))

    def log_message(self, message, tag="normal"):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def update_progress(self, val):
        self.after(0, lambda: self.progress_bar.set(val))

    def reset_status_label(self):
        pass # Optional: Add status label if needed

    def show_invalid_gst_alert(self):
        pass

    def start_process(self):
        file_path = self.file_entry.get()
        manual_gstins = [
            row.get("Username", "").strip()
            for row in self.manual_credentials
            if row.get("Username", "").strip()
        ]

        if not file_path and not manual_gstins:
            messagebox.showwarning("Warning", "Please add GSTIN first!")
            return
        
        self.reset_status_label()
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.grid()
        self.btn_open_folder.grid_remove()
        self.log_message("🚀 Starting Automation Thread...", "info")
        self.worker = GSTWorker(self, file_path, manual_gstins=manual_gstins)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if not self.worker:
            return
        self.worker.stop()
        self.btn_stop.configure(state="disabled", text="STOPPED")
        self.update_log_safe("🛑 Chrome browser closed.", "warning")
        self.update_log_safe("🛑 Process stopped by user.", "warning")

    def update_log_safe(self, message, tag="normal"):
        self.after(0, lambda: self.log_message(message, tag))

    def log_message(self, message, tag="normal"):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def update_progress(self, val):
        self.after(0, lambda: self.progress_bar.set(val))

    def reset_status_label(self):
        pass

    def show_invalid_gst_alert(self):
        pass

    def process_finished(self, msg):
        def finish():
            is_stopped = "stopped" in (msg or "").lower()
            self.log_message(f"\n🎉 DONE: {msg}", "warning" if is_stopped else "success")
            self.btn_start.configure(state="normal", text="STOPPED" if is_stopped else "START AUTOMATION")
            self.btn_stop.grid_remove()
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            if not is_stopped:
                self.btn_open_folder.grid()
            messagebox.showinfo("Info", msg)
            if is_stopped:
                self.after(1200, lambda: self.btn_start.configure(text="START AUTOMATION"))
        self.after(0, finish)

    def on_closing(self):
        if self.worker:
            self.worker.stop()
        self.destroy()

if __name__ == "__main__":
    app = GSTApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()