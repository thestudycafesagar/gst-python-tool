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
    def __init__(self, app, file_path, manual_gstins=None, credentials=None):
        self.app = app
        self.file_path = file_path
        self.manual_gstins = manual_gstins or []
        self.credentials = credentials or []
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
        self.geometry("1100x850")
        self.worker = None
        self.gstin_list = []
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        # HEADER
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.header_frame.grid_columnconfigure(0, weight=1)

        # MAIN CONTENT: scrollable tool area
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        self.scroll_container = ctk.CTkScrollableFrame(self.main_frame, fg_color="transparent")
        self.scroll_container.grid(row=0, column=0, sticky="nsew", pady=0)
        self.scroll_container.grid_columnconfigure(0, weight=1)

        self.lbl_title = ctk.CTkLabel(self.header_frame, text="GST VERIFICATION TOOL",
                                      font=("Segoe UI", 24, "bold"))
        self.lbl_title.grid(row=0, column=0)

        self.lbl_subtitle = ctk.CTkLabel(self.header_frame, text="v3.1 | Pro Edition",
                                         font=("Segoe UI", 12), text_color="gray")
        self.lbl_subtitle.grid(row=1, column=0)

        # --- GSTIN Management Row ---
        self.frame_gstin = ctk.CTkFrame(self.scroll_container)
        self.frame_gstin.grid(row=0, column=0, padx=10, pady=(0, 6), sticky="ew")
        self.frame_gstin.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.frame_gstin, text="📋 GSTIN List", font=("Segoe UI", 13, "bold")).pack(anchor="center", padx=15, pady=(14, 8))
        gstin_row = ctk.CTkFrame(self.frame_gstin, fg_color="transparent")
        gstin_row.pack(anchor="center", pady=(0, 12))
        ctk.CTkButton(gstin_row, text="➕ Add GSTIN", width=130, height=34,
                      fg_color="#059669", hover_color="#047857",
                      font=("Segoe UI", 11, "bold"), command=self.add_gstin).pack(side="left", padx=6)
        ctk.CTkButton(gstin_row, text="📂 Load Data", width=120, height=34,
                      fg_color="#4338ca", hover_color="#3730a3",
                      font=("Segoe UI", 11, "bold"), command=self.load_gstins_from_db).pack(side="left", padx=6)
        ctk.CTkButton(gstin_row, text="🗑 Delete Data", width=125, height=34,
                      fg_color="#7C3AED", hover_color="#6D28D9",
                      font=("Segoe UI", 11, "bold"), command=self.view_gstin_data).pack(side="left", padx=6)
        ctk.CTkButton(gstin_row, text="▶ Watch Demo Video", width=155, height=34,
                      fg_color="#DC2626", hover_color="#B91C1C",
                      font=("Segoe UI", 11, "bold"), command=self.open_demo_link).pack(side="left", padx=6)
        self.lbl_gstin_count = ctk.CTkLabel(self.frame_gstin, text="No GSTINs loaded",
                                             font=("Segoe UI", 11), text_color="gray")
        self.lbl_gstin_count.pack(anchor="center", pady=(0, 12))

        # --- 2. RICH LOG & PROGRESS ---
        self.frame_log = ctk.CTkFrame(self.scroll_container)
        self.frame_log.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        
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

    def add_gstin(self):
        import sqlite3 as _sq, os as _os
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add GSTIN")
        dialog.geometry("420x220")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        card = ctk.CTkFrame(dialog, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(card, text="Add New GSTIN", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 12))
        ctk.CTkLabel(card, text="GSTIN Number", font=("Segoe UI", 12)).pack(anchor="w")
        ent = ctk.CTkEntry(card, placeholder_text="e.g. 27ABCDE1234F1Z5", height=36)
        ent.pack(fill="x", pady=(4, 16))
        ent.focus_set()
        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x")
        def _save():
            g = ent.get().strip().upper()
            if not g:
                messagebox.showwarning("Missing", "Enter a GSTIN number.", parent=dialog)
                return
            db_path = _os.path.join(_os.environ.get("APPDATA", _os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
            _os.makedirs(_os.path.dirname(db_path), exist_ok=True)
            try:
                conn = _sq.connect(db_path)
                conn.execute("CREATE TABLE IF NOT EXISTS gst_gstin_list (id INTEGER PRIMARY KEY AUTOINCREMENT, gstin TEXT UNIQUE)")
                conn.execute("INSERT OR REPLACE INTO gst_gstin_list (gstin) VALUES (?)", (g,))
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dialog)
                return
            if g not in self.gstin_list:
                self.gstin_list.append(g)
            self._refresh_gstin_count()
            self.log_message(f"✅ GSTIN added: {g}", "success")
            dialog.destroy()
        ctk.CTkButton(btn_row, text="Cancel", width=100, command=dialog.destroy).pack(side="right")
        ctk.CTkButton(btn_row, text="Save GSTIN", width=110, fg_color="#059669", hover_color="#047857", command=_save).pack(side="right", padx=(0, 8))
        dialog.bind("<Return>", lambda _e: _save())

    def load_gstins_from_db(self):
        import sqlite3 as _sq, os as _os
        db_path = _os.path.join(_os.environ.get("APPDATA", _os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
        try:
            conn = _sq.connect(db_path)
            rows = conn.execute("SELECT gstin FROM gst_gstin_list ORDER BY gstin").fetchall()
            conn.close()
        except Exception:
            rows = []
        if not rows:
            messagebox.showinfo("No GSTINs", "No GSTINs saved yet. Add via the Add GSTIN button.", parent=self)
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Load GSTINs")
        dialog.geometry("400x460")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        ctk.CTkLabel(dialog, text="Select GSTINs to Load",
                     font=("Segoe UI", 14, "bold")).pack(pady=(16, 8))
        sel_all_var = ctk.BooleanVar()
        vars_ = {}
        def _toggle_all():
            state = sel_all_var.get()
            for v in vars_.values():
                v.set(state)
        ctk.CTkCheckBox(dialog, text="Select All", variable=sel_all_var,
                        command=_toggle_all,
                        font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
        scroll = ctk.CTkScrollableFrame(dialog, height=300)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        for (g,) in rows:
            v = ctk.BooleanVar()
            ctk.CTkCheckBox(scroll, text=g, variable=v).pack(anchor="w", padx=10, pady=3)
            vars_[g] = v
        def _load():
            selected = [g for g, v in vars_.items() if v.get()]
            if not selected:
                messagebox.showwarning("No Selection", "Select at least one GSTIN.", parent=dialog)
                return
            self.gstin_list = selected
            self._refresh_gstin_count()
            self.log_message(f"📂 Loaded {len(selected)} GSTIN(s) from DB", "info")
            dialog.destroy()
        foot = ctk.CTkFrame(dialog, fg_color="transparent")
        foot.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(foot, text="Cancel", width=100, command=dialog.destroy).pack(side="right")
        ctk.CTkButton(foot, text="Load Selected", width=120,
                      fg_color="#4338ca", hover_color="#3730a3",
                      command=_load).pack(side="right", padx=(0, 8))

    def clear_gstins(self):
        self.gstin_list = []
        self._refresh_gstin_count()
        self.log_message("🗑 GSTIN list cleared", "info")

    def view_gstin_data(self):
        import sqlite3 as _sq, os as _os
        db_path = _os.path.join(_os.environ.get("APPDATA", _os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
        try:
            conn = _sq.connect(db_path)
            rows = conn.execute("SELECT id, gstin FROM gst_gstin_list ORDER BY gstin").fetchall()
            conn.close()
        except Exception:
            rows = []
        dialog = ctk.CTkToplevel(self)
        dialog.title("Saved GSTINs")
        dialog.geometry("480x500")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        ctk.CTkLabel(dialog, text="Saved GSTINs", font=("Segoe UI", 14, "bold")).pack(pady=(16, 8))
        scroll = ctk.CTkScrollableFrame(dialog, height=360)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        scroll.grid_columnconfigure(0, weight=1)
        def _rebuild():
            for w in scroll.winfo_children():
                w.destroy()
            try:
                c = _sq.connect(db_path)
                rs = c.execute("SELECT id, gstin FROM gst_gstin_list ORDER BY gstin").fetchall()
                c.close()
            except Exception:
                rs = []
            if not rs:
                ctk.CTkLabel(scroll, text="No GSTINs saved yet.",
                             font=("Segoe UI", 12), text_color="gray").pack(pady=30)
                return
            for rid, gnum in rs:
                row_f = ctk.CTkFrame(scroll, fg_color=("#f8fafc", "#273549"),
                                     corner_radius=8, border_width=1,
                                     border_color=("#e2e8f0", "#334155"))
                row_f.pack(fill="x", padx=4, pady=4)
                row_f.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(row_f, text=f"  {gnum}",
                             font=("Segoe UI", 13, "bold"),
                             anchor="w").grid(row=0, column=0, sticky="w", padx=12, pady=10)
                def _del(r=rid, n=gnum):
                    if not messagebox.askyesno("Delete", f"Delete GSTIN '{n}'?", parent=dialog):
                        return
                    try:
                        c2 = _sq.connect(db_path)
                        c2.execute("DELETE FROM gst_gstin_list WHERE id=?", (r,))
                        c2.commit()
                        c2.close()
                    except Exception:
                        pass
                    if n in self.gstin_list:
                        self.gstin_list.remove(n)
                    self._refresh_gstin_count()
                    _rebuild()
                ctk.CTkButton(row_f, text="Delete", width=70, height=28,
                              fg_color="#DC2626", hover_color="#B91C1C",
                              font=("Segoe UI", 11, "bold"),
                              command=_del).grid(row=0, column=1, padx=(0, 10))
        _rebuild()
        ctk.CTkButton(dialog, text="Close", width=100, command=dialog.destroy).pack(pady=(0, 12))

    def _refresh_gstin_count(self):
        n = len(self.gstin_list)
        if n == 0:
            self.lbl_gstin_count.configure(text="No GSTINs loaded", text_color="gray")
        else:
            self.lbl_gstin_count.configure(text=f"{n} GSTIN(s) ready", text_color="#059669")

    def open_output_folder(self):
        target = os.path.join(os.getcwd(), "GST Downloaded", "GST Verifier")
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showinfo("Info", "Output folder not found.")

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
        manual_gstins = list(self.gstin_list)

        if not manual_gstins:
            messagebox.showwarning("Warning", "Please add GSTINs first!")
            return
        
        self.reset_status_label()
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.grid()
        self.btn_open_folder.grid_remove()
        self.log_message("🚀 Starting Automation Thread...", "info")
        self.worker = GSTWorker(self, None, manual_gstins=manual_gstins)
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