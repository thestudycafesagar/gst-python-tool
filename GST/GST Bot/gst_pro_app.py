import sys
import time
import threading
import pandas as pd
from io import BytesIO
from PIL import Image, ImageTk
import customtkinter as ctk  # Modern UI Library
import tkinter as tk # Standard TK for rich text support
from tkinter import filedialog, messagebox

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

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
        self.user_captcha_response = None
        self.is_waiting_for_captcha = False
        self.driver = None

    def log(self, message, tag=None):
        self.app.update_log(message, tag)

    def run(self):
        self.log("🚀 Initializing Browser Engine...", "info")
        
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized") 
        options.add_argument("--disable-blink-features=AutomationControlled") 
        
        try:
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            self.driver.maximize_window()
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                    window.navigator.chrome = { runtime: {} };
                    Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
                    Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
                """
            })
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

                    # --- CAPTCHA RETRY LOOP ---
                    while self.keep_running:
                        self.log("📷 Fetching Captcha...", "normal")
                        try:
                            # Wait for overlay to disappear
                            try:
                                WebDriverWait(driver, 2).until(
                                    EC.invisibility_of_element_located((By.CLASS_NAME, "dimmer-holder"))
                                )
                            except: pass

                            captcha_img = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "imgCaptcha"))
                            )
                        except:
                            try:
                                captcha_img = driver.find_element(By.XPATH, "//img[contains(@src, 'captcha')]")
                            except:
                                self.log("⚠️ Could not find captcha image. Reloading...", "warning")
                                driver.refresh()
                                continue

                        driver.execute_script("arguments[0].scrollIntoView();", captcha_img)
                        time.sleep(0.5) 
                        png_data = captcha_img.screenshot_as_png
                        
                        # Ask User
                        self.user_captcha_response = None
                        self.is_waiting_for_captcha = True
                        self.app.show_captcha_popup(png_data, gstin) 
                        
                        while self.user_captcha_response is None and self.keep_running:
                            time.sleep(0.1)
                        
                        self.is_waiting_for_captcha = False
                        if not self.keep_running: break

                        # Submit
                        self.log("📨 Submitting...", "normal")
                        try:
                            captcha_box = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.ID, "fo-captcha"))
                            )
                            captcha_box.clear()
                            captcha_box.send_keys(self.user_captcha_response)
                            
                            search_btn = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.ID, "lotsearch"))
                            )
                            driver.execute_script("arguments[0].click();", search_btn)
                            
                        except Exception as e:
                            self.log(f"❌ Error submitting: {e}", "error")
                            break 

                        # Verify Results
                        self.log("⏳ Verifying...", "normal")
                        
                        # Wait for ANY outcome
                        try:
                            WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, 
                                    "//strong[contains(text(),'Legal Name')] | " +
                                    "//p[contains(text(),'Legal Name')] | " +
                                    "//div[contains(text(),'Enter valid characters')] | " +
                                    "//span[contains(text(),'Invalid')]")) 
                            )
                        except:
                            self.log("⚠️ Response slow. Checking page...", "warning")

                        page_source = driver.page_source

                        # --- CRITICAL FIX: CHECK VISIBILITY OF ERROR MESSAGE ---
                        is_invalid_gst = False
                        try:
                            # Look for the specific error span
                            err_msg = driver.find_element(By.XPATH, "//span[contains(text(), 'The GSTIN/UIN that you have entered is invalid')]")
                            if err_msg.is_displayed():
                                is_invalid_gst = True
                        except:
                            pass # Element not found, so it's not invalid

                        # --- CASE 1: INVALID GSTIN (Fatal) ---
                        if is_invalid_gst:
                            self.log(f"🚫 INVALID GSTIN DETECTED: {gstin}", "fatal")
                            self.app.show_invalid_gst_alert()
                            results.append({"GSTIN": gstin, "Status": "Invalid GSTIN"})
                            time.sleep(2)
                            break 
                        
                        # --- CASE 2: WRONG CAPTCHA (Retry) ---
                        elif "Enter valid characters" in page_source or "characters shown" in page_source:
                            self.log("❌ WRONG CAPTCHA! Retrying...", "error")
                            time.sleep(1.5) 
                            continue 
                        
                        # --- CASE 3: SUCCESS ---
                        elif "Legal Name" in page_source:
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
                            break 
                        
                        else:
                            self.log("⚠️ Unknown state. Retrying...", "warning")
                            continue

                except Exception as e:
                    self.log(f"❌ Error processing {gstin}: {str(e)}", "error")
                    results.append({"GSTIN": gstin, "Status": "Error"})

            if not self.keep_running:
                self.log("🛑 Process stopped by user.", "warning")
                if results:
                    timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
                    output_file = f"GST_Report_{timestamp}.xlsx"
                    try:
                        pd.DataFrame(results).to_excel(output_file, index=False)
                        self.app.process_finished(f"Stopped by user. Partial report saved as {output_file}")
                    except Exception as e:
                        self.app.log_message(f"CRITICAL ERROR: Could not save report. {e}", "fatal")
                        self.app.process_finished("Stopped by user.")
                else:
                    self.app.process_finished("Stopped by user.")
                return
            
            # Export
            self.app.update_progress(1.0)
            timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
            output_file = f"GST_Report_{timestamp}.xlsx"
            
            try:
                pd.DataFrame(results).to_excel(output_file, index=False)
                self.app.process_finished(f"Completed! Saved as {output_file}")
            except Exception as e:
                self.app.log_message(f"CRITICAL ERROR: Could not save report. {e}", "fatal")
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
        self.geometry("700x800")
        self.worker = None
        self.manual_credentials = []
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) 

        # --- HEADER ---
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        self.lbl_title = ctk.CTkLabel(self.header_frame, text="GST VERIFICATION TOOL", 
                                      font=("Segoe UI", 24, "bold"))
        self.lbl_title.pack(side="left")
        
        self.lbl_subtitle = ctk.CTkLabel(self.header_frame, text="v3.1 | Pro Edition", 
                                         font=("Segoe UI", 12), text_color="gray")
        self.lbl_subtitle.pack(side="left", padx=10, pady=(10, 0))

        # --- 1. FILE UPLOAD ---
        self.frame_file = ctk.CTkFrame(self)
        self.frame_file.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.lbl_step1 = ctk.CTkLabel(self.frame_file, text="STEP 1: Upload Data", font=("Segoe UI", 14, "bold"))
        self.lbl_step1.pack(anchor="w", padx=15, pady=(10, 5))
        
        self.file_entry = ctk.CTkEntry(self.frame_file, placeholder_text="Add ID/Password manually...", width=400)
        self.file_entry.pack(side="left", padx=15, pady=(0, 15), expand=True, fill="x")
        
        self.btn_demo = ctk.CTkButton(self.frame_file, text="▶ View Demo", command=self.open_demo_link, fg_color="#DC2626", hover_color="#B91C1C", height=28, font=("Segoe UI", 12, "bold"))
        self.btn_demo.pack(side="right", padx=(0, 5), pady=(0, 15))
        self.btn_download = ctk.CTkButton(self.frame_file, text="➕ Add ID Password", command=self.add_id_password, fg_color="#059669", hover_color="#047857", height=28, font=("Segoe UI", 12, "bold"))
        self.btn_download.pack(side="right", padx=15, pady=(0, 15))
        self.btn_view_id = ctk.CTkButton(self.frame_file, text="👁 View ID", command=self.view_saved_user,
                         fg_color="#475569", hover_color="#334155", height=28,
                         font=("Segoe UI", 11, "bold"))
        self.btn_view_id.pack(side="right", padx=(0, 5), pady=(0, 15))
        self.btn_delete_id = ctk.CTkButton(self.frame_file, text="🗑 Delete ID", command=self.delete_saved_user,
                           fg_color="#7C3AED", hover_color="#6D28D9", height=28,
                           font=("Segoe UI", 11, "bold"))
        self.btn_delete_id.pack(side="right", padx=(0, 5), pady=(0, 15))
        self.btn_view_id.configure(state="disabled")
        self.btn_delete_id.configure(state="disabled")

        # --- 2. RICH LOG & PROGRESS ---
        self.frame_log = ctk.CTkFrame(self)
        self.frame_log.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        
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

        # --- 3. CAPTCHA SECTION ---
        self.frame_captcha = ctk.CTkFrame(self, fg_color="#1E293B", border_width=2, border_color="#444")
        self.frame_captcha.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        
        self.lbl_captcha_title = ctk.CTkLabel(self.frame_captcha, text="CAPTCHA ACTION REQUIRED", font=("Segoe UI", 14, "bold"), text_color="gray")
        self.lbl_captcha_title.pack(pady=(10, 5))
        
        self.lbl_image = ctk.CTkLabel(self.frame_captcha, text="[Waiting for Process...]", width=200, height=80, fg_color="#1E293B", corner_radius=10)
        self.lbl_image.pack(pady=5)
        
        self.entry_captcha = ctk.CTkEntry(self.frame_captcha, placeholder_text="Enter Code", justify='center', width=200, font=("Segoe UI", 16))
        self.entry_captcha.pack(pady=5)
        self.entry_captcha.bind('<Return>', lambda event: self.submit_captcha())
        self.entry_captcha.configure(state="disabled")
        
        self.btn_submit = ctk.CTkButton(self.frame_captcha, text="SUBMIT CAPTCHA", command=self.submit_captcha, state="disabled", fg_color="gray")
        self.btn_submit.pack(pady=(5, 5))
        self.cap_stop_btn = ctk.CTkButton(self.frame_captcha, text="⏹ STOP PROCESS", fg_color="#475569", hover_color="#334155",
                                          height=35, width=200, font=("Segoe UI", 11, "bold"), command=self.stop_process)
        self.cap_stop_btn.pack(pady=(5, 15))

        # --- 4. ACTION BUTTON ---
        btn_footer = ctk.CTkFrame(self, fg_color="transparent")
        btn_footer.grid(row=4, column=0, padx=20, pady=(10, 20), sticky="ew")
        btn_footer.grid_columnconfigure(0, weight=1)
        self.btn_start = ctk.CTkButton(btn_footer, text="START AUTOMATION", font=("Segoe UI", 16, "bold"), height=50, command=self.start_process, fg_color="#2563EB", hover_color="#1D4ED8")
        self.btn_start.grid(row=0, column=0, sticky="ew")
        self.btn_stop = ctk.CTkButton(btn_footer, text="⏹ STOP", font=("Segoe UI", 16, "bold"), height=50, command=self.stop_process, fg_color="#DC2626", hover_color="#B91C1C", width=150)
        self.btn_stop.grid(row=0, column=1, padx=(10, 0))
        self.btn_stop.grid_remove()
    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://www.youtube.com/watch?v=XXXXXXXXXX")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filename)
            self.manual_credentials = []
            self._refresh_manual_controls()
            self.log_message(f"✅ File loaded: {filename.split('/')[-1]}", "info")

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
            self.file_entry.insert(0, f"Selected ID: {user_id}")

    def view_saved_user(self):
        user_id = self._get_saved_user_id()
        if not user_id:
            messagebox.showinfo("Info", "No saved ID found.")
            return
        messagebox.showinfo("Saved User ID", f"Current ID: {user_id}")

    def delete_saved_user(self):
        user_id = self._get_saved_user_id()
        if not user_id:
            messagebox.showinfo("Info", "No saved ID found.")
            return
        if not messagebox.askyesno("Delete ID", f"Delete saved ID {user_id}?"):
            return
        self.manual_credentials = []
        self.file_entry.delete(0, "end")
        self._refresh_manual_controls()
        messagebox.showinfo("Deleted", "Saved ID deleted successfully.")

    def add_id_password(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add ID Password")
        dialog.geometry("420x240")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        card = ctk.CTkFrame(dialog, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(card, text="GST ID/Username").pack(anchor="w")
        ent_user = ctk.CTkEntry(card, placeholder_text="Enter GST ID/Username")
        ent_user.pack(fill="x", pady=(4, 10))

        ctk.CTkLabel(card, text="GST Password").pack(anchor="w")
        ent_pass = ctk.CTkEntry(card, placeholder_text="Enter GST Password", show="*")
        ent_pass.pack(fill="x", pady=(4, 14))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x")

        def _save():
            username = (ent_user.get() or "").strip()
            password = (ent_pass.get() or "").strip()
            if not username or not password:
                messagebox.showerror("Missing Data", "Please enter both GST ID and Password", parent=dialog)
                return

            existing_user = self._get_saved_user_id()
            if existing_user and not messagebox.askyesno(
                "Overwrite ID",
                "Your previous ID will be overwritten with this.",
                parent=dialog
            ):
                return

            self.manual_credentials = [{"Username": username, "Password": password}]
            self._refresh_manual_controls()
            messagebox.showinfo("Added", f"Credential saved for {username}", parent=dialog)
            dialog.destroy()

        ctk.CTkButton(btn_row, text="Cancel", width=110, command=dialog.destroy).pack(side="right")
        ctk.CTkButton(btn_row, text="Add", width=110, command=_save).pack(side="right", padx=(0, 8))

        ent_user.focus_set()
        dialog.bind("<Return>", lambda _e: _save())

    def log_message(self, message, tag="normal"):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def start_process(self):
        file_path = self.file_entry.get()
        manual_gstins = [
            row.get("Username", "").strip()
            for row in self.manual_credentials
            if row.get("Username", "").strip()
        ]

        if not file_path and not manual_gstins:
            messagebox.showwarning("Warning", "Please add ID/Password first!")
            return
        
        self.frame_captcha.grid()
        self.reset_status_label()
        self.cap_stop_btn.configure(state="normal", text="⏹ STOP PROCESS")
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.grid()
        self.log_message("🚀 Starting Automation Thread...", "info")
        self.worker = GSTWorker(self, file_path, manual_gstins=manual_gstins)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if not self.worker:
            return
        self.worker.stop()
        self.close_captcha_safe()
        self.btn_stop.configure(state="disabled", text="STOPPED")
        self.cap_stop_btn.configure(state="disabled", text="STOPPED")
        self.update_log("🛑 Chrome browser closed.", "warning")
        self.update_log("🛑 Process stopped by user.", "warning")

    def update_log(self, message, tag):
        self.after(0, lambda: self.log_message(message, tag))

    def update_progress(self, val):
        self.after(0, lambda: self.progress_bar.set(val))

    def close_captcha_safe(self):
        def hide_ui():
            self.frame_captcha.grid_remove()
            self.lbl_image.configure(image=None, text="[Waiting for Process...]")
            self.lbl_image.image = None
            self.entry_captcha.configure(state="disabled")
            self.entry_captcha.delete(0, "end")
            self.btn_submit.configure(state="disabled", fg_color="gray", text="SUBMIT CAPTCHA")
            self.frame_captcha.configure(border_color="#444")
            self.lbl_captcha_title.configure(text="CAPTCHA ACTION REQUIRED", text_color="gray")
        self.after(0, hide_ui)

    def show_invalid_gst_alert(self):
        def update():
            self.frame_captcha.configure(border_color="#DC2626") 
            self.lbl_captcha_title.configure(text="🚫 INVALID GSTIN DETECTED - SKIPPING", text_color="#DC2626")
            self.lbl_image.configure(image=None, text="INVALID")
            self.entry_captcha.configure(state="disabled")
            self.btn_submit.configure(state="disabled")
        self.after(0, update)

    def reset_status_label(self):
        def update():
            self.frame_captcha.configure(border_color="#444")
            self.lbl_captcha_title.configure(text="CAPTCHA ACTION REQUIRED", text_color="gray")
        self.after(0, update)

    def show_captcha_popup(self, image_data, gstin):
        def update_ui():
            if not self.worker or not self.worker.keep_running:
                return
            self.frame_captcha.grid()
            self.frame_captcha.configure(border_color="#10B981") 
            self.lbl_captcha_title.configure(text=f"ENTER CAPTCHA FOR: {gstin}", text_color="#10B981")
            
            image = Image.open(BytesIO(image_data))
            image = image.resize((200, 80), Image.Resampling.LANCZOS)
            photo = ctk.CTkImage(light_image=image, dark_image=image, size=(200, 80))
            
            self.lbl_image.configure(image=photo, text="")
            self.lbl_image.image = photo 
            
            self.entry_captcha.configure(state="normal", placeholder_text="Type here...")
            self.entry_captcha.delete(0, "end")

            self.btn_submit.configure(state="normal", fg_color="#10B981", text="SUBMIT NOW", hover_color="#047857")

            self.attributes('-topmost', True)
            self.deiconify()
            self.lift()
            def _focus():
                self.focus_force()
                self.entry_captcha.focus_set()
                self.after(1000, lambda: self.attributes('-topmost', False))
            self.after(200, _focus)

        self.after(0, update_ui)

    def submit_captcha(self):
        text = self.entry_captcha.get()
        if not text: return
        
        if self.worker and self.worker.is_waiting_for_captcha:
            self.worker.receive_captcha_input(text)
            self.entry_captcha.configure(state="disabled", placeholder_text="Verifying...")
            self.btn_submit.configure(state="disabled", fg_color="gray", text="VERIFYING...")
            self.frame_captcha.configure(border_color="#444")
            self.lbl_captcha_title.configure(text_color="gray")

    def process_finished(self, msg):
        def finish():
            is_stopped = "stopped" in (msg or "").lower()
            self.log_message(f"\n🎉 DONE: {msg}", "warning" if is_stopped else "success")
            self.close_captcha_safe()
            self.btn_start.configure(state="normal", text="STOPPED" if is_stopped else "START AUTOMATION")
            self.btn_stop.grid_remove()
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.cap_stop_btn.configure(state="normal", text="⏹ STOP PROCESS")
            self.lbl_captcha_title.configure(text="PROCESS STOPPED" if is_stopped else "PROCESS COMPLETED", text_color="gray")
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