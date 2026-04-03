import threading
import time
import os
import glob
import pandas as pd
import customtkinter as ctk
from PIL import Image
from datetime import datetime
from tkinter import filedialog, messagebox

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class GSTWorker:
    def __init__(self, app_instance, excel_path, settings):
        self.app = app_instance
        self.excel_path = excel_path
        self.settings = settings
        self.keep_running = True
        self.driver = None
        self.captcha_response = None 
        self.captcha_event = threading.Event()
        self.report_data = [] 

    def log(self, message):
        self.app.update_log_safe(message)

    def run(self):
        self.log("🚀 INITIALIZING GST ENGINE V17 (Auto-Retry + Healing)...")
        
        try:
            # 1. READ EXCEL
            df = pd.read_excel(self.excel_path)
            clean_cols = {c.lower().strip(): c for c in df.columns}
            user_col = next((clean_cols[c] for c in clean_cols if 'user' in c or 'name' in c), None)
            pass_col = next((clean_cols[c] for c in clean_cols if 'pass' in c or 'pwd' in c), None)

            if not user_col or not pass_col:
                self.app.process_finished_safe("Column Error: Need Username/Password columns")
                return

            total = len(df)
            self.log(f"📊 Loaded {total} users.")

            # 2. CREATE MAIN DOWNLOAD FOLDER
            base_dir = os.path.join(os.getcwd(), "GST_3B_Downloads")
            if not os.path.exists(base_dir): os.makedirs(base_dir)

            # 3. PROCESS LOOP
            for index, row in df.iterrows():
                if not self.keep_running: break

                username = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                
                self.app.update_progress_safe((index) / total)
                self.log(f"\n🔹 Processing: {username}")
                
                # Unique Folder Versioning
                user_root_base = os.path.join(base_dir, username)
                user_root = user_root_base
                counter = 1
                while os.path.exists(user_root):
                    user_root = f"{user_root_base}_{counter}"
                    counter += 1
                os.makedirs(user_root)

                status, reason = self.process_single_user(username, password, user_root)
                
                self.report_data.append({
                    "Username": username,
                    "Status": status,
                    "Details": reason,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Saved To": os.path.basename(user_root)
                })
                
                self.log("-" * 40)

            self.generate_excel_report()
            self.app.update_progress_safe(1.0)
            self.log("✅ ALL TASKS COMPLETED.")
            self.app.process_finished_safe("Batch Completed & Report Saved.")

        except Exception as e:
            self.log(f"❌ Critical Error: {e}")
            self.app.process_finished_safe("Error Occurred")

    def generate_excel_report(self):
        try:
            if not self.report_data: return
            report_df = pd.DataFrame(self.report_data)
            filename = f"GST_3B_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_df.to_excel(filename, index=False)
            self.log(f"📄 Summary Report saved: {filename}")
        except Exception as e:
            self.log(f"⚠️ Failed to save report: {e}")

    def handle_popups(self):
        """ Checks for Aadhaar/Generic Popups and closes them. """
        time.sleep(1.5) 
        try:
            aadhaar_skip = self.driver.find_elements(By.XPATH, "//a[contains(text(),'Remind me later')]")
            if aadhaar_skip and aadhaar_skip[0].is_displayed():
                self.log("      🛡️ Closing Aadhaar Popup...")
                aadhaar_skip[0].click()
                time.sleep(1)
        except: pass

        try:
            generic_skip = self.driver.find_elements(By.XPATH, "//button[contains(text(),'Remind Me Later')]")
            if generic_skip and generic_skip[0].is_displayed():
                self.log("      🛡️ Closing Generic Popup...")
                generic_skip[0].click()
                time.sleep(1)
        except: pass

    def ensure_return_dashboard(self, wait):
        """ 
        Checks if we are on the Return Dashboard.
        If not (e.g. kicked to Home), it handles popups and navigates back.
        """
        try:
            # Check if Financial Year dropdown is present
            if self.driver.find_elements(By.NAME, "fin"):
                return True # We are safe
            
            self.log("      🔄 Lost Dashboard. Recovering...")
            
            # 1. Handle Popups (in case we are on Home Page)
            self.handle_popups()

            # 2. Click "Return Dashboard" button
            try:
                dash_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Return Dashboard')]")))
                dash_btn.click()
                time.sleep(3)
                return True
            except:
                # Try JS Click fallback
                try:
                    btn = self.driver.find_element(By.XPATH, "//button[contains(., 'Return Dashboard')]")
                    self.driver.execute_script("arguments[0].click();", btn)
                    time.sleep(3)
                    return True
                except:
                    self.log("      ⚠️ Recovery Failed: Button missing.")
                    return False
        except Exception as e:
            self.log(f"      ⚠️ Recovery Exception: {e}")
            return False

    def process_single_user(self, username, password, user_root):
        try:
            # --- BROWSER SETUP ---
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled") 
            options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
            options.add_experimental_option('useAutomationExtension', False)
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")

            prefs = {
                "download.prompt_for_download": False,
                "directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1
            }
            options.add_experimental_option("prefs", prefs)
            
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            })

            wait = WebDriverWait(self.driver, 20)
            
            # 1. LOGIN
            login_status, login_msg = self.perform_login(username, password, wait)
            if not login_status: return "Login Failed", login_msg

            # 2. DEFINE TASKS
            q_map = {
                "Quarter 1 (Apr - Jun)": ["April", "May", "June"],
                "Quarter 2 (Jul - Sep)": ["July", "August", "September"],
                "Quarter 3 (Oct - Dec)": ["October", "November", "December"],
                "Quarter 4 (Jan - Mar)": ["January", "February", "March"]
            }

            tasks = []
            if self.settings['all_quarters']:
                for q_name, months in q_map.items():
                    for m in months:
                        tasks.append({"q": q_name, "m": m})
            else:
                selected_q = self.settings['quarter']
                if selected_q in q_map:
                    for m in q_map[selected_q]:
                        tasks.append({"q": selected_q, "m": m})
                else:
                    return "Config Error", "Invalid Quarter Selected"

            self.log(f"   📅 Queued {len(tasks)} Months...")
            
            # 3. EXECUTE LOOP
            time.sleep(2) 
            success_count = 0
            results = []

            # Create Year Folder
            fin_year = self.settings['year']
            year_folder = os.path.join(user_root, fin_year)
            if not os.path.exists(year_folder): os.makedirs(year_folder)

            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow", 
                "downloadPath": year_folder
            })

            for task in tasks:
                if not self.keep_running: return "Stopped", "User Cancelled"
                
                q_text = task['q']
                m_text = task['m']
                
                # --- RETRY MECHANISM ---
                max_retries = 3
                month_success = False
                
                for attempt in range(1, max_retries + 1):
                    # Show attempt number only if retrying
                    attempt_str = f" [Attempt {attempt}]" if attempt > 1 else ""
                    self.log(f"   ⚙️ Processing: {m_text} ({q_text[:9]}){attempt_str}")

                    try:
                        # 1. Dashboard Check
                        if not self.ensure_return_dashboard(wait):
                            self.log(f"      ⚠️ Dashboard check failed. Retrying...")
                            time.sleep(2)
                            continue # Try next attempt

                        # 2. Select Year
                        year_el = wait.until(EC.element_to_be_clickable((By.NAME, "fin")))
                        Select(year_el).select_by_visible_text(fin_year)
                        time.sleep(0.5)

                        # 3. Select Quarter
                        qtr_el = self.driver.find_element(By.NAME, "quarter")
                        Select(qtr_el).select_by_visible_text(q_text)
                        time.sleep(0.5)

                        # 4. Select Month
                        mon_el = self.driver.find_element(By.NAME, "mon")
                        Select(mon_el).select_by_visible_text(m_text)
                        time.sleep(0.5)

                        # 5. Click Search
                        self.driver.find_element(By.XPATH, "//button[contains(text(), 'Search')]").click()
                        time.sleep(4) 

                        # 6. Download
                        dl_status, dl_msg = self.download_gstr3b(wait, year_folder)
                        
                        if dl_status:
                            success_count += 1
                            results.append(f"{m_text}: ✅")
                            month_success = True
                            break # Success! Break retry loop
                        else:
                            self.log(f"      ⚠️ Failed ({dl_msg}). Retrying...")
                            
                            # Navigation Reset for next retry
                            if attempt < max_retries:
                                try: self.driver.back()
                                except: pass
                                time.sleep(2)
                                
                    except Exception as e:
                        self.log(f"      ❌ Error: {str(e)[:30]}")
                        
                        # Aggressive Recovery for next attempt
                        if attempt < max_retries:
                            self.log("      🔄 Refreshing page to fix glitch...")
                            try: 
                                self.driver.refresh()
                                time.sleep(4)
                            except: pass
                
                # If all retries failed
                if not month_success:
                    results.append(f"{m_text}: ❌ Failed")

            # 4. FINAL STATUS
            overall_status = "Success" if success_count == len(tasks) else "Partial"
            if success_count == 0: overall_status = "Failed"
            
            summary = f"Downloaded {success_count}/{len(tasks)}. Details: " + ", ".join(results)
            return overall_status, summary

        except Exception as e:
            return "Error", f"Browser Crash: {str(e)[:30]}"
        finally:
            if self.driver:
                self.driver.quit()

    def perform_login(self, username, password, wait):
        self.log("   🌐 Opening GST Portal...")
        self.driver.get("https://services.gst.gov.in/services/login")

        while True:
            if not self.keep_running: return False, "Stopped"

            try:
                wait.until(EC.visibility_of_element_located((By.ID, "username"))).clear()
                self.driver.find_element(By.ID, "username").send_keys(username)
                
                self.driver.find_element(By.ID, "user_pass").clear()
                self.driver.find_element(By.ID, "user_pass").send_keys(password)

                captcha_img = wait.until(EC.visibility_of_element_located((By.ID, "imgCaptcha")))
                captcha_img.screenshot("temp_captcha.png")
                
                self.log("   ⌨️ Waiting for Captcha...")
                self.captcha_response = None
                self.captcha_event.clear()
                self.app.request_captcha_safe("temp_captcha.png")
                self.captcha_event.wait() 

                if not self.captcha_response: return False, "Captcha Cancelled"

                self.driver.find_element(By.ID, "captcha").clear()
                self.driver.find_element(By.ID, "captcha").send_keys(self.captcha_response)
                self.driver.find_element(By.XPATH, "//button[@type='submit']").click()
                
                time.sleep(3)

                src = self.driver.page_source
                if "Invalid Username or Password" in src:
                    self.log("   ❌ Bad Credentials.")
                    return False, "Invalid Credentials"
                
                if "Enter valid Letters" in src or "Invalid Captcha" in src:
                    self.log("   ⚠️ Invalid Captcha. Retrying...")
                    time.sleep(1)
                    continue 

                if "Dashboard" in self.driver.title or "Return Dashboard" in src:
                    self.log("   ✅ Login Successful!")
                    self.app.close_captcha_safe()
                    
                    self.handle_popups()
                    
                    # Navigate to Return Dashboard
                    try:
                        dash_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Return Dashboard')]")))
                        dash_btn.click()
                        return True, "Success"
                    except:
                        try:
                            btn = self.driver.find_element(By.XPATH, "//button[contains(., 'Return Dashboard')]")
                            self.driver.execute_script("arguments[0].click();", btn)
                            return True, "Success (JS Click)"
                        except:
                            self.log("   ⚠️ Dashboard Nav Error.")
                            return False, "Dashboard Nav Failed"

            except Exception as e:
                self.log(f"   ⚠️ Login Exception: {e}")
                return False, f"Login Error: {str(e)[:20]}"

    def download_gstr3b(self, wait, download_path):
        self.log("   🔍 Searching for GSTR-3B Tile...")

        xpath_3b_btn = "//p[contains(text(),'GSTR-3B')]/ancestor::div[contains(@class,'col-sm-4')]//button[contains(text(),'Download')]"
        
        try:
            found_btn = self.driver.find_element(By.XPATH, xpath_3b_btn)
            
            if not found_btn.is_displayed():
                self.log("   ⚠️ GSTR-3B Tile hidden (Not Filed/Available).")
                self.driver.back()
                return False, "Not Filed"
                
            self.log("   ✅ Found Filed GSTR-3B.")
            
        except:
            self.log("   ⚠️ GSTR-3B Tile Not Found.")
            return False, "Not Found"

        try:
            self.log("   ⬇️ Clicking Download Button...")
            self.driver.execute_script("arguments[0].click();", found_btn)
            time.sleep(3) 
            
            self.log("   ⏳ Waiting for file...")
            file_downloaded = False
            
            for _ in range(15):
                time.sleep(1)
                files = glob.glob(os.path.join(download_path, "*.*"))
                if files:
                    latest = max(files, key=os.path.getctime)
                    if (datetime.now().timestamp() - os.path.getctime(latest)) < 30:
                        self.log(f"   ✅ Saved: {os.path.basename(latest)}")
                        file_downloaded = True
                        break
            
            self.driver.back() 
            
            if not file_downloaded:
                self.log("   ⚠️ File download timed out.")
                return False, "Timeout"
            
            return True, "Success"
                    
        except Exception as e:
            self.log(f"   ⚠️ Download Script Error: {str(e)[:20]}")
            self.driver.back()
            return False, "Script Error"


# --- GUI CLASS ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GST Bulk Downloader - Professional Edition")
        self.geometry("900x850")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) 

        self.worker = None
        self.excel_file = ""

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1a237e", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False) 
        ctk.CTkLabel(self.head, text="GST BULK DOWNLOADER", 
                      font=("Roboto Medium", 24, "bold"), text_color="white").pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(self.head, text="GSTR-3B AUTOMATION", 
                      font=("Roboto", 14), text_color="#bbdefb").pack(side="right", padx=20, pady=15)

        # SETTINGS
        self.settings_container = ctk.CTkFrame(self, fg_color="transparent")
        self.settings_container.grid(row=1, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.settings_container.grid_columnconfigure((0, 1), weight=1)

        # Credentials Card
        self.card_cred = ctk.CTkFrame(self.settings_container, border_color="#3949ab", border_width=1)
        self.card_cred.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ctk.CTkLabel(self.card_cred, text="📂 Credentials Source", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))
        self.ent_file = ctk.CTkEntry(self.card_cred, placeholder_text="Select Excel File...", height=35)
        self.ent_file.pack(fill="x", padx=15, pady=(5, 10))
        self.btn_browse = ctk.CTkButton(self.card_cred, text="Browse File", command=self.browse_file, 
                                        fg_color="#3949ab", hover_color="#283593", height=35)
        self.btn_browse.pack(fill="x", padx=15, pady=(0, 15))

        # Period Settings Card
        self.card_period = ctk.CTkFrame(self.settings_container, border_color="#3949ab", border_width=1)
        self.card_period.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        ctk.CTkLabel(self.card_period, text="📅 Period Selection", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))

        # Dynamic Year - Generate in ascending order (oldest first)
        cur_year = datetime.now().year
        start_year = cur_year - 1 if datetime.now().month < 4 else cur_year
        year_list = [f"{y}-{str(y+1)[-2:]}" for y in range(start_year - 2, start_year + 2)]

        self.frm_year = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_year.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_year, text="Financial Year:", width=100, anchor="w").pack(side="left")
        self.cb_year = ctk.CTkComboBox(self.frm_year, values=year_list, width=150)
        self.cb_year.set(year_list[0]) 
        self.cb_year.pack(side="right", expand=True, fill="x")
        
        # Checkbox
        self.chk_all_qtr_var = ctk.BooleanVar(value=False)
        self.chk_all_qtr = ctk.CTkCheckBox(self.card_period, text="Download All Quarters (Apr-Mar)", 
                                            variable=self.chk_all_qtr_var, command=self.toggle_inputs,
                                            font=("Arial", 12, "bold"))
        self.chk_all_qtr.pack(anchor="w", padx=15, pady=5)

        # Quarter & Month
        self.frm_qtr = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_qtr.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_qtr, text="Quarter:", width=100, anchor="w").pack(side="left")
        self.cb_qtr = ctk.CTkComboBox(self.frm_qtr, 
                                      values=["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)", 
                                              "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"],
                                      command=self.update_months_based_on_qtr, width=150)
        self.cb_qtr.set("Quarter 1 (Apr - Jun)")
        self.cb_qtr.pack(side="right", expand=True, fill="x")

        # Month
        self.frm_mon = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_mon.pack(fill="x", padx=15, pady=(2, 15))
        ctk.CTkLabel(self.frm_mon, text="Month:", width=100, anchor="w").pack(side="left")
        self.cb_month = ctk.CTkComboBox(self.frm_mon, values=["April", "May", "June"], width=150)
        self.cb_month.set("April")
        self.cb_month.pack(side="right", expand=True, fill="x")

        # LOGS
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=10)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self.log_frame, text="📜 Execution Logs", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), text_color="#00e676", height=150)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.log_box.configure(state="disabled")

        # CAPTCHA (Hidden)
        self.cap_frame = ctk.CTkFrame(self, border_color="#d32f2f", border_width=2, fg_color="#2b2b2b")
        self.cap_inner = ctk.CTkFrame(self.cap_frame, fg_color="transparent")
        self.cap_inner.pack(fill="both", padx=20, pady=10)
        ctk.CTkLabel(self.cap_inner, text="⚠️ CAPTCHA ACTION REQUIRED", text_color="#ef5350", font=("Arial", 14, "bold")).pack()
        self.cap_lbl_img = ctk.CTkLabel(self.cap_inner, text="")
        self.cap_lbl_img.pack(pady=10)
        self.cap_ent = ctk.CTkEntry(self.cap_inner, placeholder_text="Type Captcha Here...", 
                                    font=("Consolas", 20), justify="center", height=45, width=250)
        self.cap_ent.pack(pady=5)
        self.cap_ent.bind("<Return>", self.submit_captcha) 
        self.cap_btn = ctk.CTkButton(self.cap_inner, text="SUBMIT CAPTCHA", fg_color="#d32f2f", hover_color="#b71c1c", 
                                     height=40, width=250, font=("Arial", 12, "bold"), command=self.submit_captcha)
        self.cap_btn.pack(pady=(10, 0))

        # FOOTER
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 20))
        self.prog_bar = ctk.CTkProgressBar(self.footer, height=15, progress_color="#00e676")
        self.prog_bar.pack(fill="x", pady=(0, 10))
        self.prog_bar.set(0)
        self.btn_start = ctk.CTkButton(self.footer, text="START BATCH PROCESS", height=50, font=("Arial", 16, "bold"), 
                                       fg_color="#2e7d32", hover_color="#1b5e20", command=self.start_process)
        self.btn_start.pack(fill="x")

    def toggle_inputs(self):
        state = "disabled" if self.chk_all_qtr_var.get() else "normal"
        self.cb_qtr.configure(state=state)
        self.cb_month.configure(state=state)

    def update_months_based_on_qtr(self, choice):
        if "Quarter 1" in choice: vals = ["April", "May", "June"]
        elif "Quarter 2" in choice: vals = ["July", "August", "September"]
        elif "Quarter 3" in choice: vals = ["October", "November", "December"]
        elif "Quarter 4" in choice: vals = ["January", "February", "March"]
        else: vals = ["April"]
        self.cb_month.configure(values=vals)
        self.cb_month.set(vals[0])

    def browse_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            self.excel_file = f
            self.ent_file.delete(0, "end")
            self.ent_file.insert(0, f)

    def log_gui(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def update_log_safe(self, msg):
        self.after(0, lambda: self.log_gui(msg))

    def update_progress_safe(self, val):
        self.after(0, lambda: self.prog_bar.set(val))

    def process_finished_safe(self, msg):
        self.after(0, lambda: messagebox.showinfo("Info", msg))
        self.after(0, lambda: self.btn_start.configure(state="normal", text="START BATCH PROCESS"))

    def request_captcha_safe(self, img_path):
        def show():
            pil_img = Image.open(img_path)
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(160, 60))
            self.cap_lbl_img.configure(image=ctk_img)
            self.cap_btn.configure(state="normal", text="SUBMIT CAPTCHA", fg_color="#d32f2f")
            self.cap_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=10)
            self.cap_ent.delete(0, "end")
            self.cap_ent.focus_set()
            self.lift()
            self.attributes('-topmost',True)
            self.after_idle(self.attributes,'-topmost',False)
        self.after(0, show)

    def submit_captcha(self, event=None):
        txt = self.cap_ent.get()
        if not txt: return
        self.cap_btn.configure(state="disabled", text="VERIFYING...", fg_color="gray")
        self.worker.captcha_response = txt
        self.worker.captcha_event.set()

    def close_captcha_safe(self):
        self.after(0, lambda: self.cap_frame.grid_forget())

    def start_process(self):
        if not self.excel_file:
            messagebox.showerror("Error", "Please select Excel file")
            return
        settings = {
            "year": self.cb_year.get(),
            "month": self.cb_month.get(),
            "quarter": self.cb_qtr.get(),
            "all_quarters": self.chk_all_qtr_var.get()
        }
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.worker = GSTWorker(self, self.excel_file, settings)
        threading.Thread(target=self.worker.run, daemon=True).start()

if __name__ == "__main__":
    app = App()
    app.mainloop()