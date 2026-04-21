import threading
import time
import os
import re
import tempfile
import pandas as pd
import customtkinter as ctk
from datetime import datetime
from tkinter import filedialog, messagebox

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

YEAR_MODE_OPTIONS = [
    "Current Year",
    "Current and Last Year",
    "Current and Last 2 Years",
]

def get_taxpayer_name(driver, fallback=""):
    """
    Extracts the logged-in taxpayer's name from the Income Tax portal header.
    Uses JavaScript (innerText) as primary method for Angular reliability.
    """
    strategies = [
        lambda d: d.execute_script(
            "var el = document.querySelector('.userNameVal span:first-child'); "
            "return el ? el.innerText.trim() : '';"
        ),
        lambda d: d.execute_script(
            "var el = document.querySelector('.profileMenubtn .userNameVal span'); "
            "return el ? el.innerText.trim() : '';"
        ),
        lambda d: next(
            (s for s in (
                d.execute_script(
                    "var spans = document.querySelectorAll('.userNameVal span'); "
                    "return Array.from(spans).map(function(s){return s.innerText.trim();});"
                ) or []
            ) if s and s.lower() not in ['expand_more', '']),
            ''
        ),
        lambda d: d.find_element(By.XPATH, "//span[contains(@class,'userNameVal')]/span[1]").text.strip(),
    ]
    for strategy in strategies:
        try:
            result = strategy(driver)
            if result and len(result) > 1 and result.lower() not in ['expand_more']:
                result = result.split('\n')[0].strip()
                if result:
                    return result
        except:
            continue
    return fallback

class IncomeTaxWorker:
    """
    Handles the background automation logic (Selenium).
    Runs in a separate thread to keep the UI responsive.
    """
    def __init__(self, app_instance, excel_path, years_pref):
        self.app = app_instance
        self.excel_path = excel_path
        self.years_pref = years_pref
        self.keep_running = True
        self.report_data = [] # To store summary report

    def log(self, message):
        self.app.update_log_safe(message)

    def normalize_columns(self, df):
        """
        Intelligently find the User ID/PAN and Password columns 
        regardless of case or specific naming (e.g. 'pan number', 'userid', 'pass').
        """
        user_col = None
        pass_col = None
        
        # Normalize headers to lowercase and strip spaces
        clean_cols = {c: c.lower().strip().replace(" ", "").replace("_", "") for c in df.columns}
        
        # Search patterns
        pan_patterns = ['userid', 'user', 'pan', 'pannumber', 'panid', 'loginid']
        pass_patterns = ['password', 'pass', 'pwd', 'loginpass']

        for original, clean in clean_cols.items():
            if not user_col and any(p in clean for p in pan_patterns):
                user_col = original
            if not pass_col and any(p in clean for p in pass_patterns):
                pass_col = original
        
        return user_col, pass_col

    def run(self):
        self.log("🚀 INITIALIZING ENGINE...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        
        try:
            # 1. READ EXCEL
            df = pd.read_excel(self.excel_path)
            
            # Identify columns
            user_col, pass_col = self.normalize_columns(df)
            
            if not user_col or not pass_col:
                self.log(f"❌ ERROR: Could not identify 'User ID' or 'Password' columns.")
                self.log(f"   Found columns: {list(df.columns)}")
                self.app.process_finished_safe("Failed: Column Header Error")
                return

            self.log(f"✅ Mapped Columns -> ID: '{user_col}', Pass: '{pass_col}'")
            total_users = len(df)
            self.log(f"📊 Found {total_users} users in queue.\n")

            # 2. LOOP THROUGH USERS
            for index, row in df.iterrows():
                if not self.keep_running: 
                    self.log("🛑 Process Stopped by User.")
                    break
                
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                
                # Update Progress Bar
                progress_val = (index) / total_users
                self.app.update_progress_safe(progress_val)

                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                # --- FOLDER ROOT (actual folder created after login with name) ---
                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "Income Tax Downloaded", "ITR Bot")
                if not os.path.exists(download_root): os.makedirs(download_root, exist_ok=True)

                # START BROWSER
                status, reason, final_path = self.process_single_user(user_id, password, download_root)
                
                # Add to Report
                self.report_data.append({
                    "PAN / User ID": user_id,
                    "Status": status,
                    "Reason/Details": reason,
                    "Folder": os.path.basename(final_path) if final_path else user_id,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                
                self.log("-" * 40)
            
            # --- GENERATE SUMMARY REPORT ---
            self.generate_report()
                
            self.app.update_progress_safe(1.0)
            self.log("\n✅ BATCH PROCESSING COMPLETED!")
            self.app.process_finished_safe("All Tasks Completed. Report Generated.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe("Critical Error Occurred")

    def generate_report(self):
        try:
            if not self.report_data: return
            
            df_report = pd.DataFrame(self.report_data)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Processing_Report_{timestamp}.xlsx"
            report_dir = os.path.join(os.getcwd(), "Income Tax Downloaded", "ITR Bot")
            if not os.path.exists(report_dir): os.makedirs(report_dir, exist_ok=True)
            report_path = os.path.join(report_dir, filename)
            df_report.to_excel(report_path, index=False)
            self.log(f"📄 Summary Report saved as: {filename}")
        except Exception as e:
            self.log(f"⚠️ Failed to save report: {e}")

    def process_single_user(self, user_id, password, download_root):
        """
        Returns tuple: (Status, Reason, FolderPath)
        """
        driver = None
        # Temp PAN-only folder (Chrome needs a path at startup)
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', user_id).strip()
        main_download_folder = os.path.join(download_root, safe_name)
        if not os.path.exists(main_download_folder):
            os.makedirs(main_download_folder)
        try:
            # --- BROWSER CONFIG ---
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            options.add_argument("--disable-blink-features=AutomationControlled")

            prefs = {
                "download.default_directory": main_download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_setting_values.automatic_downloads": 1
            }
            options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            wait = WebDriverWait(driver, 15)
            actions = ActionChains(driver)

            # ============================================================
            # LOGIN RETRY LOOP
            # ============================================================
            login_success = False
            fail_reason = "Unknown Error"
            
            for login_attempt in range(1, 4): # Try up to 3 times
                if login_success: break
                
                if login_attempt > 1:
                    self.log(f"   ⚠️ Login Retry {login_attempt}/3...")
                    driver.delete_all_cookies()
                    driver.refresh()
                    time.sleep(3)

                try:
                    # STEP 1: LOGIN PAGE
                    self.log("   🌐 Opening Portal...")
                    driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")

                    try:
                        time.sleep(1)
                        driver.switch_to.alert.accept()
                    except: pass

                    # User ID
                    self.log("   🔑 Entering Credentials...")
                    pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                    pan_field.clear()
                    pan_field.send_keys(user_id)
                    
                    btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                    btn.click()

                    # Invalid PAN Check
                    time.sleep(1.5)
                    if "does not exist" in driver.page_source:
                        self.log("   ❌ Invalid PAN ID. Skipping.")
                        return "Failed", "Invalid PAN Number"

                    # Password
                    pass_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                    pass_field.clear()
                    pass_field.send_keys(password)
                    
                    try:
                        cb = driver.find_element(By.ID, "passwordCheckBox-input")
                        driver.execute_script("arguments[0].click();", cb)
                    except: pass
                    
                    # Human Pause
                    time.sleep(4) 
                    
                    login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                    driver.execute_script("arguments[0].click();", login_btn)

                    # Validation Loop (30s)
                    self.log("   ⏳ Verifying Session...")
                    for _ in range(30):
                        time.sleep(1)
                        
                        if driver.find_elements(By.ID, "e-File"):
                            self.log("   ✅ Login Successful!")
                            login_success = True
                            break
                        
                        if "Invalid Password" in driver.page_source:
                            self.log("   ❌ Wrong Password. Skipping.")
                            return "Failed", "Invalid Password"
                        
                        # Dual Login Fix
                        try:
                            dual_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual_btn and dual_btn[0].is_displayed():
                                self.log("   ⚠️ Dual Session Detected. Overriding...")
                                driver.execute_script("arguments[0].click();", dual_btn[0])
                                time.sleep(3)
                        except: pass

                        # Auth Retry
                        try:
                            if "Request is not authenticated" in driver.page_source:
                                 login_btns = driver.find_elements(By.CSS_SELECTOR, "button.large-button-primary")
                                 if login_btns:
                                     driver.execute_script("arguments[0].click();", login_btns[0])
                        except: pass

                    if login_success: break 

                except Exception as e:
                    fail_reason = f"Login Error: {str(e)[:50]}"
            
            if not login_success:
                self.log("   ❌ Login Timed Out. Skipping User.")
                return "Failed", "Login Timeout / Server Issue", main_download_folder

            # Extract taxpayer name and create NAME_PAN folder
            name_from_header = get_taxpayer_name(driver, fallback=user_id)
            if name_from_header != user_id:
                self.log(f"   👤 Taxpayer Name: {name_from_header}")
            else:
                self.log("   ⚠️ Name not found in header; using PAN as folder name.")

            folder_name = re.sub(r'[<>:"/\\|?*]', '_', f"{name_from_header}_{user_id}").strip()
            main_download_folder = os.path.join(download_root, folder_name)
            if not os.path.exists(main_download_folder):
                os.makedirs(main_download_folder)
            self.log(f"   📁 Download folder: {folder_name}")

            # STEP 3: NAVIGATION
            self.log("   🚀 Navigating to Returns...")
            nav_success = False
            for nav_attempt in range(3):
                try:
                    time.sleep(2)
                    e_file = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "e-File")))
                    driver.execute_script("arguments[0].click();", e_file)
                    
                    submenu = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Income Tax Returns')]")))
                    actions.move_to_element(submenu).perform()
                    time.sleep(1)
                    
                    view_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'View Filed Returns')]")))
                    driver.execute_script("arguments[0].click();", view_btn)
                    
                    nav_success = True
                    break 
                except:
                    self.log(f"      ⚠️ Navigation Retry {nav_attempt+1}...")
                    driver.refresh()
                    time.sleep(3)
            
            if not nav_success:
                self.log("   ❌ Navigation Failed after retries. Skipping.")
                return "Failed", "Navigation Failed (Menu Issue)"

            # ============================================================
            # STEP 4: DOWNLOAD
            # ============================================================
            self.log("   ⬇️  Scanning Files...")
            try:
                # Wait up to 10 seconds for return cards to appear
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "contextBox")))
                time.sleep(1)
                cards = driver.find_elements(By.CLASS_NAME, "contextBox")
                
                total_available = len(cards)
                
                # --- CALCULATE TARGET COUNT ---
                selected_mode = (self.years_pref or "").strip()
                if selected_mode in ("Current and Last 2 Years", "Last 3 Years"):
                    target_count = 3
                elif selected_mode in ("Current and Last Year", "Last 2 Years"):
                    target_count = 2
                elif selected_mode in ("Current Year", "Last 1 Year"):
                    target_count = 1
                elif "All" in selected_mode:
                    target_count = total_available
                else:
                    target_count = 1
                
                final_count = min(total_available, target_count)
                
                self.log(f"   ℹ️  Processing {final_count} years...")

                for i in range(final_count):
                    cards = driver.find_elements(By.CLASS_NAME, "contextBox") # Refresh DOM
                    card = cards[i]
                    
                    # 1. Extract Year Name
                    try:
                        year_text = card.find_element(By.CLASS_NAME, "contentHeadingText").text
                        safe_year = year_text.replace("A.Y.", "AY").replace(" ", "_").strip()
                    except: 
                        safe_year = f"Year_{i+1}"
                    
                    self.log(f"      📄 Found: {safe_year}")

                    # 2. Create Sub-Folder for Year
                    year_folder_path = os.path.join(main_download_folder, safe_year)
                    if not os.path.exists(year_folder_path):
                        os.makedirs(year_folder_path)

                    # 3. CHANGE BROWSER DOWNLOAD PATH DYNAMICALLY
                    try:
                        params = {'behavior': 'allow', 'downloadPath': year_folder_path}
                        driver.execute_cdp_cmd('Page.setDownloadBehavior', params)
                    except Exception as e:
                        self.log(f"         ⚠️ Path Set Error: {e}")

                    # 4. Trigger Downloads
                    def click_dl(cls, name):
                        try:
                            btn = card.find_element(By.CSS_SELECTOR, f".{cls}")
                            driver.execute_script("arguments[0].click();", btn)
                            self.log(f"         -> {name} Saving...")
                            time.sleep(0.5)
                        except: pass

                    click_dl("dformback", f"{name_from_header}-Form")
                    click_dl("drecback", f"{name_from_header}-Receipt")
                    click_dl("dxmlback", f"{name_from_header}-JSON")
                    
                    time.sleep(2) 
                
                self.log("   ✅ All Year Downloads Finished.")
                time.sleep(3)
                return "Success", f"Downloaded {final_count} years", main_download_folder
                
            except TimeoutException:
                self.log("   ⚠️ No filed returns found on account. Skipping.")
                return "Warning", "No Returns Filed", main_download_folder
            except Exception as e:
                self.log(f"   ❌ Download Error: {e}")
                return "Failed", f"Download Error: {str(e)[:30]}", main_download_folder

        except Exception as e:
            self.log(f"   ❌ Browser Error: {e}")
            return "Failed", f"Browser Crash: {str(e)[:30]}", main_download_folder
        finally:
            if driver:
                driver.quit()

# --- MODERN GUI APP ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("ITR Automation Suite Pro")
        self.geometry("800x700")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Variables
        self.excel_file_path = ""
        self.manual_credentials = []
        self.worker = None

        # --- HEADER SECTION ---
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        self.title_label = ctk.CTkLabel(self.header_frame, text="INCOME TAX BULK DOWNLOADER", 
                                      font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(side="left")
        
        self.status_label = ctk.CTkLabel(self.header_frame, text="v6.0 | Final Report Ed.", 
                                       text_color="gray", font=ctk.CTkFont(size=12))
        self.status_label.pack(side="left", padx=10, pady=(10, 0))

        # --- 1. CONFIGURATION CARD ---
        self.config_frame = ctk.CTkFrame(self)
        self.config_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        
        # 1.1 File Selection
        self.step1_label = ctk.CTkLabel(self.config_frame, text="1. CREDENTIALS SOURCE", 
                                      font=ctk.CTkFont(size=14, weight="bold"))
        self.step1_label.pack(anchor="w", padx=15, pady=(15, 5))
        
        self.file_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        self.file_frame.pack(fill="x", padx=15, pady=(0, 5))
        
        self.entry_file = ctk.CTkEntry(self.file_frame, placeholder_text="Add PAN/User ID and Password manually...")
        self.entry_file.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.file_actions_frame = ctk.CTkFrame(self.file_frame, fg_color="transparent")
        self.file_actions_frame.pack(side="right")
        # Add ID first
        self.btn_sample = ctk.CTkButton(self.file_actions_frame, text="➕ Add ID Password", command=self.add_id_password,
            fg_color="#059669", hover_color="#047857", width=150,
                font=("Segoe UI", 12, "bold"))
        self.btn_sample.pack(side="left")

        # View and Delete next
        self.btn_view_id = ctk.CTkButton(self.file_actions_frame, text="👁 View ID", command=self.view_saved_user,
             fg_color="#475569", hover_color="#334155", width=95,
             font=("Segoe UI", 11, "bold"))
        self.btn_view_id.pack(side="left", padx=(5, 0))

        self.btn_delete_id = ctk.CTkButton(self.file_actions_frame, text="🗑 Delete ID", command=self.delete_saved_user,
               fg_color="#7C3AED", hover_color="#6D28D9", width=105,
               font=("Segoe UI", 11, "bold"))
        self.btn_delete_id.pack(side="left", padx=(5, 0))

        # Demo last
        self.btn_demo = ctk.CTkButton(self.file_actions_frame, text="▶ Demo", command=self.open_demo_link,
                  fg_color="#DC2626", hover_color="#B91C1C", width=80,
                  font=("Segoe UI", 12, "bold"))
        self.btn_demo.pack(side="left", padx=(5, 0))

        self.btn_view_id.configure(state="disabled")
        self.btn_delete_id.configure(state="disabled")

        # 1.2 Year Selection
        self.step2_label = ctk.CTkLabel(self.config_frame, text="2. DOWNLOAD SETTINGS", 
                                      font=ctk.CTkFont(size=14, weight="bold"))
        self.step2_label.pack(anchor="w", padx=15, pady=(10, 5))

        self.pref_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        self.pref_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.lbl_years = ctk.CTkLabel(self.pref_frame, text="Select Number of Years:", text_color="gray")
        self.lbl_years.pack(side="left", padx=(0, 10))

        self.combo_years = ctk.CTkComboBox(self.pref_frame, 
                         values=YEAR_MODE_OPTIONS,
                         width=200, state="readonly")
        self.combo_years.set("Current Year") 
        self.combo_years.pack(side="left")

        # --- 2. TERMINAL LOG SECTION ---
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=10)
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.step3_label = ctk.CTkLabel(self.log_frame, text="3. LIVE EXECUTION LOG", 
                                      font=ctk.CTkFont(size=14, weight="bold"))
        self.step3_label.grid(row=0, column=0, sticky="w", padx=15, pady=(15, 5))

        # Terminal-like Textbox
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), 
                                    text_color="#10B981", fg_color="#0F172A",
                                    activate_scrollbars=True)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box.insert("0.0", "System Ready...\nWaiting for input...\n")
        self.log_box.configure(state="disabled")

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self.log_frame, mode="determinate")
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_bar.set(0)

        # --- 3. CONTROLS ---
        btn_footer = ctk.CTkFrame(self, fg_color="transparent")
        btn_footer.grid(row=3, column=0, sticky="ew", padx=20, pady=(10, 20))
        btn_footer.grid_columnconfigure(0, weight=1)
        self.btn_start = ctk.CTkButton(btn_footer, text="INITIATE BATCH DOWNLOAD",
                                       font=ctk.CTkFont(size=16, weight="bold"),
                                       height=50, fg_color="#059669", hover_color="#047857",
                                       command=self.start_process)
        self.btn_start.grid(row=0, column=0, sticky="ew")
        self.btn_stop = ctk.CTkButton(btn_footer, text="⏹ STOP", font=ctk.CTkFont(size=16, weight="bold"),
                                      height=50, fg_color="#DC2626", hover_color="#B91C1C",
                                      command=self.stop_process, width=150)
        self.btn_stop.grid(row=0, column=1, padx=(10, 0))
        self.btn_stop.grid_remove()
        self.btn_open_folder = ctk.CTkButton(btn_footer, text="📂 OPEN FOLDER", font=ctk.CTkFont(size=16, weight="bold"),
                                      height=50, fg_color="#2563EB", hover_color="#1D4ED8",
                                      command=self.open_output_folder, width=180)
        self.btn_open_folder.grid(row=0, column=2, padx=(10, 0))
        self.btn_open_folder.grid_remove()

    def download_sample(self):
        import shutil
        from tkinter import messagebox
        sample_path = os.path.join(os.path.dirname(__file__), "Income Tax Sample File.xlsx")
        if not os.path.exists(sample_path):
            messagebox.showerror("Not Found", f"Sample file not found:\n{sample_path}")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Income Tax Sample File.xlsx", filetypes=[("Excel", "*.xlsx")])
        if save_path:
            try:
                shutil.copy2(sample_path, save_path)
                messagebox.showinfo("Success", f"Sample file saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {e}")

    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://www.youtube.com/watch?v=XXXXXXXXXX")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.excel_file_path = filename
            self.manual_credentials = []
            self._refresh_manual_controls()
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, filename)
            self.log_to_gui(f"File Selected: {os.path.basename(filename)}")

    def _get_saved_user_id(self):
        if not self.manual_credentials:
            return ""
        return str(self.manual_credentials[0].get("User ID", "")).strip()

    def _refresh_manual_controls(self):
        has_manual = bool(self.manual_credentials)
        self.btn_view_id.configure(state="normal" if has_manual else "disabled")
        self.btn_delete_id.configure(state="normal" if has_manual else "disabled")
        if has_manual:
            user_id = self._get_saved_user_id()
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, f"Selected ID: {user_id}")

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
        self.entry_file.delete(0, "end")
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

        ctk.CTkLabel(card, text="User ID / PAN").pack(anchor="w")
        ent_user = ctk.CTkEntry(card, placeholder_text="Enter User ID / PAN")
        ent_user.pack(fill="x", pady=(4, 10))

        ctk.CTkLabel(card, text="Password").pack(anchor="w")
        pass_frm = ctk.CTkFrame(card, fg_color="transparent")
        pass_frm.pack(fill="x", pady=(4, 14))
        ent_pass = ctk.CTkEntry(pass_frm, placeholder_text="Enter Password", show="*")
        ent_pass.pack(side="left", expand=True, fill="x")

        def _toggle_pass():
            if ent_pass.cget("show") == "":
                ent_pass.configure(show="*")
                eye_btn.configure(text="👁")
            else:
                ent_pass.configure(show="")
                eye_btn.configure(text="🔒")

        eye_btn = ctk.CTkButton(pass_frm, text="👁", width=35, height=30,
                                fg_color="transparent", text_color=("#475569", "#94a3b8"),
                                hover_color=("#e2e8f0", "#334155"), command=_toggle_pass)
        eye_btn.pack(side="right", padx=(5, 0))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x")

        def _save():
            user_id = (ent_user.get() or "").strip()
            password = (ent_pass.get() or "").strip()
            if not user_id or not password:
                messagebox.showerror("Missing Data", "Please enter User ID/PAN and Password", parent=dialog)
                return

            existing_user = self._get_saved_user_id()
            if existing_user and not messagebox.askyesno(
                "Overwrite ID",
                "Your previous ID will be overwritten with this.",
                parent=dialog
            ):
                return

            self.manual_credentials = [{"User ID": user_id, "Password": password}]
            self.excel_file_path = ""
            self._refresh_manual_controls()
            messagebox.showinfo("Added", f"Credential saved for {user_id}", parent=dialog)
            dialog.destroy()

        ctk.CTkButton(btn_row, text="Cancel", width=110, command=dialog.destroy).pack(side="right")
        ctk.CTkButton(btn_row, text="Add", width=110, command=_save).pack(side="right", padx=(0, 8))

        ent_user.focus_set()
        dialog.bind("<Return>", lambda _e: _save())

    def _create_manual_excel(self):
        rows = []
        for item in self.manual_credentials:
            user_id = str(item.get("User ID", "")).strip()
            password = str(item.get("Password", "")).strip()
            if user_id and password:
                rows.append({"User ID": user_id, "Password": password})

        if not rows:
            return ""

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="it_itr_manual_") as tmp:
            temp_excel = tmp.name
        pd.DataFrame(rows, columns=["User ID", "Password"]).to_excel(temp_excel, index=False)
        return temp_excel

    def log_to_gui(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def start_process(self):
        excel_path = self.excel_file_path
        if not excel_path and self.manual_credentials:
            excel_path = self._create_manual_excel()

        if not excel_path:
            messagebox.showwarning("Missing File", "Please add ID/Password first.")
            return

        selected_pref = self.combo_years.get()

        self.btn_start.configure(state="disabled", text="PROCESSING...", fg_color="gray")
        self.btn_stop.grid()
        self.btn_open_folder.grid_remove()
        self.progress_bar.set(0)
        self.log_to_gui("-" * 30)
        self.log_to_gui("Starting Worker Thread...")

        self.worker = IncomeTaxWorker(self, excel_path, selected_pref)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if self.worker:
            self.worker.keep_running = False
        self.btn_stop.configure(state="disabled", text="Stopping...")

    # --- THREAD SAFE GUI UPDATES ---
    def update_log_safe(self, message):
        self.after(0, lambda: self.log_to_gui(message))

    def update_progress_safe(self, value):
        self.after(0, lambda: self.progress_bar.set(value))

    def process_finished_safe(self, message):
        def _finish():
            self.log_to_gui(f"\nSTATUS: {message}")
            self.btn_start.configure(state="normal", text="INITIATE BATCH DOWNLOAD", fg_color="#059669")
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.btn_stop.grid_remove()
            self.btn_open_folder.grid(row=0, column=2, padx=(10, 0))
            messagebox.showinfo("Process Complete", message)
        self.after(0, _finish)

    def open_output_folder(self):
        try:
            target = os.path.join(os.getcwd(), "Income Tax Downloaded", "ITR Bot")
            if not os.path.exists(target):
                target = os.path.join(os.getcwd(), "Income Tax Downloaded")
            os.startfile(target)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()