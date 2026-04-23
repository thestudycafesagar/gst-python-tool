import threading
import time
import os
import random
import glob
import base64
import pandas as pd
import customtkinter as ctk
from PIL import Image
from datetime import datetime
from tkinter import filedialog, messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

# Shared Stealth Driver Import
import sys
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _ROOT not in sys.path: sys.path.insert(0, _ROOT)
from stealth_driver import create_chrome_driver, build_chrome_options

# --- UI CONFIGURATION ---
# Commented out: theme is controlled globally by GST_Suite.py
# ctk.set_default_color_theme("blue")

try:
    _CAPTCHA_RESAMPLE = Image.Resampling.NEAREST
except AttributeError:
    _CAPTCHA_RESAMPLE = Image.NEAREST

class GSTWorker:
    def __init__(self, app_instance, excel_path, settings, credentials=None):
        self.app = app_instance
        self.excel_path = excel_path
        self.settings = settings
        self.credentials = credentials or []
        self.keep_running = True
        self.driver = None
        self.captcha_response = None 
        self.captcha_event = threading.Event()
        self.report_data = [] 
        
        # Save credentials for auto-recovery
        self.current_user = ""
        self.current_pass = ""

    def log(self, message):
        self.app.update_log_safe(message)

    def human_delay(self, min_s=5.0, max_s=6.5):
        """Randomized delay with 5s baseline to mimic human behavior."""
        time.sleep(random.uniform(min_s, max_s))

    def type_like_human(self, element, text):
        element.clear()
        for ch in str(text):
            element.send_keys(ch)
            time.sleep(random.uniform(0.06, 0.18))

    def run(self):
        self.log("🚀 INITIALIZING GSTR-1 PDF ENGINE (Final Version)...")
        
        try:
            # 1. LOAD CREDENTIALS (manual IDs preferred, Excel optional)
            if self.credentials:
                df = pd.DataFrame(self.credentials)
                user_col, pass_col = "Username", "Password"
                self.log(f"📊 Loaded {len(df)} users from Add ID Password.")
            else:
                if not self.excel_path:
                    self.app.process_finished_safe("Please add ID/Password first")
                    return

                df = pd.read_excel(self.excel_path)
                clean_cols = {c.lower().strip(): c for c in df.columns}
                user_col = next((clean_cols[c] for c in clean_cols if 'user' in c or 'name' in c), None)
                pass_col = next((clean_cols[c] for c in clean_cols if 'pass' in c or 'pwd' in c), None)

                if not user_col or not pass_col:
                    self.app.process_finished_safe("Column Error: Need Username/Password columns")
                    return
                self.log(f"📊 Loaded {len(df)} users from Excel.")

            if df.empty:
                self.app.process_finished_safe("No credentials found to process")
                return

            total = len(df)

            # 2. CREATE MAIN DOWNLOAD FOLDER
            base_dir = os.path.join(os.getcwd(), "GST Downloaded", "GSTR1 PDF")
            if not os.path.exists(base_dir): os.makedirs(base_dir, exist_ok=True)

            # 3. PROCESS LOOP
            stopped_by_user = False
            for index, row in df.iterrows():
                if not self.keep_running:
                    stopped_by_user = True
                    break

                self.current_user = str(row[user_col]).strip()
                self.current_pass = str(row[pass_col]).strip()
                
                self.app.update_progress_safe((index) / total)
                self.log(f"\n🔹 Processing: {self.current_user}")
                
                # Unique Folder Versioning
                user_root_base = os.path.join(base_dir, self.current_user)
                user_root = user_root_base
                counter = 1
                while os.path.exists(user_root):
                    user_root = f"{user_root_base}_{counter}"
                    counter += 1
                os.makedirs(user_root)

                status, reason = self.process_single_user(self.current_user, self.current_pass, user_root)
                
                self.report_data.append({
                    "Username": self.current_user,
                    "Status": status,
                    "Details": reason,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Saved To": os.path.basename(user_root)
                })

                if not self.keep_running:
                    stopped_by_user = True
                    break
                
                self.log("-" * 40)

            if stopped_by_user or not self.keep_running:
                if self.report_data:
                    self.generate_excel_report()
                    self.log("🛑 Process stopped by user. Partial report saved.")
                    self.app.process_finished_safe("Stopped by user. Partial report saved.")
                else:
                    self.log("🛑 Process stopped by user.")
                    self.app.process_finished_safe("Stopped by user.")
                return

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
            filename = f"GSTR1_PDF_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_df.to_excel(filename, index=False)
            self.log(f"📄 Summary Report saved: {filename}")
        except Exception as e:
            self.log(f"⚠️ Failed to save report: {e}")

    def handle_popups(self):
        """ Checks for Aadhaar/Generic Popups and closes them stealthily. """
        self.human_delay(1.5, 2.5) 
        xpath = "//a[contains(text(), 'No-Remind me later')] | //a[contains(text(),'Remind me later')] | //button[contains(text(),'Remind Me Later')]"
        try:
            popups = self.driver.find_elements(By.XPATH, xpath)
            for popup in popups:
                if popup.is_displayed():
                    self.log("      🛡️ Dismissing 'Remind me later' popup...")
                    popup.click()
                    self.human_delay(1.0, 2.0)
        except: 
            pass

    def ensure_return_dashboard(self, wait):
        """ 
        Checks if we are on the Return Dashboard.
        If logged out completely, logs back in. If on Home, navigates back.
        """
        try:
            src = self.driver.page_source
            # 1. Check for Logout / Access Denied
            if "Access Denied!" in src or "session is expired" in src or self.driver.find_elements(By.ID, "username"):
                self.log("      🔄 Session Expired! Auto Re-logging in...")
                self.perform_login(self.current_user, self.current_pass, wait)
                return True

            # 2. Check for Dashboard (Fin Year Dropdown)
            if self.driver.find_elements(By.NAME, "fin"):
                return True 
            
            self.log("      🔄 Lost Dashboard. Recovering...")
            self.handle_popups()

            try:
                # 1. Try hitting the "Back" button if it exists
                back_btns = self.driver.find_elements(By.XPATH, "//button[contains(translate(normalize-space(), 'BACK', 'back'), 'back')]")
                for b in back_btns:
                    if b.is_displayed():
                        try:
                            b.click()
                            time.sleep(2)
                            return True
                        except: pass

                # 2. Try Return Dashboard breadcrumb explicitly
                bread_btns = self.driver.find_elements(By.XPATH, "//a[contains(translate(normalize-space(), 'DASHBOARD', 'dashboard'), 'dashboard')]")
                for b in bread_btns:
                    if b.is_displayed():
                        try:
                            b.click()
                            time.sleep(2)
                            return True
                        except: pass

                dash_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Return Dashboard')] | //a[contains(., 'Return Dashboard')]")))
                dash_btn.click()
                self.human_delay(3.0, 4.0)
                return True
            except:
                try:
                    # Fallback JS Click
                    btn = self.driver.find_element(By.XPATH, "//button[contains(., 'Return Dashboard')]")
                    self.driver.execute_script("arguments[0].click();", btn)
                    self.human_delay(3.0, 4.0)
                    return True
                except:
                    self.log("      ⚠️ Recovery Failed: Dashboard Button missing.")
                    return False
        except Exception as e:
            self.log(f"      ⚠️ Recovery Exception: {e}")
            return False

    def _robust_find_clickable(self, by, value, timeout=10, refreshes=2, alert_msg="Element not found"):
        """Wait for element. If not found, refresh and retry. If still not found, show alert."""
        for attempt in range(refreshes + 1):
            try:
                el = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((by, value)))
                if el: return el
            except Exception: pass
                
            if attempt < refreshes:
                self.log(f"   ⚠️ '{value}' not found. Refreshing page (Attempt {attempt+1}/{refreshes})...")
                try: self.driver.refresh()
                except: pass
                time.sleep(4)
                
        self.log(f"   ❌ Search failed! {alert_msg}")
        self.app.after(0, lambda: messagebox.showwarning("Portal Issue", f"{alert_msg}. The browser may be detecting automation or the portal is slow."))
        return None

    def process_single_user(self, username, password, user_root):
        try:
            # --- BROWSER SETUP (SHARED STEALTH DRIVER) ---
            self.driver = create_chrome_driver(build_chrome_options(user_root))
            
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                    window.navigator.chrome = { runtime: {} };
                    Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
                    Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
                """
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

            selected_q = self.settings['quarter']
            period_mode = self.settings.get('period_mode', 'Monthly')
            if selected_q not in q_map:
                return "Config Error", "Invalid Month/Quarter Selection"

            if period_mode == "Quarterly":
                selected_m = q_map[selected_q][-1]
                tasks = [{"q": selected_q, "m": selected_m}]
                self.log(f"   📅 Mode: Quarterly ({selected_q} -> {selected_m})")
            else:
                selected_m = self.settings['month']
                if selected_m not in q_map[selected_q]:
                    return "Config Error", "Invalid Month/Quarter Selection"
                tasks = [{"q": selected_q, "m": selected_m}]
                self.log(f"   📅 Mode: Monthly ({selected_m})")
            self.log(f"   📅 Queued {len(tasks)} Months...")
            
            # 3. EXECUTE LOOP
            self.human_delay()
            success_count = 0
            results = []

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
                
                max_retries = 3
                month_success = False
                
                for attempt in range(1, max_retries + 1):
                    attempt_str = f" [Attempt {attempt}]" if attempt > 1 else ""
                    self.log(f"   ⚙️ Processing: {m_text} ({q_text[:9]}){attempt_str}")

                    try:
                        # 1. Dashboard Check & Healing
                        if not self.ensure_return_dashboard(wait):
                            self.log(f"      ⚠️ Dashboard check failed. Retrying...")
                            self.human_delay()
                            continue 

                        # 2. Selection Logic (Year -> Quarter -> Month -> Search)
                        year_el = self._robust_find_clickable(By.NAME, "fin", timeout=5, refreshes=2, alert_msg="Year dropdown missing")
                        if not year_el: raise Exception("Year missing")
                        Select(year_el).select_by_visible_text(fin_year)
                        self.human_delay(1.5, 3.0)

                        qtr_el = self._robust_find_clickable(By.NAME, "quarter", timeout=5, refreshes=1, alert_msg="Quarter dropdown missing")
                        if not qtr_el: raise Exception("Quarter missing")
                        Select(qtr_el).select_by_visible_text(q_text)
                        self.human_delay(1.5, 3.0)

                        mon_el = self._robust_find_clickable(By.NAME, "mon", timeout=5, refreshes=1, alert_msg="Month dropdown missing")
                        if not mon_el: raise Exception("Month missing")
                        Select(mon_el).select_by_visible_text(m_text)
                        self.human_delay(1.5, 3.0)

                        search_btn = self._robust_find_clickable(By.XPATH, "//button[contains(text(), 'Search')]", timeout=5, refreshes=1, alert_msg="Search missing")
                        if not search_btn: raise Exception("Search missing")
                        search_btn.click()
                        self.human_delay(3.0, 5.0) 

                        # 3. Process PDF Download
                        dl_status, dl_msg = self.process_gstr1_pdf(wait, year_folder, m_text)
                        
                        if dl_status:
                            success_count += 1
                            results.append(f"{m_text}: ✅")
                            month_success = True
                            
                            # Navigate back safely
                            try: self.driver.back() 
                            except: pass
                            self.human_delay(2.0, 3.0)
                            
                            # Double check if we need to back again (Dashboard level)
                            if "gstr1" in self.driver.current_url.lower():
                                self.driver.back()
                                self.human_delay(2.0, 3.0)

                            break 
                        else:
                            # If "Not Filed", we count it as success (process complete) but note it
                            if "Not Filed" in dl_msg:
                                success_count += 1
                                results.append(f"{m_text}: ⚠️ Not Filed")
                                month_success = True # Stop retrying
                                
                                # We might be inside the View page or on Dashboard
                                if "dashboard" not in self.driver.current_url.lower():
                                    self.driver.back()
                                    self.human_delay(2.0, 3.0)
                                break

                            self.log(f"      ⚠️ Failed ({dl_msg}). Retrying...")
                            if attempt < max_retries:
                                try: self.driver.back()
                                except: pass
                                self.human_delay(2.0, 3.0)
                                
                    except Exception as e:
                        self.log(f"      ❌ Error: {str(e)[:30]}")
                        if attempt < max_retries:
                            self.log("      🔄 Refreshing page...")
                            try: 
                                self.driver.refresh()
                                self.human_delay(3.0, 5.0)
                            except: pass
                
                if not month_success:
                    results.append(f"{m_text}: ❌ Failed")

            overall_status = "Success" if success_count == len(tasks) else "Partial"
            if success_count == 0: overall_status = "Failed"
            
            summary = f"Processed {success_count}/{len(tasks)}. Details: " + ", ".join(results)
            return overall_status, summary

        except Exception as e:
            return "Error", f"Browser Crash: {str(e)[:30]}"
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None

    def perform_login(self, username, password, wait):
        self.log("   🚀 MANUAL LOGIN MODE.")
        self.log("   👉 Please LOGIN manually in the Chrome window.")
        self.driver.maximize_window()
        self.driver.get("https://services.gst.gov.in/services/login")
        
        while self.keep_running:
            try:
                url = self.driver.current_url.lower()
                src = self.driver.page_source.lower()
                
                # Strictly check for post-login indicators
                is_logged_in = any(k in url for k in ("dashboard", "auth/home", "services/auth")) or \
                               len(self.driver.find_elements(By.XPATH, "//a[contains(@href, 'logout')]")) > 0
                
                if is_logged_in:
                    # Double check we're not still on the login page/captcha
                    if not self.driver.find_elements(By.ID, "imgCaptcha"):
                        self.log("   ✅ Login detected!")
                        break
            except Exception:
                pass
            time.sleep(2)
        
        if not self.keep_running: 
            return False, "Stopped"
        
        time.sleep(2)
        self.handle_popups()
        
        # Navigate to Return Dashboard if not already there
        try:
            dash_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Return Dashboard')]")))
            self.driver.execute_script("arguments[0].click();", dash_btn)
            return True, "Success"
        except Exception:
            try:
                btn = self.driver.find_element(By.XPATH, "//button[contains(., 'Return Dashboard')]")
                self.driver.execute_script("arguments[0].click();", btn)
                return True, "Success (JS Click)"
            except Exception:
                if "dashboard" in self.driver.current_url.lower():
                    return True, "Success (Manual/Detected)"
                self.log("   ⚠️ Dashboard Nav Error.")
                return False, "Dashboard Nav Failed"

    def process_gstr1_pdf(self, wait, download_path, month):
        """ Downloads the GSTR-1 PDF using the precise buttons provided. """
        self.log(f"      🔍 Searching for GSTR-1 Tile...")

        # 1. Click GSTR-1 'VIEW' button on the Dashboard
        xpath_r1_view = "//p[contains(text(),'GSTR1')]/ancestor::div[contains(@class,'col-')]//button[contains(normalize-space(),'VIEW')]"
        view_btn = self._robust_find_clickable(By.XPATH, xpath_r1_view, timeout=5, refreshes=1, alert_msg="GSTR-1 Target Tile missing")
        if not view_btn:
            self.log("      ⚠️ GSTR-1 Tile/View button missing.")
            return False, "Tile Not Found"
        view_btn.click()
        self.human_delay(3.0, 5.0) # Wait for View page

        # 2. Click 'VIEW SUMMARY' Button
        self.log(f"      📄 Finding 'VIEW SUMMARY'...")
        summary_xpath = "//button[contains(@data-ng-click, 'gstr1sum')] | //span[contains(text(), 'VIEW SUMMARY')]/parent::button"
        summary_btn = self._robust_find_clickable(By.XPATH, summary_xpath, timeout=5, refreshes=1, alert_msg="View Summary missing (Likely Not Filed)")
        if not summary_btn:
            return False, "Not Filed"
        
        if "disabled" in summary_btn.get_attribute("class"):
             self.log("      ⚠️ View Summary button is disabled.")
             return False, "Not Filed"

        summary_btn.click()
        self.human_delay(3.0, 5.0) # Wait for Summary page

        # 3. Click 'DOWNLOAD (PDF)' Button
        self.log(f"      ⬇️ Clicking 'DOWNLOAD (PDF)'...")
        try:
            pdf_xpath = "//button[contains(@data-ng-click, 'genratepdfNew')] | //span[contains(text(), 'DOWNLOAD (PDF)')]/parent::button"
            pdf_btn = self._robust_find_clickable(By.XPATH, pdf_xpath, timeout=5, refreshes=1, alert_msg="Download PDF button missing")
            if not pdf_btn: return False, "Download Button Missing"
            pdf_btn.click()
            
            # 4. Wait for File
            self.log("      ⏳ Waiting for file to save...")
            file_downloaded = False
            for _ in range(25): 
                time.sleep(1)
                files = glob.glob(os.path.join(download_path, "*.pdf"))
                if files:
                    latest = max(files, key=os.path.getctime)
                    if (datetime.now().timestamp() - os.path.getctime(latest)) < 30:
                        self.log(f"      ✅ Saved: {os.path.basename(latest)}")
                        file_downloaded = True
                        break
            
            if not file_downloaded:
                self.log("      ⚠️ PDF download timed out.")
                return False, "Timeout"

            return True, "Success"
            
        except Exception as e:
            self.log(f"      ⚠️ Failed to find 'DOWNLOAD (PDF)' button.")
            return False, "Download Button Missing"


# --- GUI CLASS ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GST Bulk Downloader - GSTR-1 PDF Edition")
        self.geometry("900x850")
        # ctk.set_appearance_mode("System")  # controlled by GST_Suite.py
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.worker = None
        self.excel_file = ""
        self.manual_credentials = []

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1D4ED8", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False) 
        ctk.CTkLabel(self.head, text="GST BULK DOWNLOADER", 
                      font=("Segoe UI", 24, "bold"), text_color="white").pack(side="left", padx=20, pady=10)
        
        # Theme ComboBox — removed; theme is controlled by GST_Suite.py
        # self.theme_cb = ctk.CTkComboBox(self.head, values=["System", "Dark", "Light"],
        #                                 command=self.change_theme, width=100, ...)
        # self.theme_cb.set("System")
        # self.theme_cb.pack(side="right", padx=20, pady=15)

        # CONTENT AREA (SCROLLABLE)
        self.scroll_container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll_container.grid_columnconfigure(0, weight=1)

        # SETTINGS
        self.settings_container = ctk.CTkFrame(self.scroll_container, fg_color="transparent")
        self.settings_container.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 10))
        self.settings_container.grid_columnconfigure((0, 1), weight=1)

        # Credentials
        self.card_cred = ctk.CTkFrame(self.settings_container, border_color="#334155", border_width=1)
        self.card_cred.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ctk.CTkLabel(self.card_cred, text="📂 Credentials Source", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))
        self.ent_file = ctk.CTkEntry(self.card_cred, placeholder_text="Add ID/Password manually (optional)...", height=35)
        self.ent_file.pack(fill="x", padx=15, pady=(5, 10))
        btn_row = ctk.CTkFrame(self.card_cred, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(5, 15))
        self.btn_download = ctk.CTkButton(btn_row, text="➕ Add ID Password", command=self.add_id_password, fg_color="#059669", hover_color="#047857", height=28, font=("Segoe UI", 12, "bold"))
        self.btn_download.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.btn_demo = ctk.CTkButton(btn_row, text="▶ View Demo", command=self.open_demo_link, fg_color="#DC2626", hover_color="#B91C1C", height=28, font=("Segoe UI", 12, "bold"))
        self.btn_demo.pack(side="left", expand=True, fill="x", padx=(5, 0))
        manage_row = ctk.CTkFrame(self.card_cred, fg_color="transparent")
        manage_row.pack(fill="x", padx=15, pady=(0, 10))
        self.btn_view_id = ctk.CTkButton(manage_row, text="👁 View ID", command=self.view_saved_user,
                         fg_color="#475569", hover_color="#334155", height=28, width=100,
                         font=("Segoe UI", 11, "bold"))
        self.btn_view_id.pack(side="left")
        self.btn_delete_id = ctk.CTkButton(manage_row, text="🗑 Delete ID", command=self.delete_saved_user,
                           fg_color="#7C3AED", hover_color="#6D28D9", height=28, width=110,
                           font=("Segoe UI", 11, "bold"))
        self.btn_delete_id.pack(side="left", padx=(8, 0))
        self.btn_delete_id.configure(state="disabled")



        # Period
        self.card_period = ctk.CTkFrame(self.settings_container, border_color="#334155", border_width=1)
        self.card_period.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        ctk.CTkLabel(self.card_period, text="📅 Period Selection", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))

        # Static Year List
        year_list = ["2022-23", "2023-24", "2024-25", "2025-26", "2026-27"]

        self.frm_year = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_year.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_year, text="Financial Year:", width=140, anchor="w").pack(side="left")
        self.cb_year = ctk.CTkComboBox(self.frm_year, values=year_list, width=150)
        self.cb_year.set(year_list[0]) 
        self.cb_year.pack(side="right", expand=True, fill="x")
        
        self.chk_all_qtr_var = ctk.BooleanVar(value=False)
        self.period_mode_var = ctk.StringVar(value="Monthly")
        self.frm_mode = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_mode.pack(fill="x", padx=15, pady=(4, 6))
        ctk.CTkLabel(self.frm_mode, text="Filing Frequency:", width=140, anchor="w").pack(side="left")
        self.mode_tabs = ctk.CTkSegmentedButton(
            self.frm_mode,
            values=["Monthly", "Quarterly"],
            variable=self.period_mode_var,
            command=self.toggle_inputs,
            width=180
        )
        self.mode_tabs.pack(side="right", expand=True, fill="x")

        self.frm_qtr = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_qtr.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_qtr, text="Quarter:", width=140, anchor="w").pack(side="left")
        self.cb_qtr = ctk.CTkComboBox(self.frm_qtr, 
                                      values=["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)", 
                                              "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"],
                                      command=self.update_months_based_on_qtr, width=150)
        self.cb_qtr.set("Quarter 1 (Apr - Jun)")
        self.cb_qtr.pack(side="right", expand=True, fill="x")

        self.frm_mon = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_mon.pack(fill="x", padx=15, pady=(2, 15))
        ctk.CTkLabel(self.frm_mon, text="Month:", width=140, anchor="w").pack(side="left")
        
        all_months = ["April", "May", "June", "July", "August", "September", 
                      "October", "November", "December", "January", "February", "March"]
        self.cb_month = ctk.CTkComboBox(self.frm_mon, values=all_months, 
                                        command=self.update_qtr_based_on_month, width=150)
        self.cb_month.set("April")
        self.cb_month.pack(side="right", expand=True, fill="x")
        self.toggle_inputs()

        # CAPTCHA SECTION
        self.cap_frame = ctk.CTkFrame(self.scroll_container, border_color="#DC2626", border_width=1)
        self.cap_frame.grid_columnconfigure(0, weight=1)
        
        cap_inner = ctk.CTkFrame(self.cap_frame, fg_color="transparent")
        cap_inner.pack(pady=10, padx=10, fill="x")
        
        self.cap_lbl_img = ctk.CTkLabel(cap_inner, text="", image=None)
        self.cap_lbl_img.pack(side="left", padx=10)
        
        self.cap_ent = ctk.CTkEntry(cap_inner, placeholder_text="Enter Captcha", width=120, height=35)
        self.cap_ent.pack(side="left", padx=10)
        self.cap_ent.bind("<Return>", self.submit_captcha)
        
        self.cap_btn = ctk.CTkButton(cap_inner, text="SUBMIT", command=self.submit_captcha, width=100, height=35)
        self.cap_btn.pack(side="left", padx=5)
        
        self.cap_stop_btn = ctk.CTkButton(cap_inner, text="⏹ STOP", command=self.stop_process, width=100, height=35, fg_color="#475569", hover_color="#334155")
        self.cap_stop_btn.pack(side="left", padx=5)

        # LOGS
        self.log_frame = ctk.CTkFrame(self.scroll_container)
        self.log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self.log_frame, text="📜 Execution Logs", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), text_color="#10B981", height=250)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.log_box.configure(state="disabled")



        # Footer
        self.footer = ctk.CTkFrame(self.scroll_container, fg_color="transparent")
        self.footer.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 20))
        self.prog_bar = ctk.CTkProgressBar(self.footer, height=15, progress_color="#10B981")
        self.prog_bar.pack(fill="x", pady=(0, 10))
        self.prog_bar.set(0)
        self.btn_row_footer = ctk.CTkFrame(self.footer, fg_color="transparent")
        self.btn_row_footer.pack(fill="x")
        self.btn_start = ctk.CTkButton(self.btn_row_footer, text="START PDF DOWNLOAD", height=50, font=("Segoe UI", 16, "bold"),
                                       fg_color="#047857", hover_color="#047857", command=self.start_process)
        self.btn_start.pack(side="left", expand=True, fill="x")
        self.btn_stop = ctk.CTkButton(self.btn_row_footer, text="⏹ STOP", height=50, font=("Segoe UI", 16, "bold"),
                                      fg_color="#DC2626", hover_color="#B91C1C", command=self.stop_process, width=150)
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_stop.pack_forget()

        self.btn_open_folder = ctk.CTkButton(self.btn_row_footer, text="📂 OPEN FOLDER", height=50, font=("Segoe UI", 16, "bold"),
                                      fg_color="#2563EB", hover_color="#1D4ED8", command=self.open_output_folder, width=180)
        self.btn_open_folder.pack(side="left", padx=(10, 0))
        self.btn_open_folder.pack_forget()

    def change_theme(self, choice):
        pass  # Theme controlled by GST_Suite.py

    def toggle_inputs(self, mode_choice=None):
        if mode_choice and hasattr(self, "period_mode_var"):
            self.period_mode_var.set(mode_choice)
        
        mode = self.period_mode_var.get() if hasattr(self, "period_mode_var") else "Monthly"
        
        if mode == "Monthly":
            self.cb_qtr.configure(state="disabled")
            self.cb_month.configure(state="normal")
            self.cb_month.configure(values=["April", "May", "June", "July", "August", "September", 
                                          "October", "November", "December", "January", "February", "March"])
            self.update_qtr_based_on_month(self.cb_month.get())
        else:
            self.cb_qtr.configure(state="normal")
            self.cb_month.configure(state="disabled")
            self.update_months_based_on_qtr(self.cb_qtr.get())

    def update_qtr_based_on_month(self, choice):
        mapping = {
            "April": "Quarter 1 (Apr - Jun)", "May": "Quarter 1 (Apr - Jun)", "June": "Quarter 1 (Apr - Jun)",
            "July": "Quarter 2 (Jul - Sep)", "August": "Quarter 2 (Jul - Sep)", "September": "Quarter 2 (Jul - Sep)",
            "October": "Quarter 3 (Oct - Dec)", "November": "Quarter 3 (Oct - Dec)", "December": "Quarter 3 (Oct - Dec)",
            "January": "Quarter 4 (Jan - Mar)", "February": "Quarter 4 (Jan - Mar)", "March": "Quarter 4 (Jan - Mar)"
        }
        qtr_val = mapping.get(choice, "Quarter 1 (Apr - Jun)")
        
        # Temporarily enable to set value, then restore state
        curr_state = self.cb_qtr.cget("state")
        self.cb_qtr.configure(state="normal")
        self.cb_qtr.set(qtr_val)
        self.cb_qtr.configure(state=curr_state)

    def update_months_based_on_qtr(self, choice):
        if "Quarter 1" in choice: vals = ["April", "May", "June"]
        elif "Quarter 2" in choice: vals = ["July", "August", "September"]
        elif "Quarter 3" in choice: vals = ["October", "November", "December"]
        elif "Quarter 4" in choice: vals = ["January", "February", "March"]
        else: vals = ["April", "May", "June"]
        if not hasattr(self, "cb_month"):
            return
        prev_month_state = self.cb_month.cget("state")
        if prev_month_state != "normal":
            self.cb_month.configure(state="normal")
        self.cb_month.configure(values=vals)
        mode = self.period_mode_var.get() if hasattr(self, "period_mode_var") else "Monthly"
        self.cb_month.set(vals[-1] if mode == "Quarterly" else vals[0])
        if prev_month_state != "normal":
            self.cb_month.configure(state=prev_month_state)

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
            self.ent_file.delete(0, "end")
            self.ent_file.insert(0, f"Selected ID: {user_id}")

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
        self.ent_file.delete(0, "end")
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
        pass_frm = ctk.CTkFrame(card, fg_color="transparent")
        pass_frm.pack(fill="x", pady=(4, 14))
        ent_pass = ctk.CTkEntry(pass_frm, placeholder_text="Enter GST Password", show="*")
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
            self.excel_file = ""
            self._refresh_manual_controls()
            messagebox.showinfo("Added", f"Credential saved for {username}", parent=dialog)
            dialog.destroy()

        ctk.CTkButton(btn_row, text="Cancel", width=110, command=dialog.destroy).pack(side="right")
        ctk.CTkButton(btn_row, text="Add", width=110, command=_save).pack(side="right", padx=(0, 8))

        ent_user.focus_set()
        dialog.bind("<Return>", lambda _e: _save())

    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://youtu.be/0pCbHNTEar8")

    def browse_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            self.excel_file = f
            self.manual_credentials = []
            self._refresh_manual_controls()
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
        def _finish_ui():
            messagebox.showinfo("Info", msg)
            is_stopped = "stopped" in (msg or "").lower()
            self.close_captcha_safe()
            self.btn_start.configure(state="normal", text="STOPPED" if is_stopped else "START PDF DOWNLOAD")
            self.btn_stop.pack_forget()
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.cap_stop_btn.configure(state="normal", text="⏹ STOP PROCESS")
            if is_stopped:
                self.after(1200, lambda: self.btn_start.configure(text="START PDF DOWNLOAD"))
            else:
                self.btn_open_folder.pack(side="left", padx=(10, 0))
        self.after(0, _finish_ui)

    def open_output_folder(self):
        try:
            target = os.path.join(os.getcwd(), "GST Downloaded", "GSTR1 PDF")
            if not os.path.exists(target):
                target = os.getcwd()
            os.startfile(target)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {e}")

    def request_captcha_safe(self, img_path):
        def show():
            if not self.worker or not self.worker.keep_running:
                return
            with Image.open(img_path) as raw_img:
                pil_img = raw_img.convert("RGB")

            w, h = pil_img.size
            if w <= 0 or h <= 0:
                return

            max_w, max_h = 230, 78
            scale = min(max_w / float(w), max_h / float(h), 1.0)
            if scale < 1.0:
                size = (max(1, int(w * scale)), max(1, int(h * scale)))
                display_img = pil_img.resize(size, _CAPTCHA_RESAMPLE)
            else:
                size = (w, h)
                display_img = pil_img

            self._captcha_ctk_img = ctk.CTkImage(light_image=display_img, dark_image=display_img, size=size)
            self.cap_lbl_img.image = self._captcha_ctk_img
            self.cap_lbl_img.configure(image=self._captcha_ctk_img)
            self.cap_btn.configure(state="normal", text="SUBMIT CAPTCHA", fg_color="#DC2626")
            self.cap_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
            self.cap_ent.delete(0, "end")
            self.attributes('-topmost', True)
            self.deiconify()
            self.lift()
            def _focus():
                self.focus_force()
                self.cap_ent.focus_set()
                self.after(1000, lambda: self.attributes('-topmost', False))
            self.after(200, _focus)
        self.after(0, show)

    def submit_captcha(self, event=None):
        txt = self.cap_ent.get()
        if not txt: return
        self.cap_btn.configure(state="disabled", text="VERIFYING...", fg_color="gray")
        self.worker.captcha_response = txt
        self.worker.captcha_event.set()

        self.worker.captcha_event.set()

    def close_captcha_safe(self):
        self.after(0, lambda: self.cap_frame.grid_forget())

    def start_process(self):
        credentials = list(self.manual_credentials)
        if not credentials and not self.excel_file:
            messagebox.showerror("Error", "Please add ID/Password first")
            return
        settings = {
            "year": self.cb_year.get(),
            "month": self.cb_month.get(),
            "quarter": self.cb_qtr.get(),
            "period_mode": self.period_mode_var.get(),
            "all_quarters": False,
            "manual_login": self.chk_manual_login_var.get()
        }
        self.close_captcha_safe()
        if hasattr(self, 'cap_stop_btn'):
            self.cap_stop_btn.configure(state="normal", text="⏹ STOP PROCESS")
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_open_folder.pack_forget()
        self.worker = GSTWorker(self, self.excel_file, settings, credentials=credentials)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if not self.worker:
            return

        self.worker.keep_running = False
        self.worker.captcha_response = None
        self.worker.captcha_event.set()

        try:
            if self.worker.driver:
                self.worker.driver.quit()
                self.worker.driver = None
                self.update_log_safe("🛑 Chrome browser closed.")
        except Exception as e:
            self.update_log_safe(f"⚠️ Error closing Chrome: {e}")

        self.close_captcha_safe()
        self.btn_stop.configure(state="disabled", text="STOPPED")
        if hasattr(self, 'cap_stop_btn'):
            self.cap_stop_btn.configure(state="disabled", text="STOPPED")
        self.update_log_safe("🛑 Process stopped by user.")

if __name__ == "__main__":
    app = App()
    app.mainloop()