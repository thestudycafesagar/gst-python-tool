import threading
import time
import os
import random
import glob
import sys
import base64
import pandas as pd
import customtkinter as ctk
from PIL import Image
from datetime import datetime
from tkinter import filedialog, messagebox

# EXE-compatible shared stealth driver
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from stealth_driver import create_chrome_driver, show_browser_alert, build_chrome_options

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

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

    def _save_captcha_image(self, wait, output_path):
        """Capture captcha image bytes reliably (avoids DPI crop mismatch on Windows)."""
        captcha_img = wait.until(EC.presence_of_element_located((By.ID, "imgCaptcha")))

        wait.until(lambda d: d.execute_script(
            """
            const img = document.getElementById('imgCaptcha');
            return !!(img && img.complete && (img.naturalWidth || img.width) > 0 && (img.naturalHeight || img.height) > 0);
            """
        ))

        src = captcha_img.get_attribute("src") or ""
        if src.startswith("data:image"):
            payload = src.split(",", 1)[1]
            with open(output_path, "wb") as f:
                f.write(base64.b64decode(payload))
            return

        try:
            data_url = self.driver.execute_script(
                """
                const img = document.getElementById('imgCaptcha');
                if (!img) return null;
                const w = img.naturalWidth || img.width;
                const h = img.naturalHeight || img.height;
                const canvas = document.createElement('canvas');
                canvas.width = w;
                canvas.height = h;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, w, h);
                return canvas.toDataURL('image/png');
                """
            )
            if isinstance(data_url, str) and data_url.startswith("data:image"):
                payload = data_url.split(",", 1)[1]
                with open(output_path, "wb") as f:
                    f.write(base64.b64decode(payload))
                return
        except Exception:
            pass

        with open(output_path, "wb") as f:
            f.write(captcha_img.screenshot_as_png)

    def log(self, message):
        self.app.update_log_safe(message)

    def human_delay(self, base_s=5.0, extra_s=1.5):
        time.sleep(base_s + random.uniform(0.0, extra_s))

    def type_like_human(self, element, text):
        element.clear()
        for ch in str(text):
            element.send_keys(ch)
            time.sleep(random.uniform(0.06, 0.18))

    def run(self):
        self.log("🚀 INITIALIZING GST ENGINE V17 (Auto-Retry + Healing)...")
        
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
            base_dir = os.path.join(os.getcwd(), "GST_3B_Downloads")
            if not os.path.exists(base_dir): os.makedirs(base_dir)

            # 3. PROCESS LOOP
            stopped_by_user = False
            for index, row in df.iterrows():
                if not self.keep_running:
                    stopped_by_user = True
                    break

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

    def _robust_find_clickable(self, by, value, timeout=10, refreshes=2, alert_msg="Element not found"):
        """Wait for element. If not found, refresh and retry. If still not found, show on-page alert."""
        for attempt in range(refreshes + 1):
            try:
                el = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((by, value)))
                if el:
                    return el
            except Exception:
                pass

            if attempt < refreshes:
                self.log(f"   ⚠️ '{value}' not found. Refreshing page (Attempt {attempt+1}/{refreshes})...")
                try:
                    self.driver.refresh()
                except:
                    pass
                time.sleep(4)

        self.log(f"   ❌ Search failed! {alert_msg}")
        show_browser_alert(self.driver, f"⚠️ GST Portal Issue: {alert_msg}. Portal may be slow or temporarily blocking requests. Retrying next round...")
        return None

    def _show_browser_alert(self, message):
        """Convenience wrapper so existing call-sites don't break."""
        show_browser_alert(self.driver, message)

    def ensure_return_dashboard(self, wait):
        """ Checks if we are on the Return Dashboard. If not, go directly to the URL. """
        try:
            # Check if Financial Year dropdown is present
            if self.driver.find_elements(By.NAME, "fin"):
                return True # We are safe
            
            self.log("      🔄 Lost Dashboard. Recovering via URL...")
            self.driver.execute_script("window.location.href = 'https://return.gst.gov.in/returns/auth/dashboard';")
            time.sleep(3)
            
            # Check if it threw us to the login page (session dropped)
            if "login" in self.driver.current_url.lower() or "access denied" in self.driver.page_source.lower():
                return False

            if self.driver.find_elements(By.NAME, "fin"): 
                return True
            return False
            
        except Exception as e:
            self.log(f"      ⚠️ Recovery Exception: {e}")
            return False

    def process_single_user(self, username, password, user_root):
        try:
            # --- BROWSER SETUP (ANTI-DETECT + EXE COMPATIBLE) ---
            # Options / stealth / driver all handled by stealth_driver module.
            self.driver = create_chrome_driver(build_chrome_options())


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
            
            # 3. EXECUTE LOOP
            self.human_delay()
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
                        year_el = self._robust_find_clickable(By.NAME, "fin", timeout=5, refreshes=2, alert_msg="Year dropdown missing")
                        if not year_el: raise Exception("Year dropdown missing")
                        Select(year_el).select_by_visible_text(fin_year)
                        time.sleep(0.5)

                        # 3. Select Quarter
                        qtr_el = self._robust_find_clickable(By.NAME, "quarter", timeout=5, refreshes=1, alert_msg="Quarter dropdown missing")
                        if not qtr_el: raise Exception("Quarter dropdown missing")
                        Select(qtr_el).select_by_visible_text(q_text)
                        time.sleep(0.5)

                        # 4. Select Month
                        mon_el = self._robust_find_clickable(By.NAME, "mon", timeout=5, refreshes=1, alert_msg="Month dropdown missing")
                        if not mon_el: raise Exception("Month dropdown missing")
                        Select(mon_el).select_by_visible_text(m_text)
                        time.sleep(0.5)

                        # 5. Click Search
                        search_btn = self._robust_find_clickable(By.XPATH, "//button[contains(text(), 'Search')]", timeout=5, refreshes=1, alert_msg="Search button missing")
                        if not search_btn: raise Exception("Search button missing")
                        try: search_btn.click()
                        except: self.driver.execute_script("arguments[0].click();", search_btn)
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
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None

    def perform_login(self, username, password, wait):
        self.log("   🌐 Opening GST Portal...")
        self.driver.get("https://services.gst.gov.in/services/login")

        while True:
            if not self.keep_running: return False, "Stopped"

            try:
                user_box = wait.until(EC.visibility_of_element_located((By.ID, "username")))
                self.type_like_human(user_box, username)

                pass_box = self.driver.find_element(By.ID, "user_pass")
                self.type_like_human(pass_box, password)

                self._save_captcha_image(wait, "temp_captcha.png")
                
                self.log("   ⌨️ Waiting for Captcha...")
                self.captcha_response = None
                self.captcha_event.clear()
                self.app.request_captcha_safe("temp_captcha.png")
                self.captcha_event.wait() 

                if not self.captcha_response: return False, "Captcha Cancelled"

                cap_box = self.driver.find_element(By.ID, "captcha")
                self.type_like_human(cap_box, self.captcha_response)
                self.driver.find_element(By.XPATH, "//button[@type='submit']").click()
                
                self.human_delay()

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
            found_btn = self._robust_find_clickable(By.XPATH, xpath_3b_btn, timeout=5, refreshes=1, alert_msg="GSTR-3B Target Tile missing")
            if not found_btn:
                self.log("   ⚠️ GSTR-3B Tile Not Found.")
                return False, "Not Found"

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
        self.grid_rowconfigure(3, weight=1)

        self.worker = None
        self.excel_file = ""
        self.manual_credentials = []

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1D4ED8", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False) 
        ctk.CTkLabel(self.head, text="GST BULK DOWNLOADER", 
                      font=("Segoe UI", 24, "bold"), text_color="white").pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(self.head, text="GSTR-3B AUTOMATION", 
                      font=("Segoe UI", 14), text_color="#CBD5E1").pack(side="right", padx=20, pady=15)

        # SETTINGS
        self.settings_container = ctk.CTkFrame(self, fg_color="transparent")
        self.settings_container.grid(row=1, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.settings_container.grid_columnconfigure((0, 1), weight=1)

        # Credentials Card
        self.card_cred = ctk.CTkFrame(self.settings_container, border_color="#334155", border_width=1)
        self.card_cred.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ctk.CTkLabel(self.card_cred, text="📂 Credentials Source", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))
        cred_row = ctk.CTkFrame(self.card_cred, fg_color="transparent")
        cred_row.pack(fill="x", padx=15, pady=(5, 15))

        self.ent_file = ctk.CTkEntry(cred_row, placeholder_text="Add ID/Password manually (optional)...", height=35)
        self.ent_file.pack(side="left", expand=True, fill="x", padx=(0, 10))

        action_row = ctk.CTkFrame(cred_row, fg_color="transparent")
        action_row.pack(side="right")
        self.btn_download = ctk.CTkButton(action_row, text="➕ Add ID Password", command=self.add_id_password,
                          fg_color="#059669", hover_color="#047857", height=35, width=150,
                          font=("Segoe UI", 12, "bold"))
        self.btn_download.pack(side="left", padx=(0, 8))
        self.btn_demo = ctk.CTkButton(action_row, text="▶ View Demo", command=self.open_demo_link,
                          fg_color="#DC2626", hover_color="#B91C1C", height=35, width=150,
                          font=("Segoe UI", 12, "bold"))
        self.btn_demo.pack(side="left")

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
        self.btn_view_id.configure(state="disabled")
        self.btn_delete_id.configure(state="disabled")

        # Period Settings Card
        self.card_period = ctk.CTkFrame(self.settings_container, border_color="#334155", border_width=1)
        self.card_period.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        ctk.CTkLabel(self.card_period, text="📅 Period Selection", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))

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
        
        self.chk_all_qtr_var = ctk.BooleanVar(value=False)
        self.period_mode_var = ctk.StringVar(value="Monthly")
        self.frm_mode = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_mode.pack(fill="x", padx=15, pady=(4, 6))
        ctk.CTkLabel(self.frm_mode, text="Mode:", width=100, anchor="w").pack(side="left")
        self.mode_tabs = ctk.CTkSegmentedButton(
            self.frm_mode,
            values=["Monthly", "Quarterly"],
            variable=self.period_mode_var,
            command=self.toggle_inputs,
            width=180
        )
        self.mode_tabs.pack(side="right", expand=True, fill="x")

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
        self.toggle_inputs()

        # LOGS
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self.log_frame, text="📜 Execution Logs", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), text_color="#10B981", height=150)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.log_box.configure(state="disabled")

        # CAPTCHA (Hidden)
        self.cap_frame = ctk.CTkFrame(self, border_color="#DC2626", border_width=2, fg_color="#1E293B")
        self.cap_inner = ctk.CTkFrame(self.cap_frame, fg_color="transparent")
        self.cap_inner.pack(fill="both", padx=20, pady=10)
        ctk.CTkLabel(self.cap_inner, text="⚠️ CAPTCHA ACTION REQUIRED", text_color="#EF4444", font=("Segoe UI", 14, "bold")).pack()
        self.cap_lbl_img = ctk.CTkLabel(self.cap_inner, text="")
        self.cap_lbl_img.pack(pady=10)
        self.cap_ent = ctk.CTkEntry(self.cap_inner, placeholder_text="Type Captcha Here...", 
                                    font=("Consolas", 20), justify="center", height=45, width=250)
        self.cap_ent.pack(pady=5)
        self.cap_ent.bind("<Return>", self.submit_captcha) 
        # --- CAPTCHA BUTTONS SIDE-BY-SIDE ---
        self.cap_btn_row = ctk.CTkFrame(self.cap_inner, fg_color="transparent")
        self.cap_btn_row.pack(pady=(10, 0))
        self.cap_btn = ctk.CTkButton(self.cap_btn_row, text="SUBMIT CAPTCHA", fg_color="#DC2626", hover_color="#B91C1C",
                         height=40, width=120, font=("Segoe UI", 12, "bold"), command=self.submit_captcha)
        self.cap_btn.pack(side="left", padx=(0, 10))
        self.cap_stop_btn = ctk.CTkButton(self.cap_btn_row, text="⏹ STOP PROCESS", fg_color="#475569", hover_color="#334155",
                          height=40, width=120, font=("Segoe UI", 11, "bold"), command=self.stop_process)
        self.cap_stop_btn.pack(side="left")

        # FOOTER
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 20))
        self.prog_bar = ctk.CTkProgressBar(self.footer, height=15, progress_color="#10B981")
        self.prog_bar.pack(fill="x", pady=(0, 10))
        self.prog_bar.set(0)
        self.btn_row_footer = ctk.CTkFrame(self.footer, fg_color="transparent")
        self.btn_row_footer.pack(fill="x")
        self.btn_start = ctk.CTkButton(self.btn_row_footer, text="START BATCH PROCESS", height=50, font=("Segoe UI", 16, "bold"),
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

    def toggle_inputs(self, mode_choice=None):
        if mode_choice and hasattr(self, "period_mode_var"):
            self.period_mode_var.set(mode_choice)
        self.cb_qtr.configure(state="normal")
        mode = self.period_mode_var.get() if hasattr(self, "period_mode_var") else "Monthly"
        self.cb_month.configure(state="normal")
        self.update_months_based_on_qtr(self.cb_qtr.get())
        self.cb_month.configure(state="disabled" if mode == "Quarterly" else "normal")

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
        webbrowser.open_new_tab("https://www.youtube.com/watch?v=XXXXXXXXXX")

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
            self.btn_start.configure(state="normal", text="STOPPED" if is_stopped else "START BATCH PROCESS")
            self.btn_stop.pack_forget()
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.cap_stop_btn.configure(state="normal", text="⏹ STOP PROCESS")
            if is_stopped:
                self.after(1200, lambda: self.btn_start.configure(text="START BATCH PROCESS"))
            else:
                self.btn_open_folder.pack(side="left", padx=(10, 0))
        self.after(0, _finish_ui)

    def open_output_folder(self):
        try:
            target = os.path.join(os.getcwd(), "GST_3B_Downloads")
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
            self.cap_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
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
            "all_quarters": False
        }
        self.close_captcha_safe()
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
        self.cap_stop_btn.configure(state="disabled", text="STOPPED")
        self.update_log_safe("🛑 Process stopped by user.")

if __name__ == "__main__":
    app = App()
    app.mainloop()