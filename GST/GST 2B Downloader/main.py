import threading
import time
import os
import sys
import random
import glob
import base64
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

# Shared Stealth Driver Import
import sys
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _ROOT not in sys.path: sys.path.insert(0, _ROOT)
from stealth_driver import create_chrome_driver, build_chrome_options, show_browser_alert



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

    def _mask_user(self, username):
        if not username:
            return ""
        u = str(username)
        if len(u) <= 4:
            return "*" * len(u)
        return f"{u[:2]}{'*' * (len(u) - 4)}{u[-2:]}"

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
        self.log("🚀 INITIALIZING GST ENGINE V17 (Hybrid Selection)...")
        
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
            base_dir = os.path.join(os.getcwd(), "GST Downloaded", "GSTR 2B")
            if not os.path.exists(base_dir): os.makedirs(base_dir, exist_ok=True)

            # 3. PROCESS LOOP
            stopped_by_user = False
            for index, row in df.iterrows():
                if not self.keep_running:
                    stopped_by_user = True
                    break

                username = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                
                self.app.update_progress_safe((index) / total)
                self.log(f"\n🔹 Processing: {self._mask_user(username)}")
                
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
            filename = f"GST_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_df.to_excel(filename, index=False)
            self.log(f"📄 Summary Report saved: {filename}")
        except Exception as e:
            self.log(f"⚠️ Failed to save report: {e}")

    def process_single_user(self, username, password, user_root):
        """ Returns (Overall Status, Reason String) """
        try:
            # --- BROWSER SETUP (ANTI-DETECT + EXE COMPATIBLE) ---
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-extensions")
            # --- BROWSER SETUP (SHARED STEALTH DRIVER) ---
            self.driver = create_chrome_driver(build_chrome_options(user_root))
            self.driver.maximize_window()

            # ── COMPREHENSIVE STEALTH JS ─────────────────────────────────────
            # Injected before any page script runs — masks all automation flags
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """
                    // Hide webdriver flag
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});

                    // Spoof real Chrome runtime
                    window.navigator.chrome = {
                        runtime: {},
                        loadTimes: function() {},
                        csi: function() {},
                        app: {}
                    };

                    // Realistic plugin list
                    Object.defineProperty(navigator, 'plugins', {
                        get: () => [
                            { name: 'Chrome PDF Plugin',   filename: 'internal-pdf-viewer',              length: 1 },
                            { name: 'Chrome PDF Viewer',   filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai',  length: 1 },
                            { name: 'Native Client',       filename: 'internal-nacl-plugin',             length: 2 }
                        ]
                    });

                    // Real language / platform
                    Object.defineProperty(navigator, 'languages',          { get: () => ['en-IN', 'en', 'en-US'] });
                    Object.defineProperty(navigator, 'platform',           { get: () => 'Win32' });
                    Object.defineProperty(navigator, 'hardwareConcurrency',{ get: () => 8 });
                    Object.defineProperty(navigator, 'deviceMemory',       { get: () => 8 });

                    // Permissions: avoid headless notifications flag
                    const _origQuery = window.navigator.permissions.query.bind(navigator.permissions);
                    window.navigator.permissions.query = (p) =>
                        p.name === 'notifications'
                            ? Promise.resolve({ state: Notification.permission })
                            : _origQuery(p);
                """
            })

            wait = WebDriverWait(self.driver, 20)
            
            # 1. LOGIN
            login_status, login_msg = self.perform_login(username, password, wait)
            if not login_status: return "Login Failed", login_msg

            # 2. DEFINE TASKS
            # Map Quarters to Months
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

            # Create SINGLE Year Folder
            fin_year = self.settings['year']
            year_folder = os.path.join(user_root, fin_year)
            if not os.path.exists(year_folder): os.makedirs(year_folder)

            # Set Download Path Once
            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow", 
                "downloadPath": year_folder
            })

            for task in tasks:
                if not self.keep_running: return "Stopped", "User Cancelled"
                
                q_text = task['q']
                m_text = task['m']
                
                # --- RETRY LOGIC (Max 3 Attempts) ---
                month_success = False
                fail_reason = ""
                
                for attempt in range(1, 4): 
                    self.log(f"   ⚙️ Processing: {m_text} (Attempt {attempt})")
                    
                    try:
                        # ── SESSION GUARD ─────────────────────────────────────
                        # Check for Access Denied / expired session BEFORE every
                        # attempt; re-login transparently and keep going.
                        if not self.check_session_and_relogin(username, password, wait):
                            fail_reason = "Re-login Failed"
                            break  # no point retrying if re-login itself fails

                        self.log("   ⏳ Checking dashboard readiness...")
                        time.sleep(1)

                        # Select Year
                        year_el = self._robust_find_clickable(By.NAME, "fin", timeout=8, refreshes=2, alert_msg="Year selection element (name='fin') not found")
                        if not year_el:
                            fail_reason = "Year (fin) not found. Blocked or slow."
                            break
                        Select(year_el).select_by_visible_text(fin_year)
                        self.driver.execute_script(
                            "var e=arguments[0]; angular.element(e).triggerHandler('change');", year_el)
                        time.sleep(1)

                        # Select Quarter — must trigger Angular ng-change so months load
                        qtr_el = self._robust_find_clickable(By.NAME, "quarter", timeout=5, refreshes=1, alert_msg="Quarter selection element not found")
                        if not qtr_el:
                            fail_reason = "Quarter selection not found"
                            break
                        Select(qtr_el).select_by_visible_text(q_text)
                        self.driver.execute_script(
                            "var e=arguments[0]; angular.element(e).triggerHandler('change');", qtr_el)
                        time.sleep(1.5)  # wait for month dropdown to populate

                        # Select Month
                        mon_el = self._robust_find_clickable(By.NAME, "mon", timeout=5, refreshes=1, alert_msg="Month selection element not found")
                        if not mon_el:
                            fail_reason = "Month selection not found"
                            break
                        Select(mon_el).select_by_visible_text(m_text)
                        self.driver.execute_script(
                            "var e=arguments[0]; angular.element(e).triggerHandler('change');", mon_el)
                        time.sleep(0.5)

                        # Click Search
                        search_btn = self._robust_find_clickable(By.XPATH, "//button[contains(text(), 'Search')]", timeout=5, refreshes=1, alert_msg="Search button not found")
                        if not search_btn:
                            fail_reason = "Search button not found"
                            break
                        self.driver.execute_script("arguments[0].click();", search_btn)
                        time.sleep(5)

                        # Download
                        dl_status, dl_msg = self.download_gstr2b(wait, year_folder)
                        
                        if dl_status:
                            # Validate session once more right after download page flow.
                            # Some GST redirects briefly land on access-denied/login pages.
                            if not self.check_session_and_relogin(username, password, wait):
                                fail_reason = "Re-login Failed after Download"
                                self.log("      ⚠️ Session dropped after download. Retrying month...")
                                continue
                            month_success = True
                            success_count += 1
                            results.append(f"{m_text}: ✅")
                            break # Success, exit retry loop
                        else:
                            fail_reason = dl_msg
                            if "Not Generated" in dl_msg:
                                break 
                            self.log(f"      ⚠️ Attempt {attempt} failed: {dl_msg}")

                    except Exception as e:
                        fail_reason = f"Error: {str(e)[:30]}"
                        self.log(f"      ❌ Exception: {str(e)[:50]}")
                        try: self.driver.get("https://return.gst.gov.in/returns/auth/dashboard")
                        except: pass
                
                if not month_success:
                    results.append(f"{m_text}: ❌ ({fail_reason})")

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

    def handle_popups(self):
        """Bypass Aadhaar and other GST portal popups."""
        try:
            # Aadhaar Remind Later
            aadhaar_skip = self.driver.find_elements(By.XPATH, "//a[contains(text(),'Remind me later')]")
            if aadhaar_skip and aadhaar_skip[0].is_displayed():
                aadhaar_skip[0].click()
                time.sleep(1)
        except: pass

        try:
            # Generic Remind Me Later
            generic_skip = self.driver.find_elements(By.XPATH, "//button[contains(text(),'Remind Me Later')]")
            if generic_skip and generic_skip[0].is_displayed():
                generic_skip[0].click()
                time.sleep(1)
        except: pass

    def perform_login(self, username, password, wait):
        self.log("   🔐 Attempting login (auto-fill if credentials provided)...")
        self.driver.maximize_window()
        self.driver.get("https://services.gst.gov.in/services/login")

        # Try to auto-fill credentials if available
        try:
            if username:
                try:
                    usr = WebDriverWait(self.driver, 6).until(EC.visibility_of_element_located((By.ID, "username")))
                    self.type_like_human(usr, username)
                except Exception:
                    usr = None

                try:
                    pwd = self.driver.find_element(By.ID, "user_pass")
                    self.type_like_human(pwd, password)
                except Exception:
                    pwd = None

                # Try clicking submit — captcha may be required and will be handled manually
                try:
                    btn = self.driver.find_element(By.XPATH, "//button[@type='submit']")
                    self.driver.execute_script("arguments[0].click();", btn)
                except Exception:
                    pass

                # If captcha present, show a browser banner
                try:
                    if self.driver.find_elements(By.ID, "imgCaptcha"):
                        show_browser_alert(self.driver, "Please enter captcha in the browser and submit to continue.")
                        self.log("   🟨 Captcha detected — please complete it in the browser.")
                except Exception:
                    pass
        except Exception as e:
            self.log(f"   ⚠️ Auto-fill issue: {str(e)[:20]}")

        while self.keep_running:
            try:
                url = (self.driver.current_url or "").lower()
                src = (self.driver.page_source or "").lower()
                
                # Check for post-login indicators
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
        time.sleep(2)
        self.handle_popups()

        # Navigate to Return Dashboard
        try:
            dash_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Return Dashboard')]")))
            dash_btn.click()
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
                self._show_browser_alert("⚠️ Dashboard Nav Failed")
                return False, "Dashboard Nav Failed"

            except Exception as e:
                self.log(f"   ⚠️ Login Exception: {e}")
                return False, f"Login Error: {str(e)[:20]}"

    def _robust_find_clickable(self, by, value, timeout=10, refreshes=2, alert_msg="Element not found"):
        """Wait for element. If not found, refresh and retry. If still not found, show on-page alert."""
        for attempt in range(refreshes + 1):
            try:
                # Use presence instead of element_to_be_clickable to avoid Angular overlay/click-intercept failures
                el = WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((by, value)))
                if el:
                    # Give Angular a tiny moment to render it fully
                    time.sleep(0.5)
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
        self._show_browser_alert(f"⚠️ GST Portal Issue: {alert_msg}. Portal may be slow or temporarily blocking this request. Retrying next round...")
        return None

    def _wait_for_recent_download(self, download_path, started_at, timeout=60):
        """Wait for a new downloaded file created after started_at."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                # Some GST flows expose a secondary download link after processing.
                links = self.driver.find_elements(By.XPATH, "//a[contains(text(), 'Click here to download')]")
                if links and links[0].is_displayed():
                    self.driver.execute_script("arguments[0].click();", links[0])
            except Exception:
                pass

            try:
                files = glob.glob(os.path.join(download_path, "*"))
                fresh_files = []
                for f in files:
                    if not os.path.isfile(f):
                        continue
                    lf = f.lower()
                    if lf.endswith(".crdownload") or lf.endswith(".tmp"):
                        continue
                    if os.path.getctime(f) >= (started_at - 0.2):
                        fresh_files.append(f)

                if fresh_files:
                    return max(fresh_files, key=os.path.getctime)
            except Exception:
                pass

            time.sleep(1)
        return None

    def _go_to_return_dashboard(self):
        """Move to Return Dashboard directly via URL as requested by user."""
        try:
            self.log("      🔄 Lost Dashboard. Recovering via URL...")
            self.driver.execute_script("window.location.href = 'https://return.gst.gov.in/returns/auth/dashboard';")
            time.sleep(3)
            
            # Check if it threw us to the login page (session dropped)
            url_lower = (self.driver.current_url or "").lower()
            src_lower = (self.driver.page_source or "").lower()
            if "login" in url_lower or "access denied" in src_lower or "session is expired" in src_lower:
                return False

            return True
        except Exception:
            return False

    def _session_snapshot(self, stage):
        """Small diagnostic logger for URL/session state transitions."""
        try:
            url = self.driver.current_url
            src_l = (self.driver.page_source or "").lower()
            flags = []
            if "accessdenied" in (url or "").lower() or "access denied" in src_l:
                flags.append("access-denied")
            if "session is expired" in src_l:
                flags.append("session-expired")
            if "services.gst.gov.in/services/login" in (url or ""):
                flags.append("login-page")
            state = ", ".join(flags) if flags else "ok"
            self.log(f"   🧭 [{stage}] URL: {url}")
            self.log(f"   🧭 [{stage}] Session: {state}")
        except Exception:
            pass

    def _accept_alerts_if_any(self, max_count=4):
        """Accept JS confirm/alert popups if present."""
        accepted = 0
        for _ in range(max_count):
            try:
                al = self.driver.switch_to.alert
                msg = (al.text or "").strip()
                al.accept()
                accepted += 1
                if msg:
                    self.log(f"   ℹ️ Accepted popup: {msg[:70]}")
                time.sleep(0.6)
            except Exception:
                break
        return accepted

    def _show_browser_alert(self, message):
        """Inject a visible red banner directly into the browser page.
        Shows the error ON the website — never shows a Python messagebox popup.
        Safe to call even if driver is unavailable.
        """
        if not self.driver:
            self.log(f"   ⚠️ [alert] {message}")
            return
        try:
            safe_msg = message.replace("'", "\\'").replace("`", "\\`")
            js = f"""
            (function() {{
                var _id = '__gst_tool_alert__';
                var old = document.getElementById(_id);
                if (old) old.remove();
                var d = document.createElement('div');
                d.id = _id;
                d.innerText = '{safe_msg}';
                d.style.cssText = [
                    'position:fixed', 'top:0', 'left:0', 'right:0',
                    'z-index:2147483647',
                    'background:#DC2626',
                    'color:#fff',
                    'font-size:15px',
                    'font-weight:bold',
                    'font-family:Arial,sans-serif',
                    'padding:14px 24px',
                    'text-align:center',
                    'box-shadow:0 3px 12px rgba(0,0,0,0.5)',
                    'letter-spacing:0.3px'
                ].join(';');
                document.body.prepend(d);
                setTimeout(function() {{ if (d.parentNode) d.remove(); }}, 10000);
            }})();
            """
            self.driver.execute_script(js)
        except Exception:
            # Silently log — never crash the thread over an alert
            self.log(f"   ⚠️ [portal-alert] {message}")

    def _download_gstr2b_portal_buttons(self, wait, download_path, summary_btn_xpath, details_btn_xpath):
        """Download using GST portal's new Summary/Details buttons."""
        summary_btn = self._robust_find_clickable(By.XPATH, summary_btn_xpath, timeout=8, refreshes=1, alert_msg="GSTR-2B Summary download button missing")
        if not summary_btn:
            return False, "Portal Controls Missing"

        summary_started = time.time()
        self.log("   ⬇️ Downloading GSTR-2B Summary (PDF)...")
        try:
            summary_btn.click()
        except:
            self.driver.execute_script("arguments[0].click();", summary_btn)
            
        summary_file = self._wait_for_recent_download(download_path, summary_started, timeout=60)
        if not summary_file:
            return False, "Summary Timeout"
        self.log(f"   ✅ Saved: {os.path.basename(summary_file)}")

        details_btn = self._robust_find_clickable(By.XPATH, details_btn_xpath, timeout=8, refreshes=0, alert_msg="GSTR-2B Details download button missing")
        if not details_btn:
            return False, "Details Controls Missing"

        details_started = time.time()
        self.log("   ⬇️ Downloading GSTR-2B Details (Excel)...")
        try:
            details_btn.click()
        except:
            self.driver.execute_script("arguments[0].click();", details_btn)
            
        details_file = self._wait_for_recent_download(download_path, details_started, timeout=70)
        if not details_file:
            return False, "Details Timeout"
        self.log(f"   ✅ Saved: {os.path.basename(details_file)}")
        return True, "Success"

    def _download_gstr2b_computax_controls(self, download_path):
        """Download using CompuTax page controls (pdf/xls icons)."""
        self.log("   🧩 Trying CompuTax-style controls (pdf/xls)...")

        pdf_xpath = "//span[contains(@class,'pdf') and contains(@title,'Download Return PDF')]"
        xls_xpath = "//span[contains(@class,'xls') and contains(@title,'Download Return Excel')]"
        popup_download_xpath = "//button[contains(normalize-space(),'Download Return Excel') or contains(@onclick,'DownloadReturnExcelV2')]"

        xls_btn = self._robust_find_clickable(By.XPATH, xls_xpath, timeout=8, refreshes=1, alert_msg="CompuTax Excel download icon missing")
        if not xls_btn:
            return False, "CompuTax Controls Missing"

        # 1) PDF click (best-effort)
        try:
            pdf_elems = self.driver.find_elements(By.XPATH, pdf_xpath)
            if pdf_elems:
                pdf_btn = pdf_elems[0]
                pdf_started = time.time()
                try:
                    pdf_btn.click()
                except:
                    self.driver.execute_script("arguments[0].click();", pdf_btn)
                self._accept_alerts_if_any()
                pdf_file = self._wait_for_recent_download(download_path, pdf_started, timeout=45)
                if pdf_file:
                    self.log(f"   ✅ Saved: {os.path.basename(pdf_file)}")
                else:
                    self.log("   ⚠️ PDF not downloaded via CompuTax controls (continuing).")
        except Exception:
            self.log("   ⚠️ PDF click failed on CompuTax controls (continuing).")

        # 2) Excel click (required)
        try:
            xls_started = time.time()
            try:
                xls_btn.click()
            except:
                self.driver.execute_script("arguments[0].click();", xls_btn)
            time.sleep(1.2)
            self._accept_alerts_if_any(max_count=5)

            # If popup renders a secondary "Download Return Excel" button, click it.
            popup_btns = self.driver.find_elements(By.XPATH, popup_download_xpath)
            if popup_btns:
                try:
                    try:
                        popup_btns[0].click()
                    except:
                        self.driver.execute_script("arguments[0].click();", popup_btns[0])
                    self._accept_alerts_if_any(max_count=3)
                except Exception:
                    pass

            xls_file = self._wait_for_recent_download(download_path, xls_started, timeout=80)
            if not xls_file:
                return False, "CompuTax Excel Timeout"
            self.log(f"   ✅ Saved: {os.path.basename(xls_file)}")
            return True, "Success"
        except Exception:
            return False, "CompuTax Excel Error"

    def download_gstr2b(self, wait, download_path):
        """ Returns (Bool, Message) """
        self.log("   🔍 Searching for GSTR-2B tile view button...")
        
        period_mode = self.settings.get('period_mode', 'Monthly')

        if period_mode == "Quarterly":
            view_btn_xpaths = [
                "//div[contains(@class,'col-sm-4') and .//p[contains(normalize-space(),'Quarterly View') or contains(normalize-space(),'GSTR2BQ')]]//button[contains(normalize-space(),'View')]",
                "//p[contains(normalize-space(),'Quarterly View') or contains(normalize-space(),'GSTR2BQ')]/ancestor::div[contains(@class,'col-sm-4')]//button[contains(normalize-space(),'View')]"
            ]
        else:
            view_btn_xpaths = [
                "//div[contains(@class,'col-sm-4') and .//p[contains(normalize-space(),'GSTR2B') or contains(normalize-space(),'GSTR2BQ')]]//button[normalize-space()='View']",
                "//div[contains(@class,'col-sm-4') and .//p[contains(normalize-space(),'GSTR2B') or contains(normalize-space(),'GSTR2BQ')]]//button[contains(normalize-space(),'View')]",
                "//p[contains(normalize-space(),'GSTR2B') or contains(normalize-space(),'GSTR2BQ')]/ancestor::div[contains(@class,'col-sm-4')]//button[contains(normalize-space(),'View')]",
            ]

        view_btn = None
        for attempt in range(2): # allow 1 refresh if not found
            for xpath in view_btn_xpaths:
                try:
                    view_btn = WebDriverWait(self.driver, 8).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    if view_btn:
                        break
                except Exception:
                    continue
                    
            if view_btn:
                break
            else:
                if attempt < 1:
                    self.log("   ⚠️ View button not found. Refreshing page...")
                    try:
                        self.driver.refresh()
                    except:
                        pass
                    time.sleep(4)

        if not view_btn:
            self.log("   ⚠️ GSTR-2B View button not found.")
            self.app.after(0, lambda: messagebox.showwarning("Portal Error", "Class name not found (GSTR-2B View button). GST portal may be slow or layout changed. Detected as missing!"))
            return False, "View Missing"

        summary_btn_xpath = (
            "//button[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), "
            "'DOWNLOAD GSTR-2B SUMMARY')]"
        )
        details_btn_xpath = (
            "//button[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), "
            "'DOWNLOAD GSTR-2B DETAILS') and contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'EXCEL')]"
        )
        computax_pdf_xpath = "//span[contains(@class,'pdf') and contains(@title,'Download Return PDF')]"
        computax_xls_xpath = "//span[contains(@class,'xls') and contains(@title,'Download Return Excel')]"

        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", view_btn)
            time.sleep(1)
            try:
                view_btn.click()
            except:
                self.driver.execute_script("arguments[0].click();", view_btn)
            self.log("   ✅ View clicked. Opening GSTR-2B summary page...")
            self._session_snapshot("after-view")

            WebDriverWait(self.driver, 30).until(
                lambda d: (
                    "gstr2b" in d.current_url.lower()
                    or d.find_elements(By.XPATH, summary_btn_xpath)
                    or d.find_elements(By.XPATH, details_btn_xpath)
                    or d.find_elements(By.XPATH, computax_pdf_xpath)
                    or d.find_elements(By.XPATH, computax_xls_xpath)
                )
            )

            time.sleep(1.5)
            src_l = self.driver.page_source.lower()
            if "no record" in src_l or "compute your gstr 2b" in src_l or "not generated" in src_l:
                self.log("   ⚠️ GSTR-2B not generated for selected period.")
                self._go_to_return_dashboard()
                return False, "Not Generated"

            portal_ok, portal_msg = self._download_gstr2b_portal_buttons(
                wait, download_path, summary_btn_xpath, details_btn_xpath
            )
            if not portal_ok and portal_msg == "Portal Controls Missing":
                legacy_ok, legacy_msg = self._download_gstr2b_computax_controls(download_path)
                if not legacy_ok:
                    self.log(f"   ⚠️ CompuTax-style flow failed: {legacy_msg}")
                    self.app.after(0, lambda: messagebox.showwarning("Portal Error", "Class name not found. Both Portal and CompuTax buttons failed to render."))
                    self._go_to_return_dashboard()
                    self._session_snapshot("after-dashboard-nav")
                    return False, legacy_msg
            elif not portal_ok:
                self.log(f"   ⚠️ Portal-style flow failed: {portal_msg}")
                self._go_to_return_dashboard()
                self._session_snapshot("after-dashboard-nav")
                return False, portal_msg

            self._go_to_return_dashboard()
            self._session_snapshot("after-dashboard-nav")
            return True, "Success"

        except Exception as e:
            self.log(f"   ⚠️ View/Download flow error: {str(e)[:40]}")
            self._go_to_return_dashboard()
            self._session_snapshot("after-dashboard-nav")
            return False, "Script Error"

    def check_session_and_relogin(self, username, password, wait):
        """
        Detects 'Access Denied / session expired' and automatically re-logs in.
        Returns True if the session is valid (or was successfully restored).
        Returns False if re-login itself fails.
        """
        try:
            time.sleep(1)  # let page stabilise before checking
            current_url = self.driver.current_url
            src         = self.driver.page_source

            on_login_page = (
                "services.gst.gov.in/services/login" in current_url
                or bool(self.driver.find_elements(By.ID, "username"))
            )
            hard_expired = (
                "session is expired" in src.lower()
                or "access denied" in src.lower()
                or "accessdenied" in current_url.lower()
            )
            is_expired = on_login_page or hard_expired

            if is_expired:
                self.log("   🔄 Session expired — re-logging in...")
                login_ok, login_msg = self.perform_login(username, password, wait)
                if not login_ok:
                    self.log(f"   ❌ Re-login failed: {login_msg}")
                    return False
                self.log("   ✅ Re-login successful — resuming task.")
            return True
        except Exception as e:
            self.log(f"   ⚠️ Session check error (ignored): {e}")
            return True  # assume alive if the check itself throws


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
        self._captcha_ctk_img = None
        self.runtime_log_path = os.path.join(os.getcwd(), "gst2b_runtime.log")
        self._log_file_lock = threading.Lock()

        try:
            with open(self.runtime_log_path, "w", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] GST 2B runtime log started\n")
        except Exception:
            pass

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1D4ED8", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False) 
        ctk.CTkLabel(self.head, text="GST BULK DOWNLOADER", 
                     font=("Segoe UI", 24, "bold"), text_color="white").pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(self.head, text="Powered by StudyCafe", 
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
        self.btn_delete_id.configure(state="disabled")



        # Period Settings Card
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

        # Quarter & Month
        self.frm_qtr = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_qtr.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_qtr, text="Quarter:", width=140, anchor="w").pack(side="left")
        self.cb_qtr = ctk.CTkComboBox(self.frm_qtr, 
                                      values=["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)", 
                                              "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"],
                                      command=self.update_months_based_on_qtr, width=150)
        self.cb_qtr.set("Quarter 1 (Apr - Jun)")
        self.cb_qtr.pack(side="right", expand=True, fill="x")

        # Month
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
        self.cap_frame = ctk.CTkFrame(self, border_color="#DC2626", border_width=1)
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
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self.log_frame, text="📜 Execution Logs", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), text_color="#10B981", height=150)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.log_box.configure(state="disabled")



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
        try:
            ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with self._log_file_lock:
                with open(self.runtime_log_path, "a", encoding="utf-8") as f:
                    f.write(f"[{ts}] {msg}\n")
        except Exception:
            pass
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



    def start_process(self):
        credentials = list(self.manual_credentials)
        if not credentials and not self.excel_file:
            messagebox.showerror("Error", "Please add ID/Password first")
            return
        self.update_log_safe(f"📄 Runtime log file: {self.runtime_log_path}")
        settings = {
            "year": self.cb_year.get(),
            "month": self.cb_month.get(),
            "quarter": self.cb_qtr.get(),
            "period_mode": self.period_mode_var.get(),
            "all_quarters": False
        }
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.pack(side="left", padx=(10, 0))
        if hasattr(self, 'btn_open_folder'):
            self.btn_open_folder.pack_forget()
        self.worker = GSTWorker(self, self.excel_file, settings, credentials=credentials)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if not self.worker:
            self._reset_ui_state()
            return

        self.worker.keep_running = False
        self.worker.captcha_response = None

        try:
            if self.worker.driver:
                self.worker.driver.quit()
                self.worker.driver = None
                self.update_log_safe("🛑 Chrome browser closed.")
        except Exception as e:
            self.update_log_safe(f"⚠️ Error closing Chrome: {e}")

        self.close_captcha_safe()
        self.btn_stop.pack_forget()
        self.btn_stop.configure(state="disabled", text="STOPPED")
        if hasattr(self, 'cap_stop_btn'):
            self.cap_stop_btn.configure(state="disabled", text="STOPPED")
        self.update_log_safe("🛑 Process stopped by user. Resetting state.")
        self._reset_ui_state()
        self.worker = None

    def _reset_ui_state(self):
        self.prog_bar.set(0)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.btn_start.configure(state="normal", text="START BATCH PROCESS")
        if hasattr(self, 'btn_open_folder'):
            self.btn_open_folder.pack_forget()

    def open_output_folder(self):
        try:
            target = os.path.join(os.getcwd(), "GST Downloaded", "GSTR 2B")
            if not os.path.exists(target):
                target = os.getcwd()
            os.startfile(target)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()