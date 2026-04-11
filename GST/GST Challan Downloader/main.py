import threading
import time
import os
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
from webdriver_manager.chrome import ChromeDriverManager

# --- UI CONFIGURATION ---
# Commented out: theme is controlled globally by GST_Suite.py
# ctk.set_appearance_mode("System")
# ctk.set_default_color_theme("blue")

try:
    _CAPTCHA_RESAMPLE = Image.Resampling.NEAREST
except AttributeError:
    _CAPTCHA_RESAMPLE = Image.NEAREST

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

    def run(self):
        self.log("🚀 INITIALIZING GST CHALLAN ENGINE V1...")
        
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
            base_dir = os.path.join(os.getcwd(), "GST_Challans")
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
            filename = f"Challan_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_df.to_excel(filename, index=False)
            self.log(f"📄 Summary Report saved: {filename}")
        except Exception as e:
            self.log(f"⚠️ Failed to save report: {e}")

    def process_single_user(self, username, password, user_root):
        """ Returns (Overall Status, Reason String) """
        try:
            # --- BROWSER SETUP (ANTI-DETECT) ---
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-blink-features=AutomationControlled") 
            options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
            options.add_experimental_option('useAutomationExtension', False)
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")

            prefs = {
                "download.default_directory": user_root,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1
            }
            options.add_experimental_option("prefs", prefs)
            
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            
            # Stealth JS Injection
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"""
            })

            wait = WebDriverWait(self.driver, 20)
            
            # 1. LOGIN & NAVIGATE TO DASHBOARD
            login_status, login_msg = self.perform_login(username, password, wait)
            if not login_status: return "Login Failed", login_msg

            # 2. NAVIGATE TO CHALLAN HISTORY
            challan_status, challan_msg = self.navigate_to_challan_history(wait)
            if not challan_status: return "Navigation Failed", challan_msg

            # 3. DOWNLOAD CHALLANS
            dl_status, dl_msg = self.download_challans(wait, user_root)
            return dl_status, dl_msg

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

                self._save_captcha_image(wait, "temp_captcha.png")
                
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
                    
                    # --- MODAL HANDLER ---
                    time.sleep(3)
                    try:
                        aadhaar_skip = self.driver.find_elements(By.XPATH, "//a[contains(text(),'Remind me later')]")
                        if aadhaar_skip and aadhaar_skip[0].is_displayed():
                            self.log("   ℹ️ Closing Aadhaar Popup...")
                            aadhaar_skip[0].click()
                            time.sleep(1.5)
                    except: pass

                    try:
                        generic_skip = self.driver.find_elements(By.XPATH, "//button[contains(text(),'Remind Me Later')]")
                        if generic_skip and generic_skip[0].is_displayed():
                            self.log("   ℹ️ Closing Generic Popup...")
                            generic_skip[0].click()
                            time.sleep(1.5)
                    except: pass

                    # Already on Dashboard after login — no further navigation needed
                    self.log("   ✅ Already on Dashboard.")
                    return True, "Success"

            except Exception as e:
                self.log(f"   ⚠️ Login Exception: {e}")
                return False, f"Login Error: {str(e)[:20]}"

    def navigate_to_challan_history(self, wait):
        """ Navigates via UI: Services → Payments → Challan History. Returns (status, msg) """
        try:
            # Step 1: Click the top-level "Services" dropdown in navbar
            self.log("   🖱️ Clicking 'Services' menu...")
            services_menu = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//li[contains(@class,'menuList')]//a[contains(text(),'Services') and contains(@class,'dropdown-toggle')]")
            ))
            self.driver.execute_script("arguments[0].click();", services_menu)
            time.sleep(1.5)

            # Step 2: Hover over "Payments" to reveal its sub-menu
            self.log("   🖱️ Hovering over 'Payments' sub-menu...")
            payments_link = wait.until(EC.visibility_of_element_located(
                (By.XPATH, "//a[@data-ng-show=\"udata && udata.role == 'login'\" and contains(@href,'quicklinks/payments') and normalize-space(text())='Payments']")
            ))
            # Use ActionChains to hover so sub-menu opens
            from selenium.webdriver.common.action_chains import ActionChains
            ActionChains(self.driver).move_to_element(payments_link).perform()
            time.sleep(1.5)

            # Step 3: Click "Challan History" inside the Payments sub-menu (logged-in link)
            self.log("   🖱️ Clicking 'Challan History'...")
            challan_history_link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[@data-ng-href='//payment.gst.gov.in/payment/auth/challanhistory' and @data-ng-show and contains(@data-ng-show,'login')]")
            ))
            self.driver.execute_script("arguments[0].click();", challan_history_link)
            time.sleep(3)

            # Step 4: Verify landing page
            current_url = self.driver.current_url
            self.log(f"   📄 Current URL: {current_url}")

            if "challanhistory" in current_url:
                self.log("   ✅ Reached Challan History page!")
                return True, "Success"
            else:
                return False, f"Unexpected page after nav: {current_url[:60]}"

        except Exception as e:
            self.log(f"   ❌ Challan History Navigation Error: {e}")
            return False, f"Nav Error: {str(e)[:60]}"

    def download_challans(self, wait, user_root):
        """
        Reads the Challan History table, filters rows by year and type
        (AOP = Monthly, MPQR = Quarterly), clicks each CPIN, downloads the PDF,
        then goes back. Handles pagination. Returns (status, summary_msg).
        """
        from selenium.webdriver.common.action_chains import ActionChains

        # Determine filter keyword from settings
        challan_type = self.settings.get("type", "Monthly (AOP)")
        filter_keyword = "AOP" if "Monthly" in challan_type else "MPQR"

        # Derive calendar year filter e.g. "2025" → match any row where Created On year == 2025
        selected_year_str = self.settings.get("year", "")
        try:
            selected_year = int(selected_year_str)
        except Exception:
            selected_year = None

        def in_selected_year(date_str):
            """ date_str format: DD/MM/YYYY HH:MM:SS or DD/MM/YYYY """
            if not selected_year or not date_str or date_str.strip() == "-":
                return True  # no year filter
            try:
                parts = date_str.strip().split("/")
                if len(parts) < 3: return True
                row_year = int(parts[2][:4])
                return row_year == selected_year
            except Exception:
                return True

        downloaded = 0
        skipped = 0
        page_num = 1

        self.log(f"   📋 Scanning Challan History [{filter_keyword}] for year {selected_year_str}...")

        while True:
            if not self.keep_running: break

            time.sleep(2)

            # Collect all rows on current page
            try:
                rows = self.driver.find_elements(
                    By.XPATH, "//tbody//tr[td[@data-title-text='CPIN']]")
            except Exception:
                rows = []

            if not rows:
                self.log(f"   ⚠️ No rows found on page {page_num}.")
                break

            self.log(f"   📄 Page {page_num}: {len(rows)} rows found.")

            # Build list of (cpin, created_on, reason) to process
            to_process = []
            for row in rows:
                try:
                    cpin = row.find_element(
                        By.XPATH, ".//td[@data-title-text='CPIN']//span[@data-ng-bind='user.cpin']"
                    ).text.strip()
                    created_on = row.find_element(
                        By.XPATH, ".//td[@data-title-text='Created On']"
                    ).text.strip()
                    reason = row.find_element(
                        By.XPATH, ".//td[@data-title-text='Challan Reason']"
                    ).text.strip()
                except Exception:
                    continue

                # Filter by keyword and year
                if filter_keyword.upper() not in reason.upper():
                    skipped += 1
                    continue
                if not in_selected_year(created_on):
                    skipped += 1
                    continue

                to_process.append(cpin)

            self.log(f"   ✅ {len(to_process)} matching challans on page {page_num} (skipped {skipped}).")

            # Click each CPIN and download
            for cpin in to_process:
                if not self.keep_running: break
                try:
                    self.log(f"   💾 Downloading CPIN: {cpin}...")

                    # Re-find the CPIN anchor (DOM can refresh between iterations)
                    cpin_link = wait.until(EC.element_to_be_clickable(
                        (By.XPATH,
                         f"//td[@data-title-text='CPIN']//a[span[@data-ng-bind='user.cpin' and normalize-space(text())='{cpin}']]"
                        )
                    ))
                    self.driver.execute_script("arguments[0].click();", cpin_link)
                    time.sleep(3)

                    # Click Download button on detail page
                    dl_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(@class,'btn-primary') and contains(text(),'Download')]"
                                   " | //button[@data-ng-click='pdfControllerReceipt(challanData)']"
                                   " | //button[normalize-space(text())='Download']")
                    ))
                    self.driver.execute_script("arguments[0].click();", dl_btn)
                    self.log(f"   ⬇️ Download triggered for {cpin}.")

                    # Wait for PDF to land in user_root (up to 20s)
                    deadline = time.time() + 20
                    while time.time() < deadline:
                        pdfs = glob.glob(os.path.join(user_root, "*.pdf"))
                        crdownloads = glob.glob(os.path.join(user_root, "*.crdownload"))
                        if pdfs and not crdownloads:
                            break
                        time.sleep(1)

                    downloaded += 1

                    # Navigate back to challan history
                    self.driver.back()
                    time.sleep(2)

                    # If we got redirected away from challan history, re-navigate
                    if "challanhistory" not in self.driver.current_url:
                        self.log("   🔄 Re-navigating to Challan History...")
                        nav_ok, _ = self.navigate_to_challan_history(wait)
                        if not nav_ok:
                            return "Partial", f"Downloaded {downloaded}, lost nav after back()"
                        # Re-navigate to the correct page
                        for _ in range(page_num - 1):
                            try:
                                next_btn = self.driver.find_element(
                                    By.XPATH, "//a[contains(@ng-click,'next') or contains(@class,'next')]"
                                )
                                self.driver.execute_script("arguments[0].click();", next_btn)
                                time.sleep(2)
                            except Exception:
                                break

                except Exception as e:
                    self.log(f"   ⚠️ Failed on CPIN {cpin}: {str(e)[:60]}")

            # Check for next page
            try:
                next_btn = self.driver.find_element(
                    By.XPATH,
                    "//a[contains(@ng-click,'next') or (contains(@class,'next') and not(contains(@class,'disabled')))]"
                )
                if "disabled" in (next_btn.get_attribute("class") or ""):
                    break
                self.driver.execute_script("arguments[0].click();", next_btn)
                page_num += 1
                time.sleep(2)
            except Exception:
                break  # No more pages

        if downloaded == 0:
            return "No Challans", f"No {filter_keyword} challans found for year {selected_year_str}"
        return "Success", f"Downloaded {downloaded} challan(s) [{filter_keyword}] for year {selected_year_str}"

# --- GUI CLASS ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GST Challan Downloader - Professional Edition")
        self.geometry("900x850")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.worker = None
        self.excel_file = ""

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1a237e", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False)
        ctk.CTkLabel(self.head, text="GST CHALLAN DOWNLOADER",
                     font=("Roboto Medium", 24, "bold"), text_color="white").pack(side="left", padx=20, pady=10)

        # Theme Selector — removed; theme is controlled by GST_Suite.py
        # self.theme_frame = ctk.CTkFrame(self.head, fg_color="transparent")
        # self.theme_frame.pack(side="right", padx=20, pady=15)
        # ctk.CTkLabel(self.theme_frame, text="Theme:", ...).pack(...)
        # self.theme_option = ctk.CTkOptionMenu(...)
        # self.theme_option.pack(side="left")

        ctk.CTkLabel(self.head, text="CHALLAN AUTOMATION",
                     font=("Roboto", 14), text_color="#bbdefb").pack(side="right", padx=(0, 10), pady=15)

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
        btn_row = ctk.CTkFrame(self.card_cred, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(5, 15))
        self.btn_download = ctk.CTkButton(btn_row, text="📥 Sample Excel", command=self.download_sample, fg_color="#43a047", hover_color="#2e7d32", height=28, font=("Arial", 12, "bold"))
        self.btn_download.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.btn_demo = ctk.CTkButton(btn_row, text="▶ View Demo", command=self.open_demo_link, fg_color="#e53935", hover_color="#b71c1c", height=28, font=("Arial", 12, "bold"))
        self.btn_demo.pack(side="left", expand=True, fill="x", padx=(5, 0))

        # Period Settings Card
        self.card_period = ctk.CTkFrame(self.settings_container, border_color="#3949ab", border_width=1)
        self.card_period.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        ctk.CTkLabel(self.card_period, text="📅 Period & Type", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))

        cur_year = datetime.now().year
        year_list = [str(y) for y in range(cur_year - 5, cur_year + 1)]

        self.frm_year = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_year.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_year, text="Year:", width=110, anchor="w").pack(side="left")
        self.cb_year = ctk.CTkComboBox(self.frm_year, values=year_list, width=150)
        self.cb_year.set(year_list[0])
        self.cb_year.pack(side="right", expand=True, fill="x")

        self.frm_type = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_type.pack(fill="x", padx=15, pady=(6, 15))
        ctk.CTkLabel(self.frm_type, text="Challan Type:", width=110, anchor="w").pack(side="left")
        self.cb_type = ctk.CTkComboBox(self.frm_type,
                                       values=["Monthly (AOP)", "Quarterly (MPQR)"],
                                       width=150)
        self.cb_type.set("Monthly (AOP)")
        self.cb_type.pack(side="right", expand=True, fill="x")

        # LOGS
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
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
        self.cap_stop_btn = ctk.CTkButton(self.cap_inner, text="⏹ STOP PROCESS", fg_color="#424242", hover_color="#212121",
                                          height=35, width=250, font=("Arial", 11, "bold"), command=self.stop_process)
        self.cap_stop_btn.pack(pady=(8, 5))

        # FOOTER
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 20))
        self.prog_bar = ctk.CTkProgressBar(self.footer, height=15, progress_color="#00e676")
        self.prog_bar.pack(fill="x", pady=(0, 10))
        self.prog_bar.set(0)
        self.btn_row_footer = ctk.CTkFrame(self.footer, fg_color="transparent")
        self.btn_row_footer.pack(fill="x")
        self.btn_start = ctk.CTkButton(self.btn_row_footer, text="START BATCH PROCESS", height=50, font=("Arial", 16, "bold"),
                                       fg_color="#2e7d32", hover_color="#1b5e20", command=self.start_process)
        self.btn_start.pack(side="left", expand=True, fill="x")
        self.btn_stop = ctk.CTkButton(self.btn_row_footer, text="⏹ STOP", height=50, font=("Arial", 16, "bold"),
                                      fg_color="#c62828", hover_color="#8e0000", command=self.stop_process, width=150)
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_stop.pack_forget()
    def download_sample(self):
        import shutil
        import os
        from tkinter import messagebox
        sample_path = os.path.join(os.path.dirname(__file__), "GST Challan Downloader Sample File.xlsx")
        if not os.path.exists(sample_path):
            messagebox.showerror("Download Error", f"Sample file not found: {sample_path}")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="GST Challan Downloader Sample File.xlsx", filetypes=[("Excel", "*.xlsx")])
        if save_path:
            try:
                shutil.copy2(sample_path, save_path)
                messagebox.showinfo("Success", f"Sample downloaded to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download: {e}")

    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://www.youtube.com/watch?v=XXXXXXXXXX")

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
        self.after(0, lambda: self.btn_stop.pack_forget())
        self.after(0, lambda: self.btn_stop.configure(state="normal", text="⏹ STOP"))

    def request_captcha_safe(self, img_path):
        def show():
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
            self.cap_btn.configure(state="normal", text="SUBMIT CAPTCHA", fg_color="#d32f2f")
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
        if not self.excel_file:
            messagebox.showerror("Error", "Please select Excel file")
            return
        settings = {
            "year": self.cb_year.get(),
            "type": self.cb_type.get(),  # "Monthly (AOP)" or "Quarterly (MPQR)"
        }
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.worker = GSTWorker(self, self.excel_file, settings)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if self.worker:
            self.worker.keep_running = False
        self.btn_stop.configure(state="disabled", text="STOPPING...")
        self.update_log_safe("🛑 Stop requested — will halt after current user...")

if __name__ == "__main__":
    app = App()
    app.mainloop()