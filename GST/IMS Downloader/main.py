import threading
import time
import os
import glob
import base64
import zipfile
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
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

try:
    _CAPTCHA_RESAMPLE = Image.Resampling.NEAREST
except AttributeError:
    _CAPTCHA_RESAMPLE = Image.NEAREST

LOGIN_URL = "https://services.gst.gov.in/services/login"
IMS_DASHBOARD_URL = "https://return.gst.gov.in/imsweb/auth/imsDashboard"


class IMSWorker:
    def __init__(self, app_instance, excel_path):
        self.app = app_instance
        self.excel_path = excel_path
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
        self.log("Initializing IMS Dashboard Downloader...")

        try:
            df = pd.read_excel(self.excel_path)
            clean_cols = {c.lower().strip(): c for c in df.columns}
            user_col = next((clean_cols[c] for c in clean_cols if 'user' in c or 'name' in c), None)
            pass_col = next((clean_cols[c] for c in clean_cols if 'pass' in c or 'pwd' in c), None)

            if not user_col or not pass_col:
                self.app.process_finished_safe("Column Error: Need Username/Password columns in Excel")
                return

            total = len(df)
            self.log(f"Loaded {total} users.")

            base_dir = os.path.join(os.getcwd(), "IMS_Downloads")
            if not os.path.exists(base_dir):
                os.makedirs(base_dir)

            for index, row in df.iterrows():
                if not self.keep_running:
                    break

                username = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()

                self.app.update_progress_safe(index / total)
                self.log(f"\nProcessing: {username}")

                # Unique folder per user
                user_dir_base = os.path.join(base_dir, username)
                user_dir = user_dir_base
                counter = 1
                while os.path.exists(user_dir):
                    user_dir = f"{user_dir_base}_{counter}"
                    counter += 1
                os.makedirs(user_dir)

                status, reason = self.process_single_user(username, password, user_dir)

                self.report_data.append({
                    "Username": username,
                    "Status": status,
                    "Details": reason,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Saved To": os.path.basename(user_dir)
                })

                self.log("-" * 40)

            self.generate_report()
            self.app.update_progress_safe(1.0)
            self.log("ALL TASKS COMPLETED.")
            self.app.process_finished_safe("Batch Completed & Report Saved.")

        except Exception as e:
            self.log(f"Critical Error: {e}")
            self.app.process_finished_safe("Error Occurred")

    def generate_report(self):
        try:
            if not self.report_data:
                return
            report_df = pd.DataFrame(self.report_data)
            filename = f"IMS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_df.to_excel(filename, index=False)
            self.log(f"Summary Report saved: {filename}")
        except Exception as e:
            self.log(f"Failed to save report: {e}")

    def process_single_user(self, username, password, user_dir):
        try:
            # --- BROWSER SETUP (ANTI-DETECT) ---
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            options.add_argument(
                "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
            )
            prefs = {
                "download.prompt_for_download": False,
                "directory_upgrade": True,
                "safebrowsing.enabled": True,
                "download.default_directory": user_dir,
                "profile.default_content_setting_values.automatic_downloads": 1,
            }
            options.add_experimental_option("prefs", prefs)

            self.driver = webdriver.Chrome(options=options)

            # Stealth JS Injection
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"""
            })
            self.driver.execute_cdp_cmd(
                "Page.setDownloadBehavior",
                {"behavior": "allow", "downloadPath": user_dir},
            )

            wait = WebDriverWait(self.driver, 20)

            # --- LOGIN ---
            login_ok, login_msg = self.perform_login(username, password, wait)
            if not login_ok:
                return "Login Failed", login_msg

            # --- STEP 1: CLICK 'Services' IN NAVBAR ---
            self.log("   Clicking 'Services' in navbar...")
            try:
                services_btn = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(text(), 'Services') or contains(normalize-space(),'Services')]")
                ))
                self.driver.execute_script("arguments[0].click();", services_btn)
                time.sleep(2)
            except Exception as e:
                self.log(f"   Could not click Services: {e}")
                return "Failed", "Services navbar not found"

            # --- STEP 2: CLICK 'Returns' IN NAVBAR ---
            self.log("   Clicking 'Returns' in navbar...")
            try:
                returns_btn = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(text(), 'Returns')]")
                ))
                self.driver.execute_script("arguments[0].click();", returns_btn)
                time.sleep(2)
            except Exception as e:
                self.log(f"   Could not click Returns: {e}")
                return "Failed", "Returns navbar not found"

            # --- STEP 3: CLICK 'Invoice Management System (IMS) Dashboard' LINK ---
            self.log("   Clicking IMS Dashboard link...")
            try:
                ims_link = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(@href,'imsDashboard')]")
                ))
                self.driver.execute_script("arguments[0].click();", ims_link)
                time.sleep(5)
            except Exception as e:
                self.log(f"   Could not click IMS Dashboard link: {e}")
                return "Failed", "IMS Dashboard link not found"

            # --- STEP 4: CLICK 'View' UNDER INWARD SUPPLIES ---
            self.log("   Clicking Inward Supplies > View...")
            try:
                view_btn = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//button[contains(@data-ng-click,'navigateInwsupDashboard')]")
                ))
                self.driver.execute_script("arguments[0].click();", view_btn)
                time.sleep(5)
            except Exception as e:
                self.log(f"   Could not click Inward Supplies View: {e}")
                return "Failed", "Inward Supplies View button not found"

            # --- STEP 5: DISMISS INFORMATION POPUP (click OKAY) ---
            self.log("   Handling Information popup...")
            try:
                # Wait for popup animation to complete
                time.sleep(2)
                
                # Retrieve all buttons matching "Okay" or "OKAY" inside modals
                okay_xpath = "//div[contains(@class, 'modal-content')]//button[contains(translate(text(), 'OKAY', 'okay'), 'okay')]"
                okay_buttons = self.driver.find_elements(By.XPATH, okay_xpath)
                
                popup_closed = False
                for btn in okay_buttons:
                    if btn.is_displayed():  # Crucial fix: Only interact if the button is physically visible
                        self.driver.execute_script("arguments[0].click();", btn)
                        popup_closed = True
                        self.log("   Popup closed successfully.")
                        time.sleep(2)
                        break # Stop after clicking the visible one
                        
                if not popup_closed:
                    self.log("   No visible popup found, moving to download...")
                    
            except Exception as e:
                self.log(f"   Popup error: {e}")

            # --- STEP 6: CLICK DOWNLOAD IMS DETAILS (EXCEL) ---
            self.log("   Clicking DOWNLOAD IMS DETAILS (EXCEL)...")
            try:
                dl_btn = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//button[contains(@data-ng-click,'downloadIMSSummary')] | //button[contains(text(), 'DOWNLOAD IMS DETAILS (EXCEL)')]")
                ))
                self.driver.execute_script("arguments[0].click();", dl_btn)
                time.sleep(5)
            except Exception as e:
                self.log(f"   Download button error: {e}")
                return "Failed", "Download button not found"

            # --- STEP 7: CLICK THE GENERATED DOWNLOAD LINK ---
            self.log("   Waiting for the file generation link to appear (this may take up to 2 mins)...")
            try:
                # We use a longer wait here because GST server processing takes time
                long_wait = WebDriverWait(self.driver, 120)
                file_link = long_wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(@href, 'imsExcel') or contains(., 'download file')]")
                ))
                self.driver.execute_script("arguments[0].click();", file_link)
                self.log("   Clicked the generated download link.")
                time.sleep(5)
            except Exception as e:
                self.log(f"   Generated link error: {e}")
                return "Failed", "Generated file link did not appear"

            # --- WAIT FOR FILE ---
            self.log("   Waiting for file to save to local drive...")
            file_downloaded = False
            downloaded_file_path = None
            
            for _ in range(60): # Increased wait loop for file download to 60s
                time.sleep(1)
                # Portal downloads a .zip containing the excel file
                files = glob.glob(os.path.join(user_dir, "*.zip")) + \
                        glob.glob(os.path.join(user_dir, "*.xlsx")) + \
                        glob.glob(os.path.join(user_dir, "*.xls"))
                
                # Filter out Chrome's temporary download files
                valid_files = [f for f in files if not f.endswith('.crdownload')]
                
                if valid_files:
                    latest = max(valid_files, key=os.path.getctime)
                    if (datetime.now().timestamp() - os.path.getctime(latest)) < 60:
                        self.log(f"   Downloaded: {os.path.basename(latest)}")
                        file_downloaded = True
                        downloaded_file_path = latest
                        break

            if not file_downloaded:
                return "Failed", "Download timeout (no zip/excel file appeared within 60s)"

            # --- EXTRACT ZIP IF NEEDED ---
            if downloaded_file_path and downloaded_file_path.endswith('.zip'):
                self.log("   Extracting ZIP file...")
                try:
                    with zipfile.ZipFile(downloaded_file_path, 'r') as zip_ref:
                        zip_ref.extractall(user_dir)
                    self.log("   Extraction complete.")
                except Exception as e:
                    self.log(f"   Extraction failed: {e}")
                    return "Partial Success", f"Downloaded zip, but extraction failed: {str(e)[:30]}"

            return "Success", f"File saved and extracted in {os.path.basename(user_dir)}"

        except Exception as e:
            return "Error", f"Browser error: {str(e)[:80]}"
        finally:
            if self.driver:
                self.driver.quit()

    def perform_login(self, username, password, wait):
        self.log("   Opening GST Portal...")
        self.driver.get(LOGIN_URL)

        while True:
            if not self.keep_running:
                return False, "Stopped"
            try:
                wait.until(EC.visibility_of_element_located((By.ID, "username"))).clear()
                self.driver.find_element(By.ID, "username").send_keys(username)

                self.driver.find_element(By.ID, "user_pass").clear()
                self.driver.find_element(By.ID, "user_pass").send_keys(password)

                self._save_captcha_image(wait, "temp_captcha.png")

                self.log("   Waiting for Captcha...")
                self.captcha_response = None
                self.captcha_event.clear()
                self.app.request_captcha_safe("temp_captcha.png")
                self.captcha_event.wait()

                if not self.captcha_response:
                    return False, "Captcha Cancelled"

                self.driver.find_element(By.ID, "captcha").clear()
                self.driver.find_element(By.ID, "captcha").send_keys(self.captcha_response)
                self.driver.find_element(By.XPATH, "//button[@type='submit']").click()

                time.sleep(3)

                src = self.driver.page_source
                if "Invalid Username or Password" in src:
                    self.log("   Bad Credentials.")
                    return False, "Invalid Credentials"

                if "Enter valid Letters" in src or "Invalid Captcha" in src:
                    self.log("   Invalid Captcha. Retrying...")
                    time.sleep(1)
                    continue

                if "Dashboard" in self.driver.title or "Return Dashboard" in src or "fowelcome" in self.driver.current_url:
                    self.log("   Login Successful!")
                    self.app.close_captcha_safe()

                    # --- POPUP HANDLER ---
                    time.sleep(3)
                    try:
                        aadhaar_skip = self.driver.find_elements(By.XPATH, "//a[contains(text(),'Remind me later')]")
                        if aadhaar_skip and aadhaar_skip[0].is_displayed():
                            self.log("   Closing Aadhaar Popup...")
                            aadhaar_skip[0].click()
                            time.sleep(1.5)
                    except:
                        pass

                    try:
                        generic_skip = self.driver.find_elements(By.XPATH, "//button[contains(text(),'Remind Me Later')]")
                        if generic_skip and generic_skip[0].is_displayed():
                            self.log("   Closing Generic Popup...")
                            generic_skip[0].click()
                            time.sleep(1.5)
                    except:
                        pass

                    # Click Return Dashboard to establish session on return.gst.gov.in
                    try:
                        dash_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Return Dashboard')]")))
                        self.driver.execute_script("arguments[0].click();", dash_btn)
                        time.sleep(3)
                        return True, "Success"
                    except:
                        self.log("   Dashboard Nav Error.")
                        return False, "Dashboard Nav Failed"

            except Exception as e:
                self.log(f"   Login Exception: {e}")
                return False, f"Login Error: {str(e)[:30]}"


# ─── GUI ──────────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("IMS Dashboard Downloader")
        self.geometry("800x700")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.worker = None
        self.excel_file = ""

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1a237e", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False)
        ctk.CTkLabel(
            self.head, text="IMS DASHBOARD DOWNLOADER",
            font=("Roboto Medium", 22, "bold"), text_color="white"
        ).pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(
            self.head, text="INWARD SUPPLIES AUTOMATION",
            font=("Roboto", 14), text_color="#bbdefb"
        ).pack(side="right", padx=20, pady=15)

        # CREDENTIALS CARD
        self.settings_container = ctk.CTkFrame(self, fg_color="transparent")
        self.settings_container.grid(row=1, column=0, sticky="ew", padx=20, pady=(20, 10))

        card = ctk.CTkFrame(self.settings_container, border_color="#3949ab", border_width=1)
        card.pack(fill="x")
        ctk.CTkLabel(card, text="Credentials Source (Excel)", font=("Arial", 14, "bold")).pack(
            anchor="w", padx=15, pady=(15, 5)
        )
        self.ent_file = ctk.CTkEntry(
            card, placeholder_text="Select Excel file with Username / Password columns...", height=35
        )
        self.ent_file.pack(fill="x", padx=15, pady=(5, 5))
        ctk.CTkButton(
            card, text="Browse File", command=self.browse_file,
            fg_color="#3949ab", hover_color="#283593", height=35
        ).pack(fill="x", padx=15, pady=(0, 5))
        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(5, 15))
        self.btn_download = ctk.CTkButton(btn_row, text="📥 Sample Excel", command=self.download_sample, fg_color="#43a047", hover_color="#2e7d32", height=28, font=("Arial", 12, "bold"))
        self.btn_download.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.btn_demo = ctk.CTkButton(btn_row, text="▶ View Demo", command=self.open_demo_link, fg_color="#e53935", hover_color="#b71c1c", height=28, font=("Arial", 12, "bold"))
        self.btn_demo.pack(side="left", expand=True, fill="x", padx=(5, 0))

        # LOG BOX
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self.log_frame, text="Execution Logs", font=("Arial", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=5
        )
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), text_color="#00e676", height=200)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.log_box.configure(state="disabled")

        # CAPTCHA PANEL
        self.cap_frame = ctk.CTkFrame(self, border_color="#d32f2f", border_width=2, fg_color="#2b2b2b")
        cap_inner = ctk.CTkFrame(self.cap_frame, fg_color="transparent")
        cap_inner.pack(fill="both", padx=20, pady=10)
        ctk.CTkLabel(
            cap_inner, text="CAPTCHA ACTION REQUIRED",
            text_color="#ef5350", font=("Arial", 14, "bold")
        ).pack()
        self.cap_lbl_img = ctk.CTkLabel(cap_inner, text="")
        self.cap_lbl_img.pack(pady=10)
        self.cap_ent = ctk.CTkEntry(
            cap_inner, placeholder_text="Type Captcha Here...",
            font=("Consolas", 20), justify="center", height=45, width=250
        )
        self.cap_ent.pack(pady=5)
        self.cap_ent.bind("<Return>", self.submit_captcha)
        self.cap_btn = ctk.CTkButton(
            cap_inner, text="SUBMIT CAPTCHA",
            fg_color="#d32f2f", hover_color="#b71c1c",
            height=40, width=250, font=("Arial", 12, "bold"),
            command=self.submit_captcha
        )
        self.cap_btn.pack(pady=(10, 0))
        self.cap_stop_btn = ctk.CTkButton(
            cap_inner, text="⏹ STOP PROCESS",
            fg_color="#424242", hover_color="#212121",
            height=35, width=250, font=("Arial", 11, "bold"),
            command=self.stop_process
        )
        self.cap_stop_btn.pack(pady=(8, 5))

        # FOOTER
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 20))
        self.prog_bar = ctk.CTkProgressBar(self.footer, height=15, progress_color="#00e676")
        self.prog_bar.pack(fill="x", pady=(0, 10))
        self.prog_bar.set(0)
        self.btn_row_footer = ctk.CTkFrame(self.footer, fg_color="transparent")
        self.btn_row_footer.pack(fill="x")
        self.btn_start = ctk.CTkButton(self.btn_row_footer, text="START IMS DOWNLOAD", height=50, font=("Arial", 16, "bold"),
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
        sample_path = os.path.join(os.path.dirname(__file__), "IMS Sample File.xlsx")
        if not os.path.exists(sample_path):
            messagebox.showerror("Download Error", f"Sample file not found: {sample_path}")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="IMS Sample File.xlsx", filetypes=[("Excel", "*.xlsx")])
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
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
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
        self.after(0, lambda: messagebox.showinfo("Done", msg))
        self.after(0, lambda: self.btn_start.configure(state="normal", text="START IMS DOWNLOAD"))
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
        if not txt:
            return
        self.cap_btn.configure(state="disabled", text="VERIFYING...", fg_color="gray")
        self.worker.captcha_response = txt
        self.worker.captcha_event.set()

    def close_captcha_safe(self):
        self.after(0, lambda: self.cap_frame.grid_forget())

    def start_process(self):
        if not self.excel_file:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.worker = IMSWorker(self, self.excel_file)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if self.worker:
            self.worker.keep_running = False
        self.btn_stop.configure(state="disabled", text="STOPPING...")
        self.update_log_safe("🛑 Stop requested — will halt after current user...")


if __name__ == "__main__":
    app = App()
    app.mainloop()