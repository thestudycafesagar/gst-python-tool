
import threading
import time
import os
import random
import glob
import base64
import zipfile
import pandas as pd
import customtkinter as ctk
from PIL import Image
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Shared Stealth Driver Import
import sys
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _ROOT not in sys.path: sys.path.insert(0, _ROOT)
from stealth_driver import create_chrome_driver, build_chrome_options

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

try:
    _CAPTCHA_RESAMPLE = Image.Resampling.NEAREST
except AttributeError:
    _CAPTCHA_RESAMPLE = Image.NEAREST

LOGIN_URL = "https://services.gst.gov.in/services/login"
IMS_DASHBOARD_URL = "https://return.gst.gov.in/imsweb/auth/imsDashboard"


class IMSWorker:
    def __init__(self, app_instance, excel_path, settings, credentials=None):
        self.app = app_instance
        self.excel_path = excel_path
        self.settings = settings
        self.credentials = credentials or []
        self.keep_running = True
        self.driver = None
        self.captcha_response = None
        self.report_data = []

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
        self.log("Initializing IMS Dashboard Downloader...")

        try:
            if self.credentials:
                df = pd.DataFrame(self.credentials)
                user_col, pass_col = "Username", "Password"
                self.log(f"Loaded {len(df)} users from Add ID Password.")
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
                self.log(f"Loaded {len(df)} users from Excel.")

            if df.empty:
                self.app.process_finished_safe("No credentials found to process")
                return

            total = len(df)

            base_dir = os.path.join(os.getcwd(), "GST Downloaded", "IMS")
            if not os.path.exists(base_dir):
                os.makedirs(base_dir, exist_ok=True)

            stopped_by_user = False
            for index, row in df.iterrows():
                if not self.keep_running:
                    stopped_by_user = True
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

                if not self.keep_running:
                    stopped_by_user = True
                    break

                self.log("-" * 40)

            if stopped_by_user or not self.keep_running:
                if self.report_data:
                    self.generate_report()
                    self.log("🛑 Process stopped by user. Partial report saved.")
                    self.app.process_finished_safe("Stopped by user. Partial report saved.")
                else:
                    self.log("🛑 Process stopped by user.")
                    self.app.process_finished_safe("Stopped by user.")
                return

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
            # Create reports subfolder under GST Downloaded/IMS
            base_dir = os.path.join(os.getcwd(), "GST Downloaded", "IMS", "reports")
            os.makedirs(base_dir, exist_ok=True)
            filename = os.path.join(base_dir, f"IMS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            report_df.to_excel(filename, index=False)
            self.log(f"Summary Report saved: {filename}")
        except Exception as e:
            self.log(f"Failed to save report: {e}")

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

    def process_single_user(self, username, password, user_dir):
        try:
            # --- BROWSER SETUP (SHARED STEALTH DRIVER) ---
            self.driver = create_chrome_driver(build_chrome_options(user_dir))

            wait = WebDriverWait(self.driver, 20)

            # --- LOGIN ---
            login_ok, login_msg = self.perform_login(username, password, wait)
            if not login_ok:
                return "Login Failed", login_msg

            self.human_delay()

            from selenium.webdriver.common.action_chains import ActionChains
            actions = ActionChains(self.driver)

            # --- STEP 1: CLICK 'Services' to open the smenu dropdown ---
            self.log("   Clicking 'Services' to open dropdown...")
            time.sleep(random.uniform(1.5, 2.5))
            services_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//a[@data-toggle='dropdown' and contains(normalize-space(.),'Services')]"
                ))
            )
            self.driver.execute_script("arguments[0].click();", services_btn)
            time.sleep(random.uniform(1.5, 2.0))

            # --- STEP 2: HOVER on 'Returns' tab inside smenu to reveal submenu ---
            self.log("   Hovering on 'Returns'...")
            returns_tab = WebDriverWait(self.driver, 8).until(
                EC.visibility_of_element_located((By.XPATH,
                    "//ul[contains(@class,'smenu')]//a[contains(@href,'quicklinks/returns') or normalize-space(text())='Returns']"
                ))
            )
            actions.move_to_element(returns_tab).perform()
            time.sleep(random.uniform(1.2, 1.8))

            # --- STEP 3: CLICK IMS Dashboard from isubmenu using exact href ---
            self.log("   Clicking IMS Dashboard...")
            ims_link = WebDriverWait(self.driver, 8).until(
                EC.visibility_of_element_located((By.XPATH,
                    "//ul[contains(@class,'isubmenu')]//a[contains(@href,'imsDashboard') or contains(@data-ng-href,'imsDashboard')]"
                ))
            )
            actions.move_to_element(ims_link).pause(random.uniform(0.4, 0.8)).perform()
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", ims_link)
            self.log("   ✅ IMS Dashboard clicked.")
            time.sleep(random.uniform(4, 6))

            # --- STEP 4: CLICK 'View' UNDER INWARD SUPPLIES (skip if already there) ---
            if "inwardsupplies" not in self.driver.current_url.lower():
                self.log("   Clicking Inward Supplies > View...")
                try:
                    v_btn = WebDriverWait(self.driver, 8).until(
                        EC.element_to_be_clickable((By.XPATH,
                            "//button[contains(@data-ng-click,'navigateInwsupDashboard') or contains(text(),'VIEW') or contains(text(),'View')]"
                        ))
                    )
                    self.driver.execute_script("arguments[0].click();", v_btn)
                    self.log("   ✅ Clicked View / Inward Supplies.")
                    time.sleep(4)
                except Exception as e:
                    self.log(f"   ⚠️ View button not found, continuing: {e}")
            else:
                self.log("   ✅ Already on Inward Supplies page.")

            # --- STEP 5: DISMISS INFORMATION POPUP (click OKAY if visible) ---
            self.log("   Checking for Information popup...")
            try:
                time.sleep(2)
                okay_xpath = "//div[contains(@class,'modal') and contains(@style,'display: block') or contains(@class,'modal-open')]//button[contains(translate(text(),'okay','OKAY'),'OKAY') or contains(text(),'Ok') or contains(text(),'OK')]"
                okay_buttons = self.driver.find_elements(By.XPATH, okay_xpath)
                popup_closed = False
                for btn in okay_buttons:
                    try:
                        if btn.is_displayed():
                            self.driver.execute_script("arguments[0].click();", btn)
                            popup_closed = True
                            self.log("   Popup closed.")
                            time.sleep(2)
                            break
                    except Exception:
                        pass
                if not popup_closed:
                    self.log("   No popup found, proceeding to download...")
            except Exception as e:
                self.log(f"   Popup check error: {e}")

            # --- STEP 6: CLICK DOWNLOAD IMS DETAILS (EXCEL) ---
            self.log("   Clicking DOWNLOAD IMS DETAILS (EXCEL)...")
            try:
                dl_btn = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH,
                        "//button[@data-ng-click='downloadIMSSummary()' or contains(text(),'DOWNLOAD IMS DETAILS')]"
                    ))
                )
                self.driver.execute_script("arguments[0].scrollIntoView(true);", dl_btn)
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click();", dl_btn)
                self.log("   ✅ Download button clicked.")
            except Exception as e:
                self.log(f"   ❌ Download button not found: {e}")
                return "Failed", "Download button not found"
            time.sleep(5)

            # --- CHECK: File generation in progress (20-min wait message) ---
            try:
                gen_msgs = self.driver.find_elements(By.XPATH,
                    "//div[contains(@class,'alert-success') and contains(.,'generation is in progress')]"
                )
                if gen_msgs and any(m.is_displayed() for m in gen_msgs):
                    retry_time = datetime.now().replace(second=0, microsecond=0)
                    retry_time = retry_time + timedelta(minutes=20)
                    retry_str = retry_time.strftime("%I:%M %p")
                    msg = (
                        f"File generation is in progress for {username}.\n\n"
                        f"The GST portal needs up to 20 minutes to prepare the file.\n\n"
                        f"⏰  Please try again after:  {retry_str}"
                    )
                    self.log(f"   ⏳ File generation in progress. Retry after {retry_str}")
                    self.app.after(0, lambda m=msg: messagebox.showinfo("File Generating — Retry Later", m))
                    return "Retry Later", f"File generating. Try again after {retry_str}"
            except Exception as e:
                self.log(f"   Generation check error: {e}")

            # --- STEP 7: CLICK ALL GENERATED DOWNLOAD LINKS (file 1, file 2, ...) ---
            self.log("   Waiting for download link(s) to appear (up to 2 mins)...")
            try:
                long_wait = WebDriverWait(self.driver, 120)
                # Wait until at least one download link appears
                long_wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[@ng-repeat and contains(@href,'imsExcel')]")
                ))
                time.sleep(2)
                # Grab all download links (file 1, file 2, ...)
                file_links = self.driver.find_elements(By.XPATH,
                    "//a[@ng-repeat and contains(@href,'imsExcel')]"
                )
                self.log(f"   Found {len(file_links)} download link(s).")
                for i, lnk in enumerate(file_links, 1):
                    self.driver.execute_script("arguments[0].click();", lnk)
                    self.log(f"   ✅ Clicked download file {i}.")
                    time.sleep(3)
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
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None

    def perform_login(self, username, password, wait):
        self.log("   🚀 Opening GST Portal login page...")
        self.driver.maximize_window()
        self.driver.get(LOGIN_URL)
        time.sleep(3)

        # Auto-fill Username
        try:
            user_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@id='username' or @name='username' or @placeholder='Username' or @type='text']"))
            )
            self.type_like_human(user_field, username)
            self.log(f"   ✅ Username filled: {username}")
        except Exception as e:
            self.log(f"   ⚠️ Could not auto-fill username: {e}")

        # Wait for password field to appear (GST portal shows it after username is typed)
        time.sleep(2)

        # Auto-fill Password — target the visible field by id="user_pass" (not the hidden duplicate)
        try:
            pass_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "user_pass"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", pass_field)
            self.driver.execute_script("arguments[0].focus();", pass_field)
            time.sleep(0.5)
            # Use JS to set value then trigger Angular ng-model update
            self.driver.execute_script(
                "arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                pass_field, password
            )
            self.log("   ✅ Password filled.")
        except Exception as e:
            self.log(f"   ⚠️ Could not auto-fill password: {e}")

        self.log("   👉 Please enter the CAPTCHA in the Chrome window and click LOGIN.")
        self.log("   ⏳ Waiting for you to complete login... (no time limit)")

        # Wait for user to complete captcha & login
        # Only detect AFTER leaving the login page URL
        while self.keep_running:
            try:
                url = self.driver.current_url.lower()
                src = (self.driver.page_source or "").lower()
                
                # Must have left the login page
                if "login" in url or "services/login" in url:
                    time.sleep(2)
                    continue

                # Strictly check for post-login indicators
                is_logged_in = any(k in url for k in ("dashboard", "auth/home", "services/auth", "fowelcome")) or \
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

        # Handle popups
        try:
            aadhaar_skip = self.driver.find_elements(By.XPATH, "//a[contains(text(),'Remind me later')]")
            if aadhaar_skip and aadhaar_skip[0].is_displayed():
                aadhaar_skip[0].click()
        except: pass

        try:
            generic_skip = self.driver.find_elements(By.XPATH, "//button[contains(text(),'Remind Me Later')]")
            if generic_skip and generic_skip[0].is_displayed():
                generic_skip[0].click()
        except: pass

        self.log("   ✅ Login complete. Proceeding to IMS...")
        return True, "Success"


# ─── GUI ──────────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def download_sample(self):
        # TODO: Implement actual sample download logic
        messagebox.showinfo("Download Sample", "Sample download not implemented yet.")

    def __init__(self):
        super().__init__()
        self.title("IMS Dashboard Downloader")
        self.geometry("800x700")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.worker = None
        self.excel_file = ""
        self.manual_credentials = []

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1D4ED8", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False)
        ctk.CTkLabel(
            self.head, text="IMS DASHBOARD DOWNLOADER",
            font=("Segoe UI", 22, "bold"), text_color="white"
        ).pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(
            self.head, text="INWARD SUPPLIES AUTOMATION",
            font=("Segoe UI", 14), text_color="#CBD5E1"
        ).pack(side="right", padx=20, pady=15)

        # CREDENTIALS CARD
        self.settings_container = ctk.CTkFrame(self, fg_color="transparent")
        self.settings_container.grid(row=1, column=0, sticky="ew", padx=20, pady=(20, 10))

        card = ctk.CTkFrame(self.settings_container, border_color="#334155", border_width=1)
        card.pack(fill="x")
        ctk.CTkLabel(card, text="Credentials Source", font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=15, pady=(15, 5)
        )
        self.ent_file = ctk.CTkEntry(
            card, placeholder_text="Add ID/Password manually (optional)...", height=35
        )
        self.ent_file.pack(fill="x", padx=15, pady=(5, 5))
        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(5, 15))
        self.btn_download = ctk.CTkButton(btn_row, text="➕ Add ID Password", command=self.add_id_password, fg_color="#059669", hover_color="#047857", height=28, font=("Segoe UI", 12, "bold"))
        self.btn_download.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.btn_demo = ctk.CTkButton(btn_row, text="▶ View Demo", command=self.open_demo_link, fg_color="#DC2626", hover_color="#B91C1C", height=28, font=("Segoe UI", 12, "bold"))
        self.btn_demo.pack(side="left", expand=True, fill="x", padx=(5, 0))
        manage_row = ctk.CTkFrame(card, fg_color="transparent")
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

        # LOG BOX (row=2, expands)
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(10, 0))
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self.log_frame, text="Execution Logs", font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=5
        )
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), text_color="#10B981", height=80)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.log_box.configure(state="disabled")

        # FOOTER (row=3, fixed — always visible)
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=3, column=0, sticky="sew", padx=20, pady=(8, 16))

        self.btn_row_footer = ctk.CTkFrame(self.footer, fg_color="transparent")
        self.btn_row_footer.pack(fill="x", pady=(0, 8))

        self.prog_bar = ctk.CTkProgressBar(self.footer, height=15, progress_color="#10B981")
        self.prog_bar.pack(fill="x")
        self.prog_bar.set(0)

        self.btn_start = ctk.CTkButton(
            self.btn_row_footer,
            text="START IMS DOWNLOAD",
            height=50,
            font=("Segoe UI", 16, "bold"),
            fg_color="#059669",
            hover_color="#047857",
            text_color="white",
            command=self.start_process
        )
        self.btn_start.pack(side="left", expand=True, fill="x")

        self.btn_stop = ctk.CTkButton(
            self.btn_row_footer,
            text="⏹ STOP",
            height=50,
            font=("Segoe UI", 16, "bold"),
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=self.stop_process,
            width=150
        )
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_stop.pack_forget()

        self.btn_open_folder = ctk.CTkButton(self.btn_row_footer, text="📂 OPEN FOLDER", height=50, font=("Segoe UI", 16, "bold"),
                                      fg_color="#2563EB", hover_color="#1D4ED8", command=self.open_output_folder, width=180)
        self.btn_open_folder.pack(side="left", padx=(10, 0))
        self.btn_open_folder.pack_forget()

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
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
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
            messagebox.showinfo("Done", msg)
            is_stopped = "stopped" in (msg or "").lower()
            self.btn_start.configure(state="normal", text="STOPPED" if is_stopped else "START IMS DOWNLOAD")
            self.btn_stop.pack_forget()
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            if is_stopped:
                self.after(1200, lambda: self.btn_start.configure(text="START IMS DOWNLOAD"))
            else:
                self.btn_open_folder.pack(side="left", padx=(10, 0))
        self.after(0, _finish_ui)

    def open_output_folder(self):
        try:
            target = os.path.join(os.getcwd(), "GST Downloaded", "IMS")
            if not os.path.exists(target):
                target = os.getcwd()
            os.startfile(target)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {e}")

    def clear_captcha_safe(self):
        pass

    def submit_captcha(self):
        pass

    def start_process(self):
        credentials = list(self.manual_credentials)
        if not credentials and not self.excel_file:
            messagebox.showerror("Error", "Please add ID/Password first")
            return
        settings = {
            "all_quarters": False,
            "manual_login": True
        }
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_open_folder.pack_forget()
        self.worker = IMSWorker(self, self.excel_file, settings, credentials=credentials)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if not self.worker:
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

        self.btn_stop.configure(state="disabled", text="STOPPED")
        self.update_log_safe("🛑 Process stopped by user.")


if __name__ == "__main__":
    app = App()
    app.mainloop()