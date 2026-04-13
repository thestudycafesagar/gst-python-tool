import threading
import time
import os
import random
import glob
import base64
import re
import zipfile
import shutil
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
        
        # Reporting
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
        mode = self.settings['action_mode']
        self.log(f"🚀 INITIALIZING GST ENGINE V25 (Selection Fix + Unzip)...")
        
        try:
            # 1. LOAD CREDENTIALS (manual IDs preferred, Excel optional)
            if self.credentials:
                df = pd.DataFrame(self.credentials)
                user_col, pass_col = "Username", "Password"
                self.log(f"📊 Loaded {len(df)} users from Add ID Password.")
            else:
                if not self.excel_path:
                    self.app.process_finished_safe("Please add ID/Password or select Excel file")
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
            folder_name = "GST_Requests_R1" if "Request" in mode else "GST_Downloads_R1"
            base_dir = os.path.join(os.getcwd(), folder_name)
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
                
                # Unique Folder Versioning (Per User)
                user_root = os.path.join(base_dir, username)
                if not os.path.exists(user_root): os.makedirs(user_root)

                status, reason = self.process_single_user(username, password, user_root)
                
                self.report_data.append({
                    "Username": username,
                    "Status": status,
                    "Details": reason,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Mode": mode
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
            filename = f"GST_Report_R1_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_df.to_excel(filename, index=False)
            self.log(f"📄 Summary Report saved: {filename}")
        except Exception as e:
            self.log(f"⚠️ Failed to save report: {e}")

    def process_single_user(self, username, password, user_root):
        """ Returns (Overall Status, Reason String) """
        try:
            # --- BROWSER SETUP ---
            options = webdriver.ChromeOptions()
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

            selected_q = self.settings['quarter']
            selected_m = self.settings['month']
            if selected_q not in q_map or selected_m not in q_map[selected_q]:
                return "Config Error", "Invalid Month/Quarter Selection"

            tasks = [{"q": selected_q, "m": selected_m}]
            self.log(f"   📅 Queued monthly download for {selected_m}.")
            
            # 3. EXECUTE LOOP
            self.human_delay()
            success_count = 0
            results = []

            # Set Download Path
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
                
                self.log(f"   ⚙️ Processing: {m_text} ({q_text[:9]})")
                
                # RETRY LOOP (3 Attempts)
                for attempt in range(3):
                    try:
                        # 1. Selection
                        year_el = wait.until(EC.element_to_be_clickable((By.NAME, "fin")))
                        Select(year_el).select_by_visible_text(fin_year)
                        time.sleep(0.5)

                        qtr_el = self.driver.find_element(By.NAME, "quarter")
                        Select(qtr_el).select_by_visible_text(q_text)
                        time.sleep(0.5)

                        mon_el = self.driver.find_element(By.NAME, "mon")
                        Select(mon_el).select_by_visible_text(m_text)
                        time.sleep(0.5)

                        search_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Search')]")
                        self.driver.execute_script("arguments[0].click();", search_btn)
                        time.sleep(3) 

                        # 2. Process GSTR-1
                        r1_status, r1_msg = self.process_gstr1(wait, year_folder)
                        
                        if r1_status:
                            success_count += 1
                            results.append(f"{m_text}: ✅ {r1_msg}")
                            self.reset_to_dashboard(wait)
                            break 
                        else:
                            # Soft failures (Not Ready/Generated) don't need retry
                            if "Not Ready" in r1_msg or "Request Submitted" in r1_msg or "Requested" in r1_msg:
                                success_count += 1
                                results.append(f"{m_text}: ✅ {r1_msg}")
                                self.reset_to_dashboard(wait)
                                break
                            else:
                                raise Exception(r1_msg) # Trigger retry

                    except Exception as e:
                        self.log(f"     ⚠️ Attempt {attempt+1}/3 failed: {str(e)[:20]}...")
                        try: self.reset_to_dashboard(wait)
                        except: 
                            self.driver.refresh()
                            time.sleep(5)
                        
                        if attempt == 2: 
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

    def reset_to_dashboard(self, wait):
        try:
            db_btn = self.driver.find_element(By.XPATH, "//button[contains(., 'Return Dashboard')]")
            self.driver.execute_script("arguments[0].click();", db_btn)
        except:
            try: self.driver.back()
            except: pass
        time.sleep(4) 

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
                    
                    self.human_delay()
                    try:
                        aadhaar_skip = self.driver.find_elements(By.XPATH, "//a[contains(text(),'Remind me later')]")
                        if aadhaar_skip and aadhaar_skip[0].is_displayed():
                            aadhaar_skip[0].click()
                            self.human_delay()
                    except: pass
                    try:
                        generic_skip = self.driver.find_elements(By.XPATH, "//button[contains(text(),'Remind Me Later')]")
                        if generic_skip and generic_skip[0].is_displayed():
                            generic_skip[0].click()
                            self.human_delay()
                    except: pass
                    
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

    def process_gstr1(self, wait, download_path):
        """ Handles Request vs Download logic for GSTR-1 """
        mode = self.settings['action_mode']
        
        # 1. Locate GSTR-1 Tile
        xpath_r1_tile = "//p[contains(text(),'GSTR1')]/ancestor::div[contains(@class,'col-sm-4')]//button[contains(text(),'Download')]"
        
        try:
            tile_btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_r1_tile)))
            self.log("     ✅ Found GSTR-1 Tile. Clicking Download...")
            self.driver.execute_script("arguments[0].click();", tile_btn)
            time.sleep(4)
        except:
            self.log("     ⚠️ GSTR-1 Tile Not Found or Disabled.")
            return False, "Tile Missing"

        if mode == "Request JSON":
            try:
                gen_btn_xpath = "//button[contains(text(), 'Generate JSON File to Download')]"
                gen_btn = wait.until(EC.element_to_be_clickable((By.XPATH, gen_btn_xpath)))
                self.log("     ⚙️ Clicking 'Generate JSON File' (Request)...")
                self.driver.execute_script("arguments[0].click();", gen_btn)
                time.sleep(2)
                self.log("     📤 Request Submitted.")
                return True, "Requested"
            except Exception as e:
                self.log(f"     ⚠️ Request Failed: {str(e)[:20]}")
                return False, "Request Error"

        elif mode == "Download JSON":
            try:
                # 1. Click Generate FIRST
                try:
                    gen_btn_xpath = "//button[contains(text(), 'Generate JSON File to Download')]"
                    gen_btn = wait.until(EC.element_to_be_clickable((By.XPATH, gen_btn_xpath)))
                    self.log("     ⚙️ Clicking 'Generate JSON File'...")
                    self.driver.execute_script("arguments[0].click();", gen_btn)
                    self.log("     ⏳ Waiting 5s for link...")
                    time.sleep(5)
                except:
                    self.log("     ℹ️ Generate button skipped/missing...")

                # 3. Look for Download Link
                dwn_link_xpath = "//span[contains(text(), 'Click here to download - File 1')]/parent::a | //a[contains(., 'Click here to download - File 1')]"
                
                try:
                    link = self.driver.find_element(By.XPATH, dwn_link_xpath)
                    if link.is_displayed():
                        # --- SNAPSHOT BEFORE CLICK ---
                        before_files = set(os.listdir(download_path))
                        
                        self.log("     ⬇️ Link Found! Downloading & Unzipping...")
                        self.driver.execute_script("arguments[0].click();", link)
                        
                        if self.wait_for_download_and_extract(download_path, before_files):
                            return True, "Downloaded & Extracted"
                        else:
                            return False, "Timeout"
                    else:
                        raise Exception("Hidden")
                except:
                    msg = "File not ready. Please wait 20 mins."
                    self.log(f"     ⚠️ {msg}")
                    return True, "Not Ready" 

            except Exception as e:
                self.log(f"     ⚠️ Download Error: {str(e)[:20]}")
                return False, "Script Error"
        
        return False, "Invalid Mode"

    def wait_for_download_and_extract(self, download_path, before_files):
        """ Waits for ZIP, extracts it, handles renaming (01, 02), cleans up ZIP """
        self.log("     ⏳ Waiting for new file...")
        
        target_zip = None
        
        # 1. DETECT NEW ZIP
        for _ in range(60): # Wait 60s max
            time.sleep(1)
            current_files = set(os.listdir(download_path))
            new_files = current_files - before_files
            
            for f in new_files:
                if f.endswith(".zip") and not f.endswith(".crdownload") and not f.endswith(".tmp"):
                    target_zip = os.path.join(download_path, f)
                    break
            if target_zip:
                break
        
        if not target_zip:
            return False

        # 2. EXTRACT
        try:
            time.sleep(1) 
            
            temp_extract_folder = os.path.join(download_path, "temp_extract_zone")
            if not os.path.exists(temp_extract_folder): os.makedirs(temp_extract_folder)
            
            with zipfile.ZipFile(target_zip, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_folder)
            
            # 3. MOVE & RENAME JSONs
            extracted_files = os.listdir(temp_extract_folder)
            
            for f_name in extracted_files:
                source_file = os.path.join(temp_extract_folder, f_name)
                
                if os.path.isfile(source_file):
                    fname_no_ext, ext = os.path.splitext(f_name)
                    
                    # Target Path
                    dest_name = f_name
                    dest_path = os.path.join(download_path, dest_name)
                    
                    # Collision Check -> Add 01, 02...
                    counter = 1
                    while os.path.exists(dest_path):
                        dest_name = f"{fname_no_ext}_{counter:02d}{ext}"
                        dest_path = os.path.join(download_path, dest_name)
                        counter += 1
                    
                    shutil.move(source_file, dest_path)
                    self.log(f"     ✅ Saved JSON: {dest_name}")

            # 4. CLEANUP
            shutil.rmtree(temp_extract_folder) 
            os.remove(target_zip) 
            
            return True

        except Exception as e:
            self.log(f"     ⚠️ Unzip/Move Error: {str(e)[:30]}")
            return False


# --- GUI CLASS ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GST Bulk Downloader - GSTR-1 Edition")
        self.geometry("900x850")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.worker = None
        self.excel_file = ""
        self.manual_credentials = []

        # HEADER
        self.head = ctk.CTkFrame(self, fg_color="#1a237e", corner_radius=0, height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        self.head.grid_propagate(False) 
        ctk.CTkLabel(self.head, text="GST BULK DOWNLOADER", 
                     font=("Roboto Medium", 24, "bold"), text_color="white").pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(self.head, text="GSTR-1 JSON AUTOMATION", 
                     font=("Roboto", 14), text_color="#bbdefb").pack(side="right", padx=20, pady=15)

        # SETTINGS
        self.settings_container = ctk.CTkFrame(self, fg_color="transparent")
        self.settings_container.grid(row=1, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.settings_container.grid_columnconfigure((0, 1), weight=1)

        # Credentials Card
        self.card_cred = ctk.CTkFrame(self.settings_container, border_color="#3949ab", border_width=1)
        self.card_cred.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ctk.CTkLabel(self.card_cred, text="📂 Credentials Source", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))
        self.ent_file = ctk.CTkEntry(self.card_cred, placeholder_text="Add ID/Password or select Excel file (optional)...", height=35)
        self.ent_file.pack(fill="x", padx=15, pady=(5, 10))
        self.btn_browse = ctk.CTkButton(self.card_cred, text="Browse File", command=self.browse_file, 
                                        fg_color="#3949ab", hover_color="#283593", height=35)
        
        self.btn_browse.pack(fill="x", padx=15, pady=(0, 15))
        btn_row = ctk.CTkFrame(self.card_cred, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(5, 15))
        self.btn_download = ctk.CTkButton(btn_row, text="➕ Add ID Password", command=self.add_id_password, fg_color="#43a047", hover_color="#2e7d32", height=28, font=("Arial", 12, "bold"))
        self.btn_download.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.btn_demo = ctk.CTkButton(btn_row, text="▶ View Demo", command=self.open_demo_link, fg_color="#e53935", hover_color="#b71c1c", height=28, font=("Arial", 12, "bold"))
        self.btn_demo.pack(side="left", expand=True, fill="x", padx=(5, 0))

        # Period Settings Card
        self.card_period = ctk.CTkFrame(self.settings_container, border_color="#3949ab", border_width=1)
        self.card_period.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        ctk.CTkLabel(self.card_period, text="📅 Period & Action", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 5))

        # Action Mode (New)
        self.frm_act = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_act.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_act, text="Action Mode:", width=100, anchor="w", text_color="#ffd740").pack(side="left")
        self.cb_action = ctk.CTkComboBox(self.frm_act, values=["Request JSON", "Download JSON"], width=150)
        self.cb_action.set("Request JSON")
        self.cb_action.pack(side="right", expand=True, fill="x")

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
        
        # Monthly mode only
        self.chk_all_qtr_var = ctk.BooleanVar(value=False)
        ctk.CTkLabel(self.card_period, text="Monthly download mode enabled", text_color="gray").pack(anchor="w", padx=15, pady=5)

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
        self.btn_start = ctk.CTkButton(self.btn_row_footer, text="START GSTR-1 PROCESS", height=50, font=("Arial", 16, "bold"),
                                       fg_color="#2e7d32", hover_color="#1b5e20", command=self.start_process)
        self.btn_start.pack(side="left", expand=True, fill="x")
        self.btn_stop = ctk.CTkButton(self.btn_row_footer, text="⏹ STOP", height=50, font=("Arial", 16, "bold"),
                                      fg_color="#c62828", hover_color="#8e0000", command=self.stop_process, width=150)
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_stop.pack_forget()

    def toggle_inputs(self):
        self.cb_qtr.configure(state="normal")
        self.cb_month.configure(state="normal")

    def update_months_based_on_qtr(self, choice):
        if "Quarter 1" in choice: vals = ["April", "May", "June"]
        elif "Quarter 2" in choice: vals = ["July", "August", "September"]
        elif "Quarter 3" in choice: vals = ["October", "November", "December"]
        elif "Quarter 4" in choice: vals = ["January", "February", "March"]
        else: vals = ["April", "May", "June"]
        self.cb_month.configure(values=vals)
        self.cb_month.set(vals[0])
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

            self.manual_credentials.append({"Username": username, "Password": password})
            self.excel_file = ""
            self.ent_file.delete(0, "end")
            self.ent_file.insert(0, f"Manual IDs added: {len(self.manual_credentials)}")
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
            self.btn_start.configure(state="normal", text="STOPPED" if is_stopped else "START GSTR-1 PROCESS")
            self.btn_stop.pack_forget()
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.cap_stop_btn.configure(state="normal", text="⏹ STOP PROCESS")
            if is_stopped:
                self.after(1200, lambda: self.btn_start.configure(text="START GSTR-1 PROCESS"))
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
        credentials = list(self.manual_credentials)
        if not credentials and not self.excel_file:
            messagebox.showerror("Error", "Please add ID/Password or select Excel file")
            return
        settings = {
            "year": self.cb_year.get(),
            "month": self.cb_month.get(),
            "quarter": self.cb_qtr.get(),
            "all_quarters": False,
            "action_mode": self.cb_action.get()
        }
        self.close_captcha_safe()
        self.cap_stop_btn.configure(state="normal", text="⏹ STOP PROCESS")
        self.btn_stop.configure(state="normal", text="⏹ STOP")
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.pack(side="left", padx=(10, 0))
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