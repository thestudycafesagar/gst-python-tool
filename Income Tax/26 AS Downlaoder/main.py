import threading
import time
import os
import shutil
import pandas as pd
import customtkinter as ctk
import re
from datetime import datetime
from tkinter import filedialog, messagebox
from functools import wraps



# PDF Unlocker Import
try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except Exception:
    PYPDF_AVAILABLE = False

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ============================================================
#  POPUP WINDOW CLASS (For Year Selection)
# ============================================================
class YearSelectionPopup(ctk.CTkToplevel):
    def __init__(self, parent, years_found, user_id, callback):
        super().__init__(parent)
        self.callback = callback
        self.title(f"Select Years for: {user_id}")
        self.geometry("400x550")
        self.attributes("-topmost", True)
        self.transient(parent)
        self.grab_set()

        self.label = ctk.CTkLabel(self, text=f"⚠️ User: {user_id}\nSelect Years to Download:", 
                                font=ctk.CTkFont(size=16, weight="bold"))
        self.label.pack(pady=20)

        self.scroll_frame = ctk.CTkScrollableFrame(self, width=300, height=350)
        self.scroll_frame.pack(pady=10, padx=20)

        self.check_vars = {}
        for year in years_found:
            var = ctk.StringVar(value="off")
            if years_found.index(year) < 3: 
                var.set(year)
            
            chk = ctk.CTkCheckBox(self.scroll_frame, text=year, variable=var, onvalue=year, offvalue="off")
            chk.pack(anchor="w", pady=5)
            self.check_vars[year] = var

        self.btn_confirm = ctk.CTkButton(self, text="CONFIRM & DOWNLOAD", command=self.on_confirm, 
                                       fg_color="green", hover_color="darkgreen", height=40)
        self.btn_confirm.pack(pady=20)

    def on_confirm(self):
        selected = []
        for year, var in self.check_vars.items():
            if var.get() != "off":
                selected.append(year)
        
        if not selected:
            messagebox.showwarning("Warning", "Please select at least one year!")
            return

        self.callback(selected)
        self.grab_release()
        self.destroy()

# ============================================================
#  BASE HELPER FUNCTIONS (Shared by all Workers)
# ============================================================
def retry_on_failure(max_attempts=3, delay=1, exceptions=(Exception,)):
    """Decorator to retry a function on failure with exponential backoff."""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    if attempt == max_attempts:
                        raise
                    wait_time = delay * attempt
                    time.sleep(wait_time)
            return None
        return wrapper
    return decorator

def normalize_columns(df):
    user_col = None
    pass_col = None
    dob_col = None
    clean_cols = {c: str(c).lower().strip().replace(" ", "").replace("_", "").replace(".", "") for c in df.columns}
    
    pan_patterns = ['userid', 'user', 'pan', 'pannumber', 'panid', 'loginid', 'panno']
    pass_patterns = ['password', 'pass', 'pwd', 'loginpass']
    dob_patterns = ['dob', 'dateofbirth', 'birthdate', 'incorporationdate']

    for original, clean in clean_cols.items():
        if not user_col and any(p in clean for p in pan_patterns):
            user_col = original
        elif not pass_col and any(p in clean for p in pass_patterns):
            pass_col = original
        elif not dob_col and any(p in clean for p in dob_patterns):
            dob_col = original
            
    return user_col, pass_col, dob_col

def create_unique_folder(base_dir, folder_name):
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
        
    full_path = os.path.join(base_dir, folder_name)
    if not os.path.exists(full_path):
        os.makedirs(full_path)
        return full_path
    
    counter = 1
    while True:
        new_name = f"{folder_name} ({counter})"
        full_path = os.path.join(base_dir, new_name)
        if not os.path.exists(full_path):
            os.makedirs(full_path)
            return full_path
        counter += 1

def clean_temp_files(folder, prefixes=("AIS_", "TIS_", "20")):
    """Deletes temporary downloads and un-renamed raw PDFs to prevent false captures."""
    if not os.path.exists(folder): return
    for f in os.listdir(folder):
        if f.endswith(".crdownload") or f.endswith(".tmp"):
            try: os.remove(os.path.join(folder, f))
            except: pass
        elif f.endswith(".pdf"):
            # If the PDF does not start with one of our safe prefixes, it's a raw file (like 'Form26AS.pdf')
            if not any(f.startswith(p) for p in prefixes):
                try: os.remove(os.path.join(folder, f))
                except: pass

def wait_and_rename_file(folder, year_label, logger, prefix="", start_time=None):
    """Waits for a new PDF to appear and renames it. Returns the file path if successful."""
    timeout = 30  # seconds
    if not os.path.exists(folder):
        return None

    if start_time is None:
        start_time = time.time() - 5 

    end_time = time.time() + timeout

    while time.time() < end_time:
        candidates = []
        for f in os.listdir(folder):
            if not f.lower().endswith('.pdf'):
                continue
            full = os.path.join(folder, f)
            try:
                mtime = os.path.getmtime(full)
            except Exception:
                continue
            if f.endswith('.crdownload') or f.endswith('.tmp'):
                continue
            if mtime >= start_time:
                candidates.append((full, mtime))

        if not candidates:
            time.sleep(1)
            continue

        newest_file = max(candidates, key=lambda x: x[1])[0]

        safe_year = year_label.replace("F.Y.", "").replace('/', '-').strip()
        base_target = f"{prefix}{safe_year}.pdf"
        new_name = os.path.join(folder, base_target)

        if os.path.abspath(newest_file) == os.path.abspath(new_name):
            logger(f"        📄 File already named: {base_target}")
            return new_name

        if os.path.exists(new_name):
            i = 1
            while True:
                candidate_name = os.path.join(folder, f"{prefix}{safe_year} ({i}).pdf")
                if not os.path.exists(candidate_name):
                    new_name = candidate_name
                    break
                i += 1

        try:
            os.rename(newest_file, new_name)
            logger(f"        📄 Renamed to: {os.path.basename(new_name)}")
            return new_name
        except Exception as e:
            logger(f"        ⚠️ Rename Failed: {e}")
            return None

    return None

def unlock_pdf(file_path, pan, dob_str, logger):
    """
    Decrypts PDF using standard Income Tax password logic.
    Priority 1 (AIS/TIS): pan(lowercase) + ddmmyyyy
    Priority 2 (26AS): ddmmyyyy
    """
    if not PYPDF_AVAILABLE:
        logger("        ⚠️ 'pypdf' missing. Skipping auto-unlock.")
        return False
        
    if pd.isna(dob_str) or str(dob_str).strip() == "" or str(dob_str).lower() in ['nan', 'nat', 'none']:
        logger("        ⚠️ No DOB found in Excel. Skipping auto-unlock.")
        return False

    try:
        if isinstance(dob_str, pd.Timestamp) or isinstance(dob_str, datetime):
            dob_formatted = dob_str.strftime("%d%m%Y")
        else:
            dob_dt = pd.to_datetime(str(dob_str).strip(), dayfirst=True)
            dob_formatted = dob_dt.strftime("%d%m%Y")
    except Exception as e:
        logger(f"        ⚠️ Invalid DOB format in Excel: {dob_str}")
        return False

    pan_lower = str(pan).strip().lower()
    
    passwords_to_try = [
        f"{pan_lower}{dob_formatted}", # AIS/TIS Standard
        dob_formatted                   # 26AS Standard
    ]

    try:
        reader = PdfReader(file_path)
        if not reader.is_encrypted:
            logger("        🔓 PDF is already unlocked.")
            return True

        unlocked = False
        for pwd in passwords_to_try:
            if reader.decrypt(pwd) != 0:
                unlocked = True
                break
        
        if not unlocked:
            logger(f"        🔒 Unlock failed. Invalid PAN/DOB combinations.")
            return False

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        
        unlocked_path = file_path.replace(".pdf", "_temp_unlocked.pdf")
        with open(unlocked_path, "wb") as f:
            writer.write(f)
        
        os.remove(file_path)
        os.rename(unlocked_path, file_path)
        logger("        🔓 PDF successfully decrypted!")
        return True

    except Exception as e:
        logger(f"        ⚠️ Unlock Error: {e}")
        return False


# ============================================================
#  WORKER 1: 26AS THREAD CLASS
# ============================================================
class Tax26ASWorker:
    def __init__(self, app_instance, excel_path, year_mode):
        self.app = app_instance
        self.excel_path = excel_path
        self.year_mode = year_mode
        self.keep_running = True
        self.report_data = []
        self.user_selection_event = threading.Event()
        self.current_user_selected_years = None

    def log(self, message):
        self.app.update_log_safe_26as(message)

    def set_years_and_resume(self, selected_list):
        self.current_user_selected_years = selected_list
        self.user_selection_event.set()

    def run(self):
        self.log("🚀 INITIALIZING 26AS ENGINE...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        
        try:
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)
            
            if not user_col or not pass_col:
                self.log(f"❌ ERROR: Headers missing. Need 'PAN' and 'Password'.")
                self.app.process_finished_safe_26as("Failed: Column Header Error")
                return

            self.log(f"✅ Mapped Columns -> ID: '{user_col}', Pass: '{pass_col}', DOB: '{dob_col}'")
            total_users = len(df)
            
            for index, row in df.iterrows():
                if not self.keep_running: 
                    self.log("🛑 Process Stopped by User.")
                    break
                
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None
                
                self.app.update_progress_safe_26as((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "26AS_Downloads")
                final_path = create_unique_folder(download_root, user_id)

                status, reason = self.process_single_user(user_id, password, dob, final_path)
                
                self.report_data.append({
                    "PAN": user_id, "Status": status, "Details": reason,
                    "Folder Saved": os.path.basename(final_path),
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                self.log("-" * 40)
            
            self.generate_report()
            self.app.update_progress_safe_26as(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe_26as("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe_26as("Critical Error Occurred")

    def generate_report(self):
        try:
            if not self.report_data: return
            df_report = pd.DataFrame(self.report_data)
            filename = f"26AS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df_report.to_excel(filename, index=False)
            self.log(f"📄 Report saved: {filename}")
        except: pass

    def process_single_user(self, user_id, password, dob, download_folder):
        driver = None
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument("--disable-blink-features=AutomationControlled")
            prefs = {
                "download.default_directory": download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "download_restrictions": 0,
                "safebrowsing.enabled": True,
                "safebrowsing.disable_download_protection": True
            }
            options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            
            # Set aggressive timeouts to prevent hanging
            driver.set_page_load_timeout(30)  # 30 seconds for page load
            driver.set_script_timeout(30)     # 30 seconds for script execution
            driver.implicitly_wait(10)        # 10 seconds for element search
            
            wait = WebDriverWait(driver, 20)
            actions = ActionChains(driver)

            # 1. LOGIN WITH COMPREHENSIVE RETRY
            login_success = False
            for login_attempt in range(1, 4):
                if login_success: break
                if login_attempt > 1:
                    self.log(f"   ⚠️ Login Retry {login_attempt}/3...")
                    try:
                        driver.delete_all_cookies()
                        driver.refresh()
                    except: pass
                    time.sleep(3)

                try:
                    self.log("   🌐 Opening Portal...")
                    try:
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")
                        time.sleep(2)
                    except TimeoutException:
                        self.log("   ⚠️ Page load timeout. Retrying...")
                        continue
                    except Exception as e:
                        self.log(f"   ⚠️ Page load error: {str(e)[:30]}. Retrying...")
                        continue
                    
                    try: driver.switch_to.alert.accept(); time.sleep(1)
                    except: pass

                    # Step 1: Enter PAN with retry
                    pan_entered = False
                    for pan_retry in range(3):
                        try:
                            pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                            pan_field.clear()
                            pan_field.send_keys(user_id)
                            pan_entered = True
                            break
                        except Exception as e:
                            if pan_retry == 2:
                                self.log(f"   ⚠️ Failed to enter PAN after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not pan_entered:
                        continue
                    
                    time.sleep(0.5)
                    
                    # Step 2: Click Continue button with retry
                    for cont_retry in range(3):
                        try:
                            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                            driver.execute_script("arguments[0].click();", continue_btn)
                            break
                        except Exception as e:
                            if cont_retry == 2:
                                self.log(f"   ⚠️ Failed to click Continue after 3 tries")
                                raise
                            time.sleep(1)
                    
                    time.sleep(1.5)
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN"

                    # Step 3: Enter Password with retry
                    for pwd_retry in range(3):
                        try:
                            pwd_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                            pwd_field.clear()
                            pwd_field.send_keys(password)
                            break
                        except Exception as e:
                            if pwd_retry == 2:
                                self.log(f"   ⚠️ Failed to enter password after 3 tries")
                                raise
                            time.sleep(1)
                    
                    # Show password checkbox
                    try: 
                        driver.execute_script("document.getElementById('passwordCheckBox-input').click();")
                        time.sleep(0.3)
                    except: pass
                    
                    self.log("   ⏳ Waiting for security check (3s)...")
                    time.sleep(3.5)
                    
                    # Step 4: Submit login with retry
                    for submit_retry in range(3):
                        try:
                            driver.execute_script("document.querySelector('button.large-button-primary').click();")
                            break
                        except Exception as e:
                            if submit_retry == 2:
                                self.log(f"   ⚠️ Failed to submit login after 3 tries")
                                raise
                            time.sleep(1)

                    # Step 5: Wait for successful login
                    for _ in range(20):
                        time.sleep(1)
                        try:
                            if driver.find_elements(By.ID, "e-File"):
                                self.log("   ✅ Login Successful!")
                                login_success = True; break
                        except: pass
                        
                        if "Invalid Password" in driver.page_source: return "Failed", "Invalid Password"
                        
                        try:
                            dual = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual and dual[0].is_displayed():
                                driver.execute_script("arguments[0].click();", dual[0])
                                time.sleep(2)
                        except: pass
                    if login_success: break
                except Exception as e:
                    self.log(f"   ⚠️ Login Error: {str(e)[:50]}")
                    if login_attempt < 3:
                        time.sleep(2)

            if not login_success: return "Failed", "Login Timeout"

            # 2. NAVIGATE TO 26AS MENU (with retry logic)
            self.log("   🚀 Navigating to Form 26AS...")
            nav_success = False
            for nav_attempt in range(1, 4):
                try:
                    if nav_attempt > 1:
                        self.log(f"   ⚠️ Navigation Retry {nav_attempt}/3...")
                        driver.refresh()
                        time.sleep(2)
                    
                    efile = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "e-File")))
                    driver.execute_script("arguments[0].click();", efile)
                    submenu = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Income Tax Returns')]")))
                    actions.move_to_element(submenu).perform()
                    time.sleep(0.3)
                    view_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'View Form 26AS')]")))
                    driver.execute_script("arguments[0].click();", view_btn)
                    time.sleep(2)
                    nav_success = True
                    break
                except Exception as e:
                    self.log(f"   ⚠️ Nav Attempt {nav_attempt} Failed: {str(e)[:40]}")
                    if nav_attempt == 3:
                        return "Failed", "Menu Nav Error"
            
            if not nav_success: return "Failed", "Menu Nav Error"

            # 3. DISCLAIMER (with retry logic)
            self.log("   ⚠️ Checking Disclaimer...")
            for disc_attempt in range(1, 4):
                try:
                    WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.ID, "Details")))
                    driver.execute_script("document.getElementById('Details').checked = true;")
                    try: driver.execute_script("checkModal('modalPagee');")
                    except: pass
                    time.sleep(0.3)
                    driver.execute_script("document.getElementById('btn').disabled = false;")
                    driver.execute_script("document.getElementById('btn').click();")
                    self.log("     -> Disclaimer Accepted.")
                    time.sleep(2)
                    break
                except:
                    if disc_attempt == 1:
                        self.log("     -> No Disclaimer found. Proceeding...")
                    break

            # Switch to TRACES tab with retry
            traces_tab_found = False
            for tab_attempt in range(1, 4):
                try:
                    time.sleep(1)
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        traces_tab_found = True
                        break
                    else:
                        if tab_attempt < 3:
                            self.log(f"   ⚠️ Waiting for TRACES tab... Attempt {tab_attempt}/3")
                            time.sleep(2)
                except Exception as e:
                    if tab_attempt == 3:
                        return "Failed", "TRACES tab did not open"
            
            if not traces_tab_found: return "Failed", "TRACES tab did not open"

            # 4. TRACES PORTAL (with retry logic)
            traces_nav_success = False
            for traces_attempt in range(1, 4):
                try:
                    if traces_attempt > 1:
                        self.log(f"   ⚠️ TRACES Retry {traces_attempt}/3...")
                        driver.refresh()
                        time.sleep(2)
                    
                    # Check for any popups
                    try:
                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "cbox")))
                        driver.execute_script("document.getElementById('cbox').click();")
                        driver.execute_script("document.getElementById('proceed').click();")
                        time.sleep(1)
                    except: pass

                    self.log("   🖱️ Clicking 'View Tax Credit' Link...")
                    try:
                        link_26as = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "View Tax Credit")))
                        driver.execute_script("arguments[0].click();", link_26as)
                        traces_nav_success = True
                    except TimeoutException:
                        self.log("   ⚠️ Direct link not found. Attempting Menu Navigation...")
                        menu_xpath = "//span[contains(text(), 'View/ Verify Tax Credit')]"
                        menu_item = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, menu_xpath)))
                        driver.execute_script("arguments[0].click();", menu_item)
                        time.sleep(1)
                        submenu_xpath = "//a[contains(text(), 'View Form 26AS')]"
                        sub_item = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, submenu_xpath)))
                        driver.execute_script("arguments[0].click();", sub_item)
                        traces_nav_success = True
                    
                    if traces_nav_success:
                        break
                        
                except Exception as e:
                    self.log(f"   ⚠️ TRACES Attempt {traces_attempt} Failed: {str(e)[:40]}")
                    if traces_attempt == 3:
                        return "Failed", "TRACES Navigation Failed"
            
            if not traces_nav_success: return "Failed", "TRACES Navigation Failed"
            
            try:
                self.log("   📥 Fetching Available Years...")
                dropdown_id = "AssessmentYearDropDown"
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, dropdown_id)))
                ay_select = Select(driver.find_element(By.ID, dropdown_id))
                available_years = [o.text.strip() for o in ay_select.options if "Select" not in o.text]
                if not available_years: return "Failed", "No years found"

                if self.year_mode == "Current Year": self.current_user_selected_years = available_years[:1]
                elif self.year_mode == "Current and Last Year": self.current_user_selected_years = available_years[:2]
                elif self.year_mode == "Current and Last 2 Years": self.current_user_selected_years = available_years[:3]
                else:
                    self.log(f"   🛑 PAUSED: Found {len(available_years)} years. Waiting for you...")
                    self.user_selection_event.clear()
                    self.current_user_selected_years = None
                    self.app.trigger_year_selection(available_years, user_id, self.set_years_and_resume)
                    self.user_selection_event.wait()

                years_to_download = [y for y in self.current_user_selected_years if y in available_years]
                if not years_to_download: return "Warning", "No valid years selected"

                self.log(f"   ⬇️ Downloading {len(years_to_download)} Years...")
                count = 0
                for year in years_to_download:
                    year_success = False
                    for year_attempt in range(1, 4):
                        try:
                            if year_attempt > 1:
                                self.log(f"     -> Retry {year_attempt}/3 for {year}...")
                                driver.refresh()
                                time.sleep(2)
                                dropdown_id = "AssessmentYearDropDown"
                                WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, dropdown_id)))
                            else:
                                self.log(f"     -> Processing {year}...")
                            
                            Select(driver.find_element(By.ID, dropdown_id)).select_by_visible_text(year)
                            time.sleep(1)
                            Select(driver.find_element(By.ID, "viewType")).select_by_value("HTML")
                            view_btn = driver.find_element(By.ID, "btnSubmit")
                            driver.execute_script("arguments[0].click();", view_btn)
                            
                            self.log("        Generating HTML Data...")
                            time.sleep(3)
                            year_success = True
                            break
                        except Exception as e:
                            if year_attempt == 3:
                                self.log(f"        ⚠️ Failed to process {year} after 3 attempts")
                                continue
                    
                    if not year_success:
                        continue 
                    
                    # PDF Download with retry
                    pdf_download_success = False
                    for pdf_attempt in range(1, 4):
                        try:
                            if pdf_attempt > 1:
                                self.log(f"        ⚠️ PDF Download Retry {pdf_attempt}/3...")
                                time.sleep(2)
                            
                            clean_temp_files(download_folder)

                            dl_click_time = time.time()
                            pdf_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "pdfBtn")))
                            driver.execute_script("arguments[0].click();", pdf_btn)
                            self.log("        ✅ Export Triggered.")
                            
                            saved_path = wait_and_rename_file(download_folder, year, self.log, prefix="", start_time=dl_click_time)
                            if saved_path:
                                count += 1
                                unlock_pdf(saved_path, user_id, dob, self.log)
                                pdf_download_success = True
                                break
                            else:
                                self.log("        ❌ File capture failed.")
                                if pdf_attempt < 3:
                                    continue
                        except Exception as e:
                            self.log(f"       ⚠️ PDF Export Attempt {pdf_attempt} Failed: {str(e)[:30]}")
                            if pdf_attempt == 3:
                                self.log("       ❌ PDF Export failed after 3 attempts.")

                return "Success", f"Downloaded {count} files"

            except Exception as e: return "Failed", f"TRACES Error: {str(e)[:20]}"
        except Exception: return "Failed", "Browser Crash"
        finally:
            if driver: driver.quit()


# ============================================================
#  WORKER 2: AIS THREAD CLASS
# ============================================================
class AISWorker:
    def __init__(self, app_instance, excel_path, year_mode):
        self.app = app_instance
        self.excel_path = excel_path
        self.year_mode = year_mode
        self.keep_running = True
        self.report_data = []
        self.user_selection_event = threading.Event()
        self.current_user_selected_years = None

    def log(self, message):
        self.app.update_log_safe_ais(message)

    def set_years_and_resume(self, selected_list):
        self.current_user_selected_years = selected_list
        self.user_selection_event.set()

    def run(self):
        self.log("🚀 INITIALIZING AIS ENGINE...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        
        try:
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)
            
            if not user_col or not pass_col:
                self.log(f"❌ ERROR: Headers missing.")
                self.app.process_finished_safe_ais("Failed: Column Header Error")
                return

            total_users = len(df)
            
            for index, row in df.iterrows():
                if not self.keep_running: break
                
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None
                
                self.app.update_progress_safe_ais((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "AIS_Downloads")
                final_path = create_unique_folder(download_root, user_id)

                status, reason = self.process_single_user(user_id, password, dob, final_path)
                
                self.report_data.append({
                    "PAN": user_id, "Status": status, "Details": reason,
                    "Folder Saved": os.path.basename(final_path),
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                self.log("-" * 40)
            
            self.generate_report()
            self.app.update_progress_safe_ais(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe_ais("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe_ais("Critical Error Occurred")

    def generate_report(self):
        try:
            if not self.report_data: return
            df_report = pd.DataFrame(self.report_data)
            filename = f"AIS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df_report.to_excel(filename, index=False)
            self.log(f"📄 Report saved: {filename}")
        except: pass

    def process_single_user(self, user_id, password, dob, download_folder):
        driver = None
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument("--disable-blink-features=AutomationControlled")
            prefs = {
                "download.default_directory": download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "download_restrictions": 0,
                "safebrowsing.enabled": True,
                "safebrowsing.disable_download_protection": True
            }
            options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            
            # Set aggressive timeouts to prevent hanging
            driver.set_page_load_timeout(30)  # 30 seconds for page load
            driver.set_script_timeout(30)     # 30 seconds for script execution
            driver.implicitly_wait(10)        # 10 seconds for element search
            
            wait = WebDriverWait(driver, 20)
            
            # 1. LOGIN WITH COMPREHENSIVE RETRY (AIS)
            login_success = False
            for login_attempt in range(1, 4):
                if login_success: break
                if login_attempt > 1:
                    self.log(f"   ⚠️ Login Retry {login_attempt}/3...")
                    try:
                        driver.delete_all_cookies()
                        driver.refresh()
                    except: pass
                    time.sleep(3)

                try:
                    self.log("   🌐 Opening Portal...")
                    try:
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")
                        time.sleep(2)
                    except TimeoutException:
                        self.log("   ⚠️ Page load timeout. Retrying...")
                        continue
                    except Exception as e:
                        self.log(f"   ⚠️ Page load error: {str(e)[:30]}. Retrying...")
                        continue
                    
                    try: driver.switch_to.alert.accept(); time.sleep(1)
                    except: pass

                    # Step 1: Enter PAN with retry
                    pan_entered = False
                    for pan_retry in range(3):
                        try:
                            pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                            pan_field.clear()
                            pan_field.send_keys(user_id)
                            pan_entered = True
                            break
                        except Exception as e:
                            if pan_retry == 2:
                                self.log(f"   ⚠️ Failed to enter PAN after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not pan_entered:
                        continue
                    
                    time.sleep(0.5)
                    
                    # Step 2: Click Continue button with retry
                    continue_success = False
                    for cont_retry in range(3):
                        try:
                            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                            driver.execute_script("arguments[0].click();", continue_btn)
                            time.sleep(1)
                            
                            # Check for ITD-EXEC2002 error
                            if "ITD-EXEC2002" in driver.page_source or "Something seems to have gone wrong" in driver.page_source:
                                self.log(f"   ⚠️ ITD-EXEC2002 error detected, retrying Continue button...")
                                if cont_retry < 2:
                                    continue
                                else:
                                    raise Exception("ITD-EXEC2002 error persists after retries")
                            
                            continue_success = True
                            break
                        except Exception as e:
                            if cont_retry == 2:
                                self.log(f"   ⚠️ Failed to click Continue after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not continue_success:
                        continue
                    
                    time.sleep(1.5)
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN"

                    # Step 3: Enter Password with retry
                    for pwd_retry in range(3):
                        try:
                            pwd_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                            pwd_field.clear()
                            pwd_field.send_keys(password)
                            break
                        except Exception as e:
                            if pwd_retry == 2:
                                self.log(f"   ⚠️ Failed to enter password after 3 tries")
                                raise
                            time.sleep(1)
                    
                    # Show password checkbox
                    try: 
                        driver.execute_script("document.getElementById('passwordCheckBox-input').click();")
                        time.sleep(0.3)
                    except: pass
                    
                    self.log("   ⏳ Waiting for security check (3s)...")
                    time.sleep(3.5)
                    
                    # Step 4: Submit login with retry
                    submit_success = False
                    for submit_retry in range(3):
                        try:
                            driver.execute_script("document.querySelector('button.large-button-primary').click();")
                            time.sleep(1)
                            
                            # Check for ITD-EXEC2002 error
                            if "ITD-EXEC2002" in driver.page_source or "Something seems to have gone wrong" in driver.page_source:
                                self.log(f"   ⚠️ ITD-EXEC2002 error detected, retrying Login button...")
                                if submit_retry < 2:
                                    continue
                                else:
                                    raise Exception("ITD-EXEC2002 error persists after retries")
                            
                            submit_success = True
                            break
                        except Exception as e:
                            if submit_retry == 2:
                                self.log(f"   ⚠️ Failed to submit login after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not submit_success:
                        continue

                    # Step 5: Wait for successful login
                    for _ in range(20):
                        time.sleep(1)
                        try:
                            if driver.find_elements(By.ID, "e-File"):
                                self.log("   ✅ Login Successful!")
                                login_success = True; break
                        except: pass
                        
                        if "Invalid Password" in driver.page_source: return "Failed", "Invalid Password"
                        
                        try:
                            dual = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual and dual[0].is_displayed():
                                driver.execute_script("arguments[0].click();", dual[0])
                                time.sleep(2)
                        except: pass
                    if login_success: break
                except Exception as e:
                    self.log(f"   ⚠️ Login Error: {str(e)[:50]}")
                    if login_attempt < 3:
                        time.sleep(2)

            if not login_success: return "Failed", "Login Timeout"

            # 2. NAVIGATE TO AIS (with retry logic)
            self.log("   🚀 Navigating to AIS...")
            ais_nav_success = False
            for ais_nav_attempt in range(1, 4):
                try:
                    if ais_nav_attempt > 1:
                        self.log(f"   ⚠️ AIS Navigation Retry {ais_nav_attempt}/3...")
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/dashboard")
                        time.sleep(2)
                    
                    ais_span = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'mdc-button__label') and contains(text(), 'AIS')]")))
                    driver.execute_script("arguments[0].click();", ais_span)
                    try:
                        proceed_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Proceed')]")))
                        driver.execute_script("arguments[0].click();", proceed_btn)
                    except: pass
                    time.sleep(2)
                    ais_nav_success = True
                    break
                except Exception as e:
                    self.log(f"   ⚠️ AIS Nav Attempt {ais_nav_attempt} Failed: {str(e)[:40]}")
                    if ais_nav_attempt == 3:
                        return "Failed", "Dashboard AIS Menu Not Found"
            
            if not ais_nav_success: return "Failed", "Dashboard AIS Menu Not Found"

            # Switch to AIS tab with retry
            ais_tab_found = False
            for tab_attempt in range(1, 4):
                try:
                    time.sleep(1)
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        ais_tab_found = True
                        break
                    else:
                        if tab_attempt < 3:
                            self.log(f"   ⚠️ Waiting for AIS tab... Attempt {tab_attempt}/3")
                            time.sleep(2)
                except Exception as e:
                    if tab_attempt == 3:
                        return "Failed", "AIS Tab did not open"
            
            if not ais_tab_found: return "Failed", "AIS Tab did not open"

            # 4. AIS INTERNAL LOGIC (with retry)
            ais_internal_success = False
            for ais_internal_attempt in range(1, 4):
                try:
                    if ais_internal_attempt > 1:
                        self.log(f"   ⚠️ AIS Internal Menu Retry {ais_internal_attempt}/3...")
                        driver.refresh()
                        time.sleep(2)
                    
                    ais_internal_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]")))
                    driver.execute_script("arguments[0].click();", ais_internal_menu)
                    time.sleep(2)
                    ais_internal_success = True
                    break
                except Exception as e:
                    self.log(f"   ⚠️ AIS Internal Attempt {ais_internal_attempt} Failed: {str(e)[:40]}")
                    if ais_internal_attempt == 3:
                        return "Failed", "AIS Internal Menu Failed"
            
            if not ais_internal_success: return "Failed", "AIS Internal Menu Failed"
            
            try:
                self.log("   📥 Fetching Available Years...")
                try:
                    dropdown_toggle = wait.until(EC.presence_of_element_located((By.ID, "dropdownMenuButton")))
                    driver.execute_script("arguments[0].click();", dropdown_toggle)
                    time.sleep(0.5)
                except: pass

                year_buttons = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//button[contains(@class, 'dropdown-item') and contains(text(), 'F.Y.')]")))
                available_years = []
                for btn in year_buttons:
                    txt = btn.get_attribute("textContent").strip()
                    if txt and txt not in available_years:
                        available_years.append(txt)

                try: driver.execute_script("arguments[0].click();", dropdown_toggle)
                except: pass

                if not available_years: return "Failed", "No years found in AIS"

                if self.year_mode == "Current Year": self.current_user_selected_years = available_years[:1]
                elif self.year_mode == "Current and Last Year": self.current_user_selected_years = available_years[:2]
                elif self.year_mode == "Current and Last 2 Years": self.current_user_selected_years = available_years[:3]
                else:
                    self.log(f"   🛑 PAUSED: Found {len(available_years)} years. Waiting for you...")
                    self.user_selection_event.clear()
                    self.current_user_selected_years = None
                    self.app.trigger_year_selection(available_years, user_id, self.set_years_and_resume)
                    self.user_selection_event.wait()

                years_to_download = [y for y in self.current_user_selected_years if y in available_years]
                if not years_to_download: return "Warning", "No valid years selected"

                self.log(f"   ⬇️ Downloading {len(years_to_download)} Years...")
                count = 0
                for year in years_to_download:
                    year_success = False
                    for year_attempt in range(1, 4):
                        try:
                            if year_attempt > 1:
                                self.log(f"     -> Retry {year_attempt}/3 for {year}...")
                                driver.refresh()
                                time.sleep(2)
                                # Re-click AIS menu
                                ais_internal_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]")))
                                driver.execute_script("arguments[0].click();", ais_internal_menu)
                                time.sleep(2)
                            else:
                                self.log(f"     -> Processing {year}...")
                            
                            # Open dropdown and select year
                            try:
                                dropdown_toggle = driver.find_element(By.ID, "dropdownMenuButton")
                                driver.execute_script("arguments[0].click();", dropdown_toggle)
                                time.sleep(0.5)
                            except: pass

                            year_xpath = f"//button[contains(@class, 'dropdown-item') and contains(text(), '{year}')]"
                            target_yr_btn = wait.until(EC.presence_of_element_located((By.XPATH, year_xpath)))
                            driver.execute_script("arguments[0].click();", target_yr_btn)
                            time.sleep(2)
                            
                            # Click download icon
                            dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Download AIS related documents']")))
                            driver.execute_script("arguments[0].click();", dl_icon)
                            time.sleep(1)
                            year_success = True
                            break
                        except Exception as e:
                            if year_attempt == 3:
                                self.log(f"        ⚠️ Failed to process {year} after 3 attempts")
                    
                    if not year_success:
                        continue
                    
                    # PDF Download with retry
                    pdf_download_success = False
                    for pdf_attempt in range(1, 4):
                        try:
                            if pdf_attempt > 1:
                                self.log(f"        ⚠️ PDF Download Retry {pdf_attempt}/3...")
                                time.sleep(2)
                            
                            clean_temp_files(download_folder, prefixes=("AIS_",))

                            modal_dl_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Download' and contains(@class, 'btn-outline-primary')]")))
                            modal_click_time = time.time()
                            driver.execute_script("arguments[0].click();", modal_dl_btn)
                            self.log("        Generating Document...")
                            
                            # --- INDIVIDUAL VS COMPANY LOGIC ---
                            direct_download = False
                            prefix = "AIS_"
                            for _ in range(8): # Monitor folder for 8 seconds
                                time.sleep(1)
                                for f in os.listdir(download_folder):
                                    f_path = os.path.join(download_folder, f)
                                    if os.path.isfile(f_path) and os.path.getmtime(f_path) >= modal_click_time - 2:
                                        if f.endswith(".crdownload") or f.endswith(".pdf"):
                                            direct_download = True
                                            break
                                if direct_download: break

                            if direct_download:
                                self.log("        ✅ Direct download detected.")
                                saved_path = wait_and_rename_file(download_folder, year, self.log, prefix=prefix, start_time=modal_click_time-2)
                                if saved_path:
                                    count += 1
                                    unlock_pdf(saved_path, user_id, dob, self.log)
                                    pdf_download_success = True
                                    break
                                else:
                                    self.log("        ❌ File capture failed.")
                                try:
                                    close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                    driver.execute_script("arguments[0].click();", close_btn)
                                except: pass
                            else:
                                self.log("        ℹ️ No direct download. Checking Activity History...")
                                try:
                                    close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                    driver.execute_script("arguments[0].click();", close_btn)
                                    time.sleep(0.5)
                                except: pass

                                history_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Go To Activity History']")))
                                driver.execute_script("arguments[0].click();", history_btn)
                                time.sleep(3)

                                clean_temp_files(download_folder, prefixes=("AIS_",))
                                
                                hist_click_time = time.time()
                                final_dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "(//img[@alt='Download'])[1]")))
                                driver.execute_script("arguments[0].click();", final_dl_icon)
                                self.log("        ✅ Export Triggered from History.")
                                
                                saved_path = wait_and_rename_file(download_folder, year, self.log, prefix=prefix, start_time=hist_click_time-2)
                                if saved_path:
                                    count += 1
                                    unlock_pdf(saved_path, user_id, dob, self.log)
                                    pdf_download_success = True
                                    break
                                else:
                                    self.log("        ❌ File capture failed.")
                                    
                        except Exception as e:
                            self.log(f"       ⚠️ Download Attempt {pdf_attempt} Failed: {str(e)[:30]}")
                            if pdf_attempt == 3:
                                self.log("       ❌ PDF Export failed after 3 attempts.")

                    # Return to AIS tab for the next year
                    try:
                        ais_internal_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]")))
                        driver.execute_script("arguments[0].click();", ais_internal_menu)
                        time.sleep(2)
                    except: pass

                return "Success", f"Downloaded {count} files"

            except Exception as e: return "Failed", f"AIS Portal Error: {str(e)[:20]}"

        except Exception as e: return "Failed", "Browser Crash"
        finally:
            if driver: driver.quit()


# ============================================================
#  WORKER 3: TIS THREAD CLASS
# ============================================================
class TISWorker:
    def __init__(self, app_instance, excel_path, year_mode):
        self.app = app_instance
        self.excel_path = excel_path
        self.year_mode = year_mode
        self.keep_running = True
        self.report_data = []
        self.user_selection_event = threading.Event()
        self.current_user_selected_years = None

    def log(self, message):
        self.app.update_log_safe_tis(message)

    def set_years_and_resume(self, selected_list):
        self.current_user_selected_years = selected_list
        self.user_selection_event.set()

    def run(self):
        self.log("🚀 INITIALIZING TIS ENGINE...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        try:
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)
            if not user_col or not pass_col:
                self.log(f"❌ ERROR: Headers missing.")
                self.app.process_finished_safe_tis("Failed: Column Header Error")
                return

            total_users = len(df)
            for index, row in df.iterrows():
                if not self.keep_running: break
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None
                
                self.app.update_progress_safe_tis((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "TIS_Downloads")
                final_path = create_unique_folder(download_root, user_id)

                status, reason = self.process_single_user(user_id, password, dob, final_path)
                self.report_data.append({
                    "PAN": user_id, "Status": status, "Details": reason,
                    "Folder Saved": os.path.basename(final_path),
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                self.log("-" * 40)

            self.generate_report()
            self.app.update_progress_safe_tis(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe_tis("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe_tis("Critical Error Occurred")

    def generate_report(self):
        try:
            if not self.report_data: return
            df_report = pd.DataFrame(self.report_data)
            filename = f"TIS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df_report.to_excel(filename, index=False)
            self.log(f"📄 Report saved: {filename}")
        except: pass

    def process_single_user(self, user_id, password, dob, download_folder):
        driver = None
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument("--disable-blink-features=AutomationControlled")
            prefs = {
                "download.default_directory": download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "download_restrictions": 0,
                "safebrowsing.enabled": True,
                "safebrowsing.disable_download_protection": True
            }
            options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            
            # Set aggressive timeouts to prevent hanging
            driver.set_page_load_timeout(30)  # 30 seconds for page load
            driver.set_script_timeout(30)     # 30 seconds for script execution
            driver.implicitly_wait(10)        # 10 seconds for element search
            
            wait = WebDriverWait(driver, 20)

            # 1. LOGIN
            login_success = False
            for login_attempt in range(1, 4):
                if login_success: break
                if login_attempt > 1:
                    try:
                        driver.delete_all_cookies()
                        driver.refresh()
                    except: pass
                    time.sleep(3)

                try:
                    self.log("   🌐 Opening Portal...")
                    try:
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")
                    except TimeoutException:
                        self.log("   ⚠️ Page load timeout. Retrying...")
                        continue
                    except Exception as e:
                        self.log(f"   ⚠️ Page load error: {str(e)[:30]}. Retrying...")
                        continue
                    
                    time.sleep(2)
                    try: driver.switch_to.alert.accept(); time.sleep(1)
                    except: pass

                    # Step 1: Enter PAN with retry
                    pan_entered = False
                    for pan_retry in range(3):
                        try:
                            pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                            pan_field.clear()
                            pan_field.send_keys(user_id)
                            pan_entered = True
                            break
                        except Exception as e:
                            if pan_retry == 2:
                                self.log(f"   ⚠️ Failed to enter PAN after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not pan_entered:
                        continue
                    
                    time.sleep(0.5)
                    
                    # Step 2: Click Continue button with retry
                    continue_success = False
                    for cont_retry in range(3):
                        try:
                            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                            driver.execute_script("arguments[0].click();", continue_btn)
                            time.sleep(1)
                            
                            # Check for ITD-EXEC2002 error
                            if "ITD-EXEC2002" in driver.page_source or "Something seems to have gone wrong" in driver.page_source:
                                self.log(f"   ⚠️ ITD-EXEC2002 error detected, retrying Continue button...")
                                if cont_retry < 2:
                                    continue
                                else:
                                    raise Exception("ITD-EXEC2002 error persists after retries")
                            
                            continue_success = True
                            break
                        except Exception as e:
                            if cont_retry == 2:
                                self.log(f"   ⚠️ Failed to click Continue after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not continue_success:
                        continue
                    
                    time.sleep(1.5)
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN"

                    # Step 3: Enter Password with retry
                    for pwd_retry in range(3):
                        try:
                            pwd_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                            pwd_field.clear()
                            pwd_field.send_keys(password)
                            break
                        except Exception as e:
                            if pwd_retry == 2:
                                self.log(f"   ⚠️ Failed to enter password after 3 tries")
                                raise
                            time.sleep(1)
                    
                    # Show password checkbox
                    try: 
                        driver.execute_script("document.getElementById('passwordCheckBox-input').click();")
                        time.sleep(0.3)
                    except: pass
                    
                    self.log("   ⏳ Waiting for security check (3s)...")
                    time.sleep(3.5)
                    
                    # Step 4: Submit login with retry
                    submit_success = False
                    for submit_retry in range(3):
                        try:
                            driver.execute_script("document.querySelector('button.large-button-primary').click();")
                            time.sleep(1)
                            
                            # Check for ITD-EXEC2002 error
                            if "ITD-EXEC2002" in driver.page_source or "Something seems to have gone wrong" in driver.page_source:
                                self.log(f"   ⚠️ ITD-EXEC2002 error detected, retrying Login button...")
                                if submit_retry < 2:
                                    continue
                                else:
                                    raise Exception("ITD-EXEC2002 error persists after retries")
                            
                            submit_success = True
                            break
                        except Exception as e:
                            if submit_retry == 2:
                                self.log(f"   ⚠️ Failed to submit login after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not submit_success:
                        continue

                    # Step 5: Wait for successful login
                    for _ in range(20):
                        time.sleep(1)
                        try:
                            if driver.find_elements(By.ID, "e-File"):
                                self.log("   ✅ Login Successful!")
                                login_success = True; break
                        except: pass
                        
                        if "Invalid Password" in driver.page_source: return "Failed", "Invalid Password"
                        
                        try:
                            dual = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual and dual[0].is_displayed():
                                driver.execute_script("arguments[0].click();", dual[0])
                                time.sleep(2)
                        except: pass
                    if login_success: break
                except Exception as e:
                    self.log(f"   ⚠️ Login Error: {str(e)[:50]}")
                    if login_attempt < 3:
                        time.sleep(2)

            if not login_success: return "Failed", "Login Timeout"

            # 2. NAVIGATE TO AIS (with retry logic)
            self.log("   🚀 Navigating to AIS (for TIS)...")
            tis_nav_success = False
            for tis_nav_attempt in range(1, 4):
                try:
                    if tis_nav_attempt > 1:
                        self.log(f"   ⚠️ TIS Navigation Retry {tis_nav_attempt}/3...")
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/dashboard")
                        time.sleep(2)
                    
                    ais_span = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'mdc-button__label') and contains(text(), 'AIS')]") ))
                    driver.execute_script("arguments[0].click();", ais_span)
                    try:
                        proceed_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Proceed')]") ))
                        driver.execute_script("arguments[0].click();", proceed_btn)
                    except: pass
                    time.sleep(2)
                    tis_nav_success = True
                    break
                except Exception as e:
                    self.log(f"   ⚠️ TIS Nav Attempt {tis_nav_attempt} Failed: {str(e)[:40]}")
                    if tis_nav_attempt == 3:
                        return "Failed", "Dashboard AIS Menu Not Found"
            
            if not tis_nav_success: return "Failed", "Dashboard AIS Menu Not Found"

            # Switch to TIS tab with retry
            tis_tab_found = False
            for tab_attempt in range(1, 4):
                try:
                    time.sleep(1)
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        tis_tab_found = True
                        break
                    else:
                        if tab_attempt < 3:
                            self.log(f"   ⚠️ Waiting for TIS tab... Attempt {tab_attempt}/3")
                            time.sleep(2)
                except Exception as e:
                    if tab_attempt == 3:
                        return "Failed", "AIS Tab did not open"
            
            if not tis_tab_found: return "Failed", "AIS Tab did not open"

            # 3. TIS INTERNAL LOGIC (with retry)
            tis_internal_success = False
            for tis_internal_attempt in range(1, 4):
                try:
                    if tis_internal_attempt > 1:
                        self.log(f"   ⚠️ TIS Internal Menu Retry {tis_internal_attempt}/3...")
                        driver.refresh()
                        time.sleep(2)
                    
                    ais_internal_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]") ))
                    driver.execute_script("arguments[0].click();", ais_internal_menu)
                    time.sleep(2)
                    tis_internal_success = True
                    break
                except Exception as e:
                    self.log(f"   ⚠️ TIS Internal Attempt {tis_internal_attempt} Failed: {str(e)[:40]}")
                    if tis_internal_attempt == 3:
                        return "Failed", "TIS Internal Menu Failed"
            
            if not tis_internal_success: return "Failed", "TIS Internal Menu Failed"
            
            try:
                self.log("   📥 Fetching Available Years...")
                try:
                    dropdown_toggle = wait.until(EC.presence_of_element_located((By.ID, "dropdownMenuButton")))
                    driver.execute_script("arguments[0].click();", dropdown_toggle)
                    time.sleep(0.5)
                except: pass

                year_buttons = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//button[contains(@class, 'dropdown-item') and contains(text(), 'F.Y.')]") ))
                available_years = []
                for btn in year_buttons:
                    txt = btn.get_attribute("textContent").strip()
                    if txt and txt not in available_years:
                        available_years.append(txt)
                try: driver.execute_script("arguments[0].click();", dropdown_toggle)
                except: pass

                if not available_years: return "Failed", "No years found in AIS"

                if self.year_mode == "Current Year": self.current_user_selected_years = available_years[:1]
                elif self.year_mode == "Current and Last Year": self.current_user_selected_years = available_years[:2]
                elif self.year_mode == "Current and Last 2 Years": self.current_user_selected_years = available_years[:3]
                else:
                    self.log(f"   🛑 PAUSED: Found {len(available_years)} years. Waiting for you...")
                    self.user_selection_event.clear()
                    self.current_user_selected_years = None
                    self.app.trigger_year_selection(available_years, user_id, self.set_years_and_resume)
                    self.user_selection_event.wait()

                years_to_download = [y for y in self.current_user_selected_years if y in available_years]
                if not years_to_download: return "Warning", "No valid years selected"

                self.log(f"   ⬇️ Downloading {len(years_to_download)} Years (TIS)...")
                count = 0
                for year in years_to_download:
                    year_success = False
                    for year_attempt in range(1, 4):
                        try:
                            if year_attempt > 1:
                                self.log(f"     -> Retry {year_attempt}/3 for {year}...")
                                driver.refresh()
                                time.sleep(2)
                                ais_internal_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]") ))
                                driver.execute_script("arguments[0].click();", ais_internal_menu)
                                time.sleep(2)
                            else:
                                self.log(f"     -> Processing {year}...")
                            
                            
                            try:
                                dropdown_toggle = driver.find_element(By.ID, "dropdownMenuButton")
                                driver.execute_script("arguments[0].click();", dropdown_toggle)
                                time.sleep(0.5)
                            except: pass

                            year_xpath = f"//button[contains(@class, 'dropdown-item') and contains(text(), '{year}')]"
                            target_yr_btn = wait.until(EC.presence_of_element_located((By.XPATH, year_xpath)))
                            driver.execute_script("arguments[0].click();", target_yr_btn)
                            time.sleep(2)

                            tis_dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[contains(@title, 'Download TIS related documents') or contains(@alt, 'Download TIS related documents')]") ))
                            driver.execute_script("arguments[0].click();", tis_dl_icon)
                            time.sleep(1)
                            year_success = True
                            break
                        except Exception as e:
                            if year_attempt == 3:
                                self.log(f"        ⚠️ Failed to process {year} after 3 attempts")
                    
                    if not year_success:
                        continue
                    
                    # PDF Download with retry
                    pdf_download_success = False
                    for pdf_attempt in range(1, 4):
                        try:
                            if pdf_attempt > 1:
                                self.log(f"        ⚠️ PDF Download Retry {pdf_attempt}/3...")
                                time.sleep(2)
                            
                            clean_temp_files(download_folder, prefixes=("TIS_",))
                            modal_dl_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Download' and contains(@class, 'btn-outline-primary')]") ))
                            
                            modal_click_time = time.time()
                            driver.execute_script("arguments[0].click();", modal_dl_btn)
                            self.log("        Generating Document...")
                        
                            # --- INDIVIDUAL VS COMPANY LOGIC ---
                            direct_download = False
                            prefix = "TIS_"
                            for _ in range(8):
                                time.sleep(1)
                                for f in os.listdir(download_folder):
                                    f_path = os.path.join(download_folder, f)
                                    if os.path.isfile(f_path) and os.path.getmtime(f_path) >= modal_click_time - 2:
                                        if f.endswith(".crdownload") or f.endswith(".pdf"):
                                            direct_download = True
                                            break
                                if direct_download: break

                            if direct_download:
                                self.log("        ✅ Direct download detected.")
                                saved_path = wait_and_rename_file(download_folder, year, self.log, prefix=prefix, start_time=modal_click_time-2)
                                if saved_path:
                                    count += 1
                                    unlock_pdf(saved_path, user_id, dob, self.log)
                                    pdf_download_success = True
                                    break
                                else:
                                    self.log("        ❌ File capture failed.")
                                try:
                                    close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                    driver.execute_script("arguments[0].click();", close_btn)
                                except: pass
                            else:
                                self.log("        ℹ️ No direct download. Checking Activity History...")
                                try:
                                    close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                    driver.execute_script("arguments[0].click();", close_btn)
                                    time.sleep(0.5)
                                except: pass

                                history_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Go To Activity History']")))
                                driver.execute_script("arguments[0].click();", history_btn)
                                time.sleep(3)
                                
                                clean_temp_files(download_folder, prefixes=("TIS_",))
                                hist_click_time = time.time()
                                final_dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "(//img[@alt='Download'])[1]")))
                                driver.execute_script("arguments[0].click();", final_dl_icon)
                                self.log("        ✅ Export Triggered from History.")
                                
                                saved_path = wait_and_rename_file(download_folder, year, self.log, prefix=prefix, start_time=hist_click_time-2)
                                if saved_path:
                                    count += 1
                                    unlock_pdf(saved_path, user_id, dob, self.log)
                                    pdf_download_success = True
                                    break
                                else:
                                    self.log("        ❌ File capture failed.")
                                    
                        except Exception as e:
                            self.log(f"       ⚠️ Download Attempt {pdf_attempt} Failed: {str(e)[:30]}")
                            if pdf_attempt == 3:
                                self.log("       ❌ PDF Export failed after 3 attempts.")

                    # Return to AIS tab for next year
                    try:
                        ais_internal_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]") ))
                        driver.execute_script("arguments[0].click();", ais_internal_menu)
                        time.sleep(2)
                    except: pass

                return "Success", f"Downloaded {count} files"

            except Exception as e:
                return "Failed", f"TIS Portal Error: {str(e)[:20]}"

        except Exception as e:
            return "Failed", "Browser Crash"
        finally:
            if driver: driver.quit()


# ============================================================
#  WORKER 4: FILED RETURN REPORT WORKER
# ============================================================
class FiledReturnWorker:
    def __init__(self, app_instance, excel_path, year_mode):
        self.app = app_instance
        self.excel_path = excel_path
        self.year_mode = year_mode
        self.keep_running = True
        self.report_data = []
        self.user_selection_event = threading.Event()
        self.current_user_selected_years = None

    def log(self, message):
        self.app.update_log_safe_filed(message)

    def set_years_and_resume(self, selected_list):
        self.current_user_selected_years = selected_list
        self.user_selection_event.set()

    def run(self):
        self.log("🚀 INITIALIZING FILED RETURN REPORT ENGINE...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        
        try:
            # Read the Excel file
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)
            
            if not user_col or not pass_col:
                self.log("❌ ERROR: Headers missing. Need 'PAN' and 'Password'.")
                self.app.process_finished_safe_filed("Failed: Column Header Error")
                return
            
            self.log(f"✅ Mapped Columns -> ID: '{user_col}', Pass: '{pass_col}', DOB: '{dob_col}'")
            total_users = len(df)
            
            # Process each user
            for index, row in df.iterrows():
                if not self.keep_running:
                    self.log("🛑 Process Stopped by User.")
                    break
                
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None
                
                self.app.update_progress_safe_filed((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                status, reason = self.process_single_user(user_id, password, dob)
                
                self.log(f"   📊 Result: {status} - {reason}")
                self.log("-" * 40)
            
            # Generate the report
            self.generate_report()
            
            self.app.update_progress_safe_filed(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe_filed("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe_filed("Critical Error Occurred")

    def process_single_user(self, user_id, password, dob):
        driver = None
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument("--disable-blink-features=AutomationControlled")

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            
            # Set aggressive timeouts to prevent hanging
            driver.set_page_load_timeout(30)  # 30 seconds for page load
            driver.set_script_timeout(30)     # 30 seconds for script execution
            driver.implicitly_wait(10)        # 10 seconds for element search
            
            wait = WebDriverWait(driver, 20)
            actions = ActionChains(driver)

            # 1. LOGIN WITH COMPREHENSIVE RETRY
            login_success = False
            for login_attempt in range(1, 4):
                if login_success: break
                if login_attempt > 1:
                    self.log(f"   ⚠️ Login Retry {login_attempt}/3...")
                    try:
                        driver.delete_all_cookies()
                        driver.refresh()
                    except: pass
                    time.sleep(3)

                try:
                    self.log("   🌐 Opening Portal...")
                    try:
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")
                        time.sleep(2)
                    except TimeoutException:
                        self.log("   ⚠️ Page load timeout. Retrying...")
                        continue
                    except Exception as e:
                        self.log(f"   ⚠️ Page load error: {str(e)[:30]}. Retrying...")
                        continue
                    
                    try: driver.switch_to.alert.accept(); time.sleep(1)
                    except: pass

                    # Step 1: Enter PAN with retry
                    pan_entered = False
                    for pan_retry in range(3):
                        try:
                            pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                            pan_field.clear()
                            pan_field.send_keys(user_id)
                            pan_entered = True
                            break
                        except Exception as e:
                            if pan_retry == 2:
                                self.log(f"   ⚠️ Failed to enter PAN after 3 tries")
                                raise
                            time.sleep(1)
                    
                    if not pan_entered:
                        continue
                    
                    time.sleep(0.5)
                    
                    # Step 2: Click Continue button with retry
                    for cont_retry in range(3):
                        try:
                            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                            driver.execute_script("arguments[0].click();", continue_btn)
                            break
                        except Exception as e:
                            if cont_retry == 2:
                                self.log(f"   ⚠️ Failed to click Continue after 3 tries")
                                raise
                            time.sleep(1)
                    
                    time.sleep(1.5)
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN"

                    # Step 3: Enter Password with retry
                    for pwd_retry in range(3):
                        try:
                            pwd_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                            pwd_field.clear()
                            pwd_field.send_keys(password)
                            break
                        except Exception as e:
                            if pwd_retry == 2:
                                self.log(f"   ⚠️ Failed to enter password after 3 tries")
                                raise
                            time.sleep(1)
                    
                    # Show password checkbox
                    try: 
                        driver.execute_script("document.getElementById('passwordCheckBox-input').click();")
                        time.sleep(0.3)
                    except: pass
                    
                    self.log("   ⏳ Waiting for security check (3s)...")
                    time.sleep(3.5)
                    
                    # Step 4: Submit login with retry
                    for submit_retry in range(3):
                        try:
                            driver.execute_script("document.querySelector('button.large-button-primary').click();")
                            break
                        except Exception as e:
                            if submit_retry == 2:
                                self.log(f"   ⚠️ Failed to submit login after 3 tries")
                                raise
                            time.sleep(1)

                    # Step 5: Wait for successful login
                    for _ in range(20):
                        time.sleep(1)
                        try:
                            if driver.find_elements(By.ID, "e-File"):
                                self.log("   ✅ Login Successful!")
                                login_success = True; break
                        except: pass
                        
                        if "Invalid Password" in driver.page_source: return "Failed", "Invalid Password"
                        
                        try:
                            dual = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual and dual[0].is_displayed():
                                driver.execute_script("arguments[0].click();", dual[0])
                                time.sleep(2)
                        except: pass
                    if login_success: break
                except Exception as e:
                    self.log(f"   ⚠️ Login Error: {str(e)[:50]}")
                    if login_attempt < 3:
                        time.sleep(2)

            if not login_success: return "Failed", "Login Timeout"

            # 2. NAVIGATE TO VIEW FILED RETURNS (with retry logic)
            self.log("   🚀 Navigating to View Filed Returns...")
            nav_success = False
            for nav_attempt in range(1, 4):
                try:
                    if nav_attempt > 1:
                        self.log(f"   ⚠️ Navigation Retry {nav_attempt}/3...")
                        time.sleep(2)
                    
                    # Step 1: Click e-File menu
                    self.log("   📂 Clicking e-File menu...")
                    efile = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[@class='mdc-button__label' and contains(text(), 'e-File')]"))
                    )
                    driver.execute_script("arguments[0].click();", efile)
                    time.sleep(1)
                    
                    # Step 2: Hover over Income Tax Returns
                    self.log("   📋 Hovering over Income Tax Returns...")
                    itr_menu = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//button[contains(@class, 'mat-mdc-menu-item')]//span[contains(text(), 'Income Tax Returns')]"))
                    )
                    actions.move_to_element(itr_menu).perform()
                    time.sleep(0.5)
                    
                    # Step 3: Click View Filed Returns
                    self.log("   🔍 Clicking View Filed Returns...")
                    view_filed = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'mat-mdc-menu-item')]//span[contains(text(), 'View Filed Returns')]"))
                    )
                    driver.execute_script("arguments[0].click();", view_filed)
                    time.sleep(3)
                    
                    self.log("   ✅ Navigation Successful!")
                    nav_success = True
                    break
                    
                except Exception as e:
                    self.log(f"   ⚠️ Nav Attempt {nav_attempt} Failed: {str(e)[:40]}")
                    if nav_attempt == 3:
                        return "Failed", "Navigation to View Filed Returns Failed"
            
            if not nav_success: return "Failed", "Menu Navigation Error"

            # 3. DATA EXTRACTION FROM FILED RETURNS PAGE
            self.log("   📊 Starting data extraction...")
            time.sleep(2)
            
            try:
                # Check filing count
                try:
                    filing_count_elem = driver.find_element(By.CLASS_NAME, "filingCount")
                    filing_count = filing_count_elem.text
                    self.log(f"   📄 {filing_count}")
                except:
                    self.log("   ⚠️ Could not find filing count")
                
                # Click pagination dropdown to show maximum records
                self.log("   🔽 Setting pagination to show all records...")
                try:
                    # Click the specific pagination select with id="paginatorselect"
                    pagination_field = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "paginatorselect"))
                    )
                    pagination_select = pagination_field.find_element(By.TAG_NAME, "mat-select")
                    driver.execute_script("arguments[0].click();", pagination_select)
                    time.sleep(1)
                    
                    # Get all options and click the highest one
                    options = driver.find_elements(By.XPATH, "//mat-option[@role='option']")
                    if options:
                        # Click the last option (highest number)
                        driver.execute_script("arguments[0].click();", options[-1])
                        self.log(f"   ✅ Set pagination to maximum records")
                        time.sleep(2)
                except Exception as e:
                    self.log(f"   ⚠️ Pagination selection skipped: {str(e)[:30]}")
                
                # Scrape all assessment years available
                self.log("   🔍 Extracting Assessment Years...")
                year_elements = driver.find_elements(By.XPATH, "//mat-label[@class='contentHeadingText']")
                available_years = []
                for elem in year_elements:
                    year_text = elem.text.strip()
                    if "A.Y." in year_text:
                        available_years.append(year_text)
                
                if not available_years:
                    self.log("   ⚠️ No filed returns found")
                    return "Success", "No Filed Returns Found"
                
                self.log(f"   📋 Found {len(available_years)} Assessment Years: {', '.join(available_years)}")
                
                # Apply year selection logic based on year_mode
                if self.year_mode == "Current Year":
                    self.current_user_selected_years = available_years[:1]
                elif self.year_mode == "Current and Last Year":
                    self.current_user_selected_years = available_years[:2]
                elif self.year_mode == "Current and Last 2 Years":
                    self.current_user_selected_years = available_years[:3]
                else:  # Manual Selection
                    self.log(f"   🛑 PAUSED: Found {len(available_years)} years. Waiting for selection...")
                    self.user_selection_event.clear()
                    self.current_user_selected_years = None
                    self.app.trigger_year_selection(available_years, user_id, self.set_years_and_resume)
                    self.user_selection_event.wait()
                
                years_to_extract = [y for y in self.current_user_selected_years if y in available_years]
                if not years_to_extract:
                    self.log("   ⚠️ No matching years selected")
                    return "Success", "No Matching Years"
                
                self.log(f"   ⬇️ Extracting data for {len(years_to_extract)} years: {', '.join(years_to_extract)}")
                
                # Scrape data from all cards
                self.log("   📥 Extracting return details...")
                cards = driver.find_elements(By.XPATH, "//mat-card[contains(@class, 'contextBox')]")
                
                extracted_data = []
                for idx, card in enumerate(cards):
                    try:
                        # Assessment Year
                        ay = card.find_element(By.CLASS_NAME, "contentHeadingText").text.strip()
                        
                        # Skip if this year is not in the selected years
                        if ay not in years_to_extract:
                            continue
                        
                        # Filing Type
                        filing_type = card.find_element(By.CLASS_NAME, "leftSideVal").text.strip()
                        
                        # First status and date from stepper
                        first_status = "N/A"
                        first_date = "N/A"
                        try:
                            status_divs = card.find_elements(By.CLASS_NAME, "matStepStatus")
                            date_divs = card.find_elements(By.CLASS_NAME, "matStepDate")
                            if status_divs:
                                first_status = status_divs[0].text.strip()
                            if date_divs:
                                first_date = date_divs[0].text.strip()
                        except:
                            pass
                        
                        # ITR Type
                        itr_type = "N/A"
                        # Acknowledgement Number
                        ack_no = "N/A"
                        # Filed By
                        filed_by = "N/A"
                        # Filing Date
                        filing_date = "N/A"
                        # Filing Section
                        filing_section = "N/A"
                        
                        try:
                            right_labels = card.find_elements(By.CLASS_NAME, "rightsideLabel")
                            right_values = card.find_elements(By.CLASS_NAME, "fieldVal")
                            
                            for i, label_elem in enumerate(right_labels):
                                label = label_elem.text.strip().lower()
                                if i < len(right_values):
                                    value = right_values[i].text.strip()
                                    if "itr" in label:
                                        itr_type = value
                                    elif "acknowledgement" in label:
                                        ack_no = value
                                    elif "filed by" in label:
                                        filed_by = value
                                    elif "filing date" in label:
                                        filing_date = value
                                    elif "filing section" in label:
                                        filing_section = value
                        except:
                            pass
                        
                        extracted_data.append({
                            "PAN": user_id,
                            "Assessment Year": ay,
                            "Filing Type": filing_type,
                            "Current Status": first_status,
                            "Status Date": first_date,
                            "ITR Type": itr_type,
                            "Acknowledgement No": ack_no,
                            "Filed By": filed_by,
                            "Filing Date": filing_date,
                            "Filing Section": filing_section
                        })
                        
                        self.log(f"   ✅ Extracted: {ay} - {first_status}")
                        
                    except Exception as e:
                        self.log(f"   ⚠️ Error extracting card {idx+1}: {str(e)[:30]}")
                
                # Store extracted data for report
                if extracted_data:
                    for data in extracted_data:
                        self.report_data.append(data)
                    self.log(f"   ✅ Extracted {len(extracted_data)} return records")
                    return "Success", f"Extracted {len(extracted_data)} returns"
                else:
                    return "Success", "No data extracted"
                    
            except Exception as e:
                self.log(f"   ❌ Data extraction error: {str(e)[:50]}")
                return "Failed", f"Extraction Error: {str(e)[:20]}"

        except Exception as e:
            return "Failed", f"Browser Error: {str(e)[:30]}"
        finally:
            if driver:
                driver.quit()

    def generate_report(self):
        try:
            if not self.report_data:
                self.log("⚠️ No data to generate report")
                return
            
            df_report = pd.DataFrame(self.report_data)
            
            # Generate filename based on timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Filed_Return_Report_{timestamp}.xlsx"
            
            self.log(f"📝 Generating report: {filename}")
            
            # Create Excel writer with formatting
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df_report.to_excel(writer, index=False, sheet_name='Filed Returns')
                
                # Get the worksheet
                worksheet = writer.sheets['Filed Returns']
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            self.log(f"✅ Report saved successfully: {filename}")
            self.log(f"📊 Total records in report: {len(df_report)}")
            
        except Exception as e:
            self.log(f"❌ Report generation error: {str(e)}")


# ============================================================
#  WORKER 5: COMBINED WORKER (26AS + AIS + TIS per PAN)
# ============================================================
class CombinedWorker:
    def __init__(self, app_instance, excel_path, year_mode):
        self.app = app_instance
        self.excel_path = excel_path
        self.year_mode = year_mode
        self.keep_running = True
        self.report_data = []

    def log(self, message):
        try: self.app.update_log_safe_26as(message)
        except: pass
        try: self.app.update_log_safe_ais(message)
        except: pass
        try: self.app.update_log_safe_tis(message)
        except: pass

    def run(self):
        self.log("🚀 INITIALIZING COMBINED ENGINE (26AS + AIS + TIS)...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        try:
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)
            if not user_col or not pass_col:
                self.log("❌ ERROR: Headers missing.")
                # Reset all button states
                self.app.after(0, lambda: self.app.btn_start_26as.configure(state="normal", text="START 26AS DOWNLOAD", fg_color="#1f538d"))
                self.app.after(0, lambda: self.app.btn_start_ais.configure(state="normal", text="START AIS DOWNLOAD", fg_color="#1f538d"))
                self.app.after(0, lambda: self.app.btn_start_tis.configure(state="normal", text="START TIS DOWNLOAD", fg_color="#1f538d"))
                return

            total_users = len(df)
            for index, row in df.iterrows():
                if not self.keep_running: break
                
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None
                
                self.app.update_progress_safe_26as((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                base_dir = os.getcwd()
                combined_root = os.path.join(base_dir, "Combined_Downloads")
                if not os.path.exists(combined_root): os.makedirs(combined_root, exist_ok=True)
                
                user_folder = create_unique_folder(combined_root, user_id)

                folder_26as = os.path.join(user_folder, "1) 26AS")
                folder_ais = os.path.join(user_folder, "2) AIS")
                folder_tis = os.path.join(user_folder, "3) TIS")
                for p in (folder_26as, folder_ais, folder_tis):
                    if not os.path.exists(p): os.makedirs(p, exist_ok=True)

                w26 = Tax26ASWorker(self.app, self.excel_path, self.year_mode)
                wais = AISWorker(self.app, self.excel_path, self.year_mode)
                wtis = TISWorker(self.app, self.excel_path, self.year_mode)

                try:
                    status26, reason26 = w26.process_single_user(user_id, password, dob, folder_26as)
                    self.log(f"   26AS -> {status26}: {reason26}")
                except Exception as e:
                    status26, reason26 = "Failed", str(e)
                    self.log(f"   26AS -> Exception: {e}")

                try:
                    statusAIS, reasonAIS = wais.process_single_user(user_id, password, dob, folder_ais)
                    self.log(f"   AIS -> {statusAIS}: {reasonAIS}")
                except Exception as e:
                    statusAIS, reasonAIS = "Failed", str(e)
                    self.log(f"   AIS -> Exception: {e}")

                try:
                    statusTIS, reasonTIS = wtis.process_single_user(user_id, password, dob, folder_tis)
                    self.log(f"   TIS -> {statusTIS}: {reasonTIS}")
                except Exception as e:
                    statusTIS, reasonTIS = "Failed", str(e)
                    self.log(f"   TIS -> Exception: {e}")

                self.report_data.append({
                    "PAN": user_id,
                    "26AS Status": status26, "26AS Details": reason26,
                    "AIS Status": statusAIS, "AIS Details": reasonAIS,
                    "TIS Status": statusTIS, "TIS Details": reasonTIS,
                    "Folder Saved": os.path.basename(user_folder),
                    "Timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
                self.log("-" * 40)

            try:
                if self.report_data:
                    df_report = pd.DataFrame(self.report_data)
                    filename = f"Combined_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    df_report.to_excel(filename, index=False)
                    self.log(f"📄 Combined report saved: {filename}")
            except Exception as e:
                self.log(f"⚠️ Failed to save combined report: {e}")

            # Reset all button states
            self.app.after(0, lambda: self.app.btn_start_26as.configure(state="normal", text="START 26AS DOWNLOAD", fg_color="#1f538d"))
            self.app.after(0, lambda: self.app.btn_start_ais.configure(state="normal", text="START AIS DOWNLOAD", fg_color="#1f538d"))
            self.app.after(0, lambda: self.app.btn_start_tis.configure(state="normal", text="START TIS DOWNLOAD", fg_color="#1f538d"))
            self.app.after(0, lambda: messagebox.showinfo("Done", "Combined download finished"))

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {e}")
            # Reset buttons even on error
            self.app.after(0, lambda: self.app.btn_start_26as.configure(state="normal", text="START 26AS DOWNLOAD", fg_color="#1f538d"))
            self.app.after(0, lambda: self.app.btn_start_ais.configure(state="normal", text="START AIS DOWNLOAD", fg_color="#1f538d"))
            self.app.after(0, lambda: self.app.btn_start_tis.configure(state="normal", text="START TIS DOWNLOAD", fg_color="#1f538d"))

# ============================================================
#  MAIN APP GUI
# ============================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Suite Pro")
        self.geometry("900x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.worker = None

        # --- Header ---
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.title_label = ctk.CTkLabel(self.header_frame, text="INCOME TAX AUTOMATION SUITE", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(side="left")
        
        # --- Main Tab View ---
        self.tabview = ctk.CTkTabview(self, width=860)
        self.tabview.add("26AS")
        self.tabview.add("AIS")
        self.tabview.add("TIS")
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

        # ==========================================
        #  DIRECT TABS (26AS, AIS, TIS)
        # ==========================================
        self.tab_26as = self.tabview.tab("26AS")
        self.tab_ais = self.tabview.tab("AIS")
        self.tab_tis = self.tabview.tab("TIS")

        self.chk_download_all_var = ctk.StringVar(value="off")
        self._build_26as_ui()
        self._build_ais_ui()
        self._build_tis_ui()

    # --- UI BUILDERS ---
    def _build_26as_ui(self):
        self.excel_file_path_26as = ""
        ctk.CTkCheckBox(self.tab_26as, text="Download all Three documents (26AS, AIS, TIS)", variable=self.chk_download_all_var, onvalue="on", offvalue="off").pack(anchor='nw', padx=10, pady=(6,4))
        self.config_26as = ctk.CTkFrame(self.tab_26as)
        self.config_26as.pack(fill="x", padx=10, pady=(2, 5))

        ctk.CTkLabel(self.config_26as, text="1. CREDENTIALS SOURCE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        f_frame = ctk.CTkFrame(self.config_26as, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file_26as = ctk.CTkEntry(f_frame, placeholder_text="Excel File (Headers: PAN, Password, DOB)...")
        self.entry_file_26as.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("26as"), width=100).pack(side="right")

        pref_frame = ctk.CTkFrame(self.config_26as, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(pref_frame, text="Download Return:", text_color="gray").pack(side="left", padx=(0, 10))
        self.combo_years_26as = ctk.CTkComboBox(pref_frame, values=["Current Year", "Current and Last Year", "Current and Last 2 Years", "Manual Selection (Popup)"], width=250, state="readonly")
        self.combo_years_26as.set("Current Year")
        self.combo_years_26as.pack(side="left")

        self.log_frame_26as = ctk.CTkFrame(self.tab_26as)
        self.log_frame_26as.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.log_frame_26as.grid_rowconfigure(1, weight=1)
        self.log_frame_26as.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.log_frame_26as, text="3. LIVE LOG", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(5, 5))
        self.log_box_26as = ctk.CTkTextbox(self.log_frame_26as, font=("Consolas", 12), activate_scrollbars=True)
        self.log_box_26as.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box_26as.configure(state="disabled")
        
        self.progress_26as = ctk.CTkProgressBar(self.log_frame_26as, mode="determinate")
        self.progress_26as.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_26as.set(0)

        self.btn_start_26as = ctk.CTkButton(self.tab_26as, text="START 26AS DOWNLOAD", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=lambda: self.start_process("26as"))
        self.btn_start_26as.pack(fill="x", padx=20, pady=(0, 20))

    def _build_ais_ui(self):
        self.excel_file_path_ais = ""
        ctk.CTkCheckBox(self.tab_ais, text="Download all Three documents (26AS, AIS, TIS)", variable=self.chk_download_all_var, onvalue="on", offvalue="off").pack(anchor='nw', padx=10, pady=(6,4))
        self.config_ais = ctk.CTkFrame(self.tab_ais)
        self.config_ais.pack(fill="x", padx=10, pady=(2, 5))

        ctk.CTkLabel(self.config_ais, text="1. CREDENTIALS SOURCE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        f_frame = ctk.CTkFrame(self.config_ais, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file_ais = ctk.CTkEntry(f_frame, placeholder_text="Excel File (Headers: PAN, Password, DOB)...")
        self.entry_file_ais.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("ais"), width=100).pack(side="right")

        pref_frame = ctk.CTkFrame(self.config_ais, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(pref_frame, text="Download Return:", text_color="gray").pack(side="left", padx=(0, 10))
        self.combo_years_ais = ctk.CTkComboBox(pref_frame, values=["Current Year", "Current and Last Year", "Current and Last 2 Years", "Manual Selection (Popup)"], width=250, state="readonly")
        self.combo_years_ais.set("Current Year")
        self.combo_years_ais.pack(side="left")

        self.log_frame_ais = ctk.CTkFrame(self.tab_ais)
        self.log_frame_ais.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.log_frame_ais.grid_rowconfigure(1, weight=1)
        self.log_frame_ais.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.log_frame_ais, text="3. LIVE LOG", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(5, 5))
        self.log_box_ais = ctk.CTkTextbox(self.log_frame_ais, font=("Consolas", 12), activate_scrollbars=True)
        self.log_box_ais.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box_ais.configure(state="disabled")
        
        self.progress_ais = ctk.CTkProgressBar(self.log_frame_ais, mode="determinate")
        self.progress_ais.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_ais.set(0)

        self.btn_start_ais = ctk.CTkButton(self.tab_ais, text="START AIS DOWNLOAD", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=lambda: self.start_process("ais"))
        self.btn_start_ais.pack(fill="x", padx=20, pady=(0, 20))

    def _build_tis_ui(self):
        self.excel_file_path_tis = ""
        ctk.CTkCheckBox(self.tab_tis, text="Download all Three documents (26AS, AIS, TIS)", variable=self.chk_download_all_var, onvalue="on", offvalue="off").pack(anchor='nw', padx=10, pady=(6,4))
        self.config_tis = ctk.CTkFrame(self.tab_tis)
        self.config_tis.pack(fill="x", padx=10, pady=(2, 5))

        ctk.CTkLabel(self.config_tis, text="1. CREDENTIALS SOURCE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        f_frame = ctk.CTkFrame(self.config_tis, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file_tis = ctk.CTkEntry(f_frame, placeholder_text="Excel File (Headers: PAN, Password, DOB)...")
        self.entry_file_tis.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("tis"), width=100).pack(side="right")

        pref_frame = ctk.CTkFrame(self.config_tis, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(pref_frame, text="Download Return:", text_color="gray").pack(side="left", padx=(0, 10))
        self.combo_years_tis = ctk.CTkComboBox(pref_frame, values=["Current Year", "Current and Last Year", "Current and Last 2 Years", "Manual Selection (Popup)"], width=250, state="readonly")
        self.combo_years_tis.set("Current Year")
        self.combo_years_tis.pack(side="left")

        self.log_frame_tis = ctk.CTkFrame(self.tab_tis)
        self.log_frame_tis.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.log_frame_tis.grid_rowconfigure(1, weight=1)
        self.log_frame_tis.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.log_frame_tis, text="3. LIVE LOG", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(5, 5))
        self.log_box_tis = ctk.CTkTextbox(self.log_frame_tis, font=("Consolas", 12), activate_scrollbars=True)
        self.log_box_tis.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box_tis.configure(state="disabled")

        self.progress_tis = ctk.CTkProgressBar(self.log_frame_tis, mode="determinate")
        self.progress_tis.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_tis.set(0)

        self.btn_start_tis = ctk.CTkButton(self.tab_tis, text="START TIS DOWNLOAD", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=lambda: self.start_process("tis"))
        self.btn_start_tis.pack(fill="x", padx=20, pady=(0, 20))

    # --- GUI Handlers ---
    def trigger_year_selection(self, years_list, user_id, callback):
        self.after(0, lambda: self._show_popup(years_list, user_id, callback))

    def _show_popup(self, years_list, user_id, callback):
        YearSelectionPopup(self, years_list, user_id, callback)

    def browse_file(self, mode):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            if mode == "26as":
                self.excel_file_path_26as = filename
                self.entry_file_26as.delete(0, "end")
                self.entry_file_26as.insert(0, filename)
                self.log_to_gui_26as(f"File Loaded: {os.path.basename(filename)}")
            elif mode == "ais":
                self.excel_file_path_ais = filename
                self.entry_file_ais.delete(0, "end")
                self.entry_file_ais.insert(0, filename)
                self.log_to_gui_ais(f"File Loaded: {os.path.basename(filename)}")
            elif mode == "tis":
                self.excel_file_path_tis = filename
                self.entry_file_tis.delete(0, "end")
                self.entry_file_tis.insert(0, filename)
                self.log_to_gui_tis(f"File Loaded: {os.path.basename(filename)}")

    def start_process(self, mode):
        if getattr(self, 'chk_download_all_var', None) and self.chk_download_all_var.get() == "on":
            excel_path = None
            year_mode = None
            if mode == "26as":
                if not self.excel_file_path_26as: return messagebox.showwarning("Error", "Select file first")
                excel_path = self.excel_file_path_26as; year_mode = self.combo_years_26as.get()
            elif mode == "ais":
                if not self.excel_file_path_ais: return messagebox.showwarning("Error", "Select file first")
                excel_path = self.excel_file_path_ais; year_mode = self.combo_years_ais.get()
            elif mode == "tis":
                if not self.excel_file_path_tis: return messagebox.showwarning("Error", "Select file first")
                excel_path = self.excel_file_path_tis; year_mode = self.combo_years_tis.get()

            try: self.btn_start_26as.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            except: pass
            try: self.btn_start_ais.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            except: pass
            try: self.btn_start_tis.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            except: pass

            self.worker = CombinedWorker(self, excel_path, year_mode)
            threading.Thread(target=self.worker.run, daemon=True).start()
            return

        if mode == "26as":
            if not self.excel_file_path_26as: return messagebox.showwarning("Error", "Select file first")
            self.btn_start_26as.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            self.progress_26as.set(0)
            self.worker = Tax26ASWorker(self, self.excel_file_path_26as, self.combo_years_26as.get())
            threading.Thread(target=self.worker.run, daemon=True).start()
        elif mode == "ais":
            if not self.excel_file_path_ais: return messagebox.showwarning("Error", "Select file first")
            self.btn_start_ais.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            self.progress_ais.set(0)
            self.worker = AISWorker(self, self.excel_file_path_ais, self.combo_years_ais.get())
            threading.Thread(target=self.worker.run, daemon=True).start()
        elif mode == "tis":
            if not self.excel_file_path_tis: return messagebox.showwarning("Error", "Select file first")
            self.btn_start_tis.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            self.progress_tis.set(0)
            self.worker = TISWorker(self, self.excel_file_path_tis, self.combo_years_tis.get())
            threading.Thread(target=self.worker.run, daemon=True).start()

    # --- 26AS SAFE UPDATERS ---
    def log_to_gui_26as(self, msg):
        self.log_box_26as.configure(state="normal")
        self.log_box_26as.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box_26as.see("end")
        self.log_box_26as.configure(state="disabled")

    def update_log_safe_26as(self, msg): self.after(0, lambda: self.log_to_gui_26as(msg))
    def update_progress_safe_26as(self, val): self.after(0, lambda: self.progress_26as.set(val))
    def process_finished_safe_26as(self, msg):
        self.after(0, lambda: self.log_to_gui_26as(f"\nSTATUS: {msg}"))
        self.after(0, lambda: self.btn_start_26as.configure(state="normal", text="START 26AS DOWNLOAD", fg_color="#1f538d"))
        self.after(0, lambda: messagebox.showinfo("Done", msg))

    # --- AIS SAFE UPDATERS ---
    def log_to_gui_ais(self, msg):
        self.log_box_ais.configure(state="normal")
        self.log_box_ais.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box_ais.see("end")
        self.log_box_ais.configure(state="disabled")

    def update_log_safe_ais(self, msg): self.after(0, lambda: self.log_to_gui_ais(msg))
    def update_progress_safe_ais(self, val): self.after(0, lambda: self.progress_ais.set(val))
    def process_finished_safe_ais(self, msg):
        self.after(0, lambda: self.log_to_gui_ais(f"\nSTATUS: {msg}"))
        self.after(0, lambda: self.btn_start_ais.configure(state="normal", text="START AIS DOWNLOAD", fg_color="#1f538d"))
        self.after(0, lambda: messagebox.showinfo("Done", msg))

    # --- TIS SAFE UPDATERS ---
    def log_to_gui_tis(self, msg):
        self.log_box_tis.configure(state="normal")
        self.log_box_tis.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box_tis.see("end")
        self.log_box_tis.configure(state="disabled")

    def update_log_safe_tis(self, msg): self.after(0, lambda: self.log_to_gui_tis(msg))
    def update_progress_safe_tis(self, val): self.after(0, lambda: self.progress_tis.set(val))
    def process_finished_safe_tis(self, msg):
        self.after(0, lambda: self.log_to_gui_tis(f"\nSTATUS: {msg}"))
        self.after(0, lambda: self.btn_start_tis.configure(state="normal", text="START TIS DOWNLOAD", fg_color="#1f538d"))
        self.after(0, lambda: messagebox.showinfo("Done", msg))


if __name__ == "__main__":
    app = App()
    app.mainloop()