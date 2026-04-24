import threading
import time
import os
import shutil
import tempfile
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

# Assessment Year options (update as needed)
ASSESSMENT_YEAR_OPTIONS = [
    "2027-2028",
    "2026-2027",
    "2025-2026",
    "2024-2025",
    "2023-2024",
    "2022-2023",
    "Manual Selection (Popup)",
]

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
    # Sanitize folder name to remove characters invalid on Windows
    folder_name = re.sub(r'[<>:"/\\|?*]', '_', folder_name).strip()
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

def get_taxpayer_name(driver, fallback=""):
    """
    Extracts the logged-in taxpayer's name from the Income Tax portal header.
    Uses JavaScript (innerText) as the primary method since Angular apps often
    render content after Selenium's .text has already been read.
    Returns a sanitized name string, or fallback if extraction fails.
    """
    strategies = [
        # Strategy 1: JavaScript innerText on the userNameVal span (most reliable for Angular)
        lambda d: d.execute_script(
            "var el = document.querySelector('.userNameVal span:first-child'); "
            "return el ? el.innerText.trim() : '';"
        ),
        # Strategy 2: JS querySelector on the button containing the name
        lambda d: d.execute_script(
            "var el = document.querySelector('.profileMenubtn .userNameVal span'); "
            "return el ? el.innerText.trim() : '';"
        ),
        # Strategy 3: JS - grab all spans inside userNameVal and return first non-empty
        lambda d: next(
            (s for s in (
                d.execute_script(
                    "var spans = document.querySelectorAll('.userNameVal span'); "
                    "return Array.from(spans).map(function(s){return s.innerText.trim();});"
                ) or []
            ) if s and not s.lower() in ['expand_more', '']),
            ''
        ),
        # Strategy 4: Selenium XPath - span directly inside userNameVal
        lambda d: d.find_element(By.XPATH, "//span[contains(@class,'userNameVal')]/span[1]").text.strip(),
        # Strategy 5: Selenium - the button that shows the name tag
        lambda d: d.find_element(By.XPATH, "//button[contains(@class,'profileMenubtn')]//span[contains(@class,'userNameVal')]//span[1]").text.strip(),
    ]
    for strategy in strategies:
        try:
            result = strategy(driver)
            if result and len(result) > 1 and result.lower() not in ['expand_more']:
                # Final sanitize: remove any leftover icon text
                result = result.split('\n')[0].strip()
                if result:
                    return result
        except:
            continue
    return fallback

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

def wait_and_rename_file(folder, year_label, logger, prefix="", start_time=None, taxpayer_name=None):
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
        if taxpayer_name:
            base_target = f"{taxpayer_name}-{prefix}{safe_year}.pdf"
        else:
            base_target = f"{prefix}{safe_year}.pdf"
        new_name = os.path.join(folder, base_target)

        if os.path.abspath(newest_file) == os.path.abspath(new_name):
            logger(f"        📄 File already named: {base_target}")
            return new_name

        if os.path.exists(new_name):
            i = 1
            while True:
                if taxpayer_name:
                    candidate_name = os.path.join(folder, f"{taxpayer_name}-{prefix}{safe_year} ({i}).pdf")
                else:
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
                download_root = os.path.join(base_dir, "Income Tax Downloaded", "26 AS")

                status, reason, final_path = self.process_single_user(user_id, password, dob, download_root)
                
                self.report_data.append({
                    "PAN": user_id, "Status": status, "Details": reason,
                    "Folder Saved": os.path.basename(final_path) if final_path else user_id,
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
            report_dir = os.path.join(os.getcwd(), "Income Tax Downloaded", "reports")
            os.makedirs(report_dir, exist_ok=True)
            report_path = os.path.join(report_dir, filename)
            df_report.to_excel(report_path, index=False)
            self.log(f"📄 Report saved: {report_path}")
        except Exception as e:
            self.log(f"❌ Failed to save report: {e}")

    def process_single_user(self, user_id, password, dob, download_root):
        driver = None
        download_folder = tempfile.gettempdir()  # temporary path until name is known
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
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
                        
                        if "Invalid Password" in driver.page_source: return "Failed", "Invalid Password", download_folder
                        
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

            if not login_success: return "Failed", "Login Timeout", download_folder

            # Extract taxpayer name from dashboard header and create NAME_PAN folder
            name_from_header = get_taxpayer_name(driver, fallback=user_id)
            if name_from_header != user_id:
                self.log(f"   👤 Taxpayer Name: {name_from_header}")
            else:
                self.log("   ⚠️ Name not found in header; using PAN as folder name.")

            # Create the proper NAME_PAN folder and redirect Chrome downloads via CDP
            folder_name = f"{name_from_header}_{user_id}"
            download_folder = create_unique_folder(download_root, folder_name)
            self.log(f"   📁 Download folder: {os.path.basename(download_folder)}")
            try:
                driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                    'behavior': 'allow',
                    'downloadPath': download_folder
                })
            except Exception as cdp_e:
                self.log(f"   ⚠️ CDP redirect failed ({str(cdp_e)[:30]}); folder still created.")

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

                if self.year_mode in ASSESSMENT_YEAR_OPTIONS and "-" in self.year_mode and self.year_mode.split("-")[0].isdigit():
                    # Match absolute year like "2024-2025" to "2024-25"
                    start_yr, end_yr = self.year_mode.split("-")
                    target_ay = f"{start_yr}-{end_yr[2:]}"
                    self.current_user_selected_years = [target_ay] if target_ay in available_years else []

                else:
                    self.user_selection_event.wait()

                years_to_download = [y for y in self.current_user_selected_years if y in available_years]
                if not years_to_download: return "Warning", "No valid years selected"

                self.log(f"   ⬇️ Downloading {len(years_to_download)} Years...")
                count = 0
                for year in years_to_download:
                    # Create subfolder for the year (financial year wise)
                    safe_year_folder = year.replace('/', '-').replace(' ', '_').strip()
                    year_folder_path = os.path.join(download_folder, safe_year_folder)
                    os.makedirs(year_folder_path, exist_ok=True)
                    
                    # Update download behavior for this specific year folder
                    try:
                        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                            'behavior': 'allow',
                            'downloadPath': year_folder_path
                        })
                    except: pass

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
                            
                            clean_temp_files(year_folder_path)

                            dl_click_time = time.time()
                            pdf_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "pdfBtn")))
                            driver.execute_script("arguments[0].click();", pdf_btn)
                            self.log("        ✅ Export Triggered.")
                            
                            saved_path = wait_and_rename_file(year_folder_path, year, self.log, prefix="", start_time=dl_click_time, taxpayer_name=name_from_header)
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

                return "Success", f"Downloaded {count} files", download_folder

            except Exception as e: return "Failed", f"TRACES Error: {str(e)[:20]}", download_folder
        except Exception: return "Failed", "Browser Crash", download_folder
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
            options.add_argument("--start-maximized")
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
                if self.year_mode in ASSESSMENT_YEAR_OPTIONS and "-" in self.year_mode and self.year_mode.split("-")[0].isdigit():
                    # Match absolute year like "2024-2025" to "2024-25"
                    start_yr, end_yr = self.year_mode.split("-")
                    target_ay = f"{start_yr}-{end_yr[2:]}"
                    self.current_user_selected_years = [target_ay] if target_ay in available_years else []
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
#  WORKER: AIS & TIS COMBINED (single login, both files)
# ============================================================
class AISTISWorker:
    def __init__(self, app_instance, excel_path, year_mode):
        self.app = app_instance
        self.excel_path = excel_path
        self.year_mode = year_mode
        self.keep_running = True
        self.report_data = []
        self.user_selection_event = threading.Event()
        self.current_user_selected_years = None

    def log(self, message):
        self.app.update_log_safe_aistis(message)

    def set_years_and_resume(self, selected_list):
        self.current_user_selected_years = selected_list
        self.user_selection_event.set()

    def run(self):
        self.log("🚀 INITIALIZING AIS & TIS ENGINE (single login)...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        try:
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)
            if not user_col or not pass_col:
                self.log("❌ ERROR: Headers missing.")
                self.app.process_finished_safe_aistis("Failed: Column Header Error")
                return
            total_users = len(df)
            for index, row in df.iterrows():
                if not self.keep_running: break
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None
                self.app.update_progress_safe_aistis(index / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")
                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "Income Tax Downloaded", "AIS-TIS")
                status, reason, final_path = self.process_single_user(user_id, password, dob, download_root)
                self.report_data.append({
                    "PAN": user_id, "Status": status, "Details": reason,
                    "Folder Saved": os.path.basename(final_path) if final_path else user_id,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                self.log("-" * 40)
            self.generate_report()
            self.app.update_progress_safe_aistis(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe_aistis("All Tasks Completed.")
        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe_aistis("Critical Error Occurred")

    def process_single_user(self, user_id, password, dob, download_root):
        driver = None
        download_folder = tempfile.gettempdir()  # temporary path until name is known
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
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
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            driver.implicitly_wait(10)
            wait = WebDriverWait(driver, 20)

            # 1. LOGIN
            login_success = False
            for login_attempt in range(1, 4):
                if login_success: break
                if login_attempt > 1:
                    try: driver.delete_all_cookies(); driver.refresh()
                    except: pass
                    time.sleep(3)
                try:
                    self.log("   🌐 Opening Portal...")
                    try:
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")
                    except TimeoutException:
                        self.log("   ⚠️ Page load timeout. Retrying..."); continue
                    except Exception as e:
                        self.log(f"   ⚠️ Page load error: {str(e)[:30]}. Retrying..."); continue
                    time.sleep(2)
                    try: driver.switch_to.alert.accept(); time.sleep(1)
                    except: pass
                    pan_entered = False
                    for pan_retry in range(3):
                        try:
                            pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                            pan_field.clear(); pan_field.send_keys(user_id)
                            pan_entered = True; break
                        except:
                            if pan_retry == 2: raise
                            time.sleep(1)
                    if not pan_entered: continue
                    time.sleep(0.5)
                    continue_success = False
                    for cont_retry in range(3):
                        try:
                            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                            driver.execute_script("arguments[0].click();", continue_btn); time.sleep(1)
                            if "ITD-EXEC2002" in driver.page_source or "Something seems to have gone wrong" in driver.page_source:
                                if cont_retry < 2: continue
                                else: raise Exception("ITD-EXEC2002 persists")
                            continue_success = True; break
                        except:
                            if cont_retry == 2: raise
                            time.sleep(1)
                    if not continue_success: continue
                    time.sleep(1.5)
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN", download_folder
                    for pwd_retry in range(3):
                        try:
                            pwd_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                            pwd_field.clear(); pwd_field.send_keys(password); break
                        except:
                            if pwd_retry == 2: raise
                            time.sleep(1)
                    try:
                        driver.execute_script("document.getElementById('passwordCheckBox-input').click();")
                        time.sleep(0.3)
                    except: pass
                    self.log("   ⏳ Waiting for security check (3s)...")
                    time.sleep(3.5)
                    submit_success = False
                    for submit_retry in range(3):
                        try:
                            driver.execute_script("document.querySelector('button.large-button-primary').click();")
                            time.sleep(1)
                            if "ITD-EXEC2002" in driver.page_source or "Something seems to have gone wrong" in driver.page_source:
                                if submit_retry < 2: continue
                                else: raise Exception("ITD-EXEC2002 persists")
                            submit_success = True; break
                        except:
                            if submit_retry == 2: raise
                            time.sleep(1)
                    if not submit_success: continue
                    for _ in range(20):
                        time.sleep(1)
                        try:
                            if driver.find_elements(By.ID, "e-File"):
                                self.log("   ✅ Login Successful!")
                                login_success = True; break
                        except: pass
                        if "Invalid Password" in driver.page_source: return "Failed", "Invalid Password", download_folder
                        try:
                            dual = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual and dual[0].is_displayed():
                                driver.execute_script("arguments[0].click();", dual[0]); time.sleep(2)
                        except: pass
                    if login_success: break
                except Exception as e:
                    self.log(f"   ⚠️ Login Error: {str(e)[:50]}")
                    if login_attempt < 3: time.sleep(2)

            if not login_success: return "Failed", "Login Timeout", download_folder

            name_from_header = get_taxpayer_name(driver, fallback=user_id)
            if name_from_header != user_id:
                self.log(f"   👤 Taxpayer: {name_from_header}")
            else:
                self.log("   ⚠️ Name not found; using PAN as folder name.")
            folder_name = f"{user_id}_{name_from_header}"
            download_folder = create_unique_folder(download_root, folder_name)
            self.log(f"   📁 Folder: {os.path.basename(download_folder)}")
            try:
                driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': download_folder})
            except Exception as cdp_e:
                self.log(f"   ⚠️ CDP redirect failed ({str(cdp_e)[:30]}); folder still created.")

            # 2. NAVIGATE TO AIS PORTAL
            self.log("   🚀 Navigating to AIS Portal...")
            nav_success = False
            for nav_attempt in range(1, 4):
                try:
                    if nav_attempt > 1:
                        self.log(f"   ⚠️ Navigation Retry {nav_attempt}/3...")
                        driver.get("https://eportal.incometax.gov.in/iec/foservices/#/dashboard")
                        time.sleep(2)
                    ais_span = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'mdc-button__label') and contains(text(), 'AIS')]")))
                    driver.execute_script("arguments[0].click();", ais_span)
                    try:
                        proceed_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Proceed')]")))
                        driver.execute_script("arguments[0].click();", proceed_btn)
                    except: pass
                    time.sleep(2)
                    nav_success = True; break
                except Exception as e:
                    self.log(f"   ⚠️ Nav Attempt {nav_attempt} Failed: {str(e)[:40]}")
                    if nav_attempt == 3: return "Failed", "Dashboard AIS Menu Not Found", download_folder
            if not nav_success: return "Failed", "Dashboard AIS Menu Not Found", download_folder

            tab_found = False
            for tab_attempt in range(1, 4):
                try:
                    time.sleep(1)
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        tab_found = True; break
                    else:
                        if tab_attempt < 3:
                            self.log(f"   ⚠️ Waiting for portal tab... Attempt {tab_attempt}/3")
                            time.sleep(2)
                except:
                    if tab_attempt == 3: return "Failed", "AIS Tab did not open", download_folder
            if not tab_found: return "Failed", "AIS Tab did not open", download_folder

            internal_success = False
            for internal_attempt in range(1, 4):
                try:
                    if internal_attempt > 1:
                        self.log(f"   ⚠️ Internal Menu Retry {internal_attempt}/3...")
                        driver.refresh(); time.sleep(2)
                    ais_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]")))
                    driver.execute_script("arguments[0].click();", ais_menu); time.sleep(2)
                    internal_success = True; break
                except Exception as e:
                    self.log(f"   ⚠️ Internal Menu Attempt {internal_attempt} Failed: {str(e)[:40]}")
                    if internal_attempt == 3: return "Failed", "AIS Internal Menu Failed", download_folder
            if not internal_success: return "Failed", "AIS Internal Menu Failed", download_folder

            try:
                self.log("   📥 Fetching Available Years...")
                try:
                    dropdown_toggle = wait.until(EC.presence_of_element_located((By.ID, "dropdownMenuButton")))
                    driver.execute_script("arguments[0].click();", dropdown_toggle); time.sleep(0.5)
                except: pass
                year_buttons = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//button[contains(@class, 'dropdown-item') and contains(text(), 'F.Y.')]")))
                available_years = []
                for btn in year_buttons:
                    txt = btn.get_attribute("textContent").strip()
                    if txt and txt not in available_years: available_years.append(txt)
                try: driver.execute_script("arguments[0].click();", dropdown_toggle)
                except: pass
                if not available_years: return "Failed", "No years found", download_folder

                if self.year_mode in ASSESSMENT_YEAR_OPTIONS and "-" in self.year_mode and self.year_mode.split("-")[0].isdigit():
                    # AIS/TIS portal uses Financial Year. 
                    # Conversion: AY 2024-2025 -> FY 2023-24
                    ay_start = int(self.year_mode.split("-")[0])
                    fy_start = ay_start - 1
                    fy_end = ay_start
                    target_fy = f"F.Y. {fy_start}-{str(fy_end)[2:]}"
                    self.current_user_selected_years = [target_fy] if target_fy in available_years else []

                else:
                    if self.year_mode == "Manual Selection (Popup)":
                        self.log(f"   🛑 PAUSED: Found {len(available_years)} years. Waiting for you...")
                        self.user_selection_event.clear()
                        self.current_user_selected_years = None
                        self.app.trigger_year_selection(available_years, user_id, self.set_years_and_resume)
                        self.user_selection_event.wait()
                    else:
                        self.current_user_selected_years = available_years[:1]

                years_to_download = [y for y in self.current_user_selected_years if y in available_years]
                if not years_to_download: return "Warning", "No valid years selected", download_folder

                self.log(f"   ⬇️ Downloading {len(years_to_download)} Year(s) — AIS + TIS each...")
                ais_count = 0
                tis_count = 0
                AIS_MENU_XPATH = "//span[contains(text(), 'AIS') and contains(@class, 'opacity-6')]"

                for year in years_to_download:
                    self.log(f"   📅 Year: {year}")
                    try:
                        driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': download_folder})
                    except: pass

                    # AIS download
                    ais_year_ok = False
                    for year_attempt in range(1, 4):
                        try:
                            if year_attempt > 1:
                                self.log(f"     -> AIS Retry {year_attempt}/3 for {year}...")
                                driver.refresh(); time.sleep(2)
                                ais_m = wait.until(EC.element_to_be_clickable((By.XPATH, AIS_MENU_XPATH)))
                                driver.execute_script("arguments[0].click();", ais_m); time.sleep(2)
                            else:
                                self.log(f"     -> AIS: {year}...")
                            try:
                                dt = driver.find_element(By.ID, "dropdownMenuButton")
                                driver.execute_script("arguments[0].click();", dt); time.sleep(0.5)
                            except: pass
                            yr_btn = wait.until(EC.presence_of_element_located((By.XPATH, f"//button[contains(@class,'dropdown-item') and contains(text(),'{year}')]")))
                            driver.execute_script("arguments[0].click();", yr_btn); time.sleep(2)
                            dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Download AIS related documents']")))
                            driver.execute_script("arguments[0].click();", dl_icon); time.sleep(1)
                            ais_year_ok = True; break
                        except Exception as e:
                            if year_attempt == 3: self.log(f"        ⚠️ AIS failed for {year} after 3 attempts")

                    if ais_year_ok:
                        for pdf_attempt in range(1, 4):
                            try:
                                if pdf_attempt > 1:
                                    self.log(f"        ⚠️ AIS PDF Retry {pdf_attempt}/3..."); time.sleep(2)
                                clean_temp_files(download_folder, prefixes=("AIS_",))
                                wait.until(EC.element_to_be_clickable((By.XPATH, "//p[contains(text(), 'Annual Information Statement')]/ancestor::app-download//button[contains(@class, 'btn-outline-primary')]")))
                                modal_dl_btn = driver.find_element(By.XPATH, "//p[contains(text(), 'Annual Information Statement')]/ancestor::app-download//button[contains(@class, 'btn-outline-primary')]")
                                modal_click_time = time.time()
                                driver.execute_script("arguments[0].click();", modal_dl_btn)
                                self.log("        Generating AIS Document...")
                                direct_download = False
                                for _ in range(8):
                                    time.sleep(1)
                                    for f in os.listdir(download_folder):
                                        fp = os.path.join(download_folder, f)
                                        if os.path.isfile(fp) and os.path.getmtime(fp) >= modal_click_time - 2:
                                            if f.endswith(".crdownload") or f.endswith(".pdf"):
                                                direct_download = True; break
                                    if direct_download: break
                                if direct_download:
                                    self.log("        ✅ AIS Direct download detected.")
                                    saved_path = wait_and_rename_file(download_folder, year, self.log, prefix="AIS_", start_time=modal_click_time-2, taxpayer_name=name_from_header)
                                    if saved_path:
                                        ais_count += 1; unlock_pdf(saved_path, user_id, dob, self.log); break
                                    else:
                                        self.log("        ❌ AIS file capture failed.")
                                    try:
                                        close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                        driver.execute_script("arguments[0].click();", close_btn)
                                    except: pass
                                else:
                                    self.log("        ℹ️ No AIS direct download. Checking Activity History...")
                                    try:
                                        close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                        driver.execute_script("arguments[0].click();", close_btn); time.sleep(0.5)
                                    except: pass
                                    history_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Go To Activity History']")))
                                    driver.execute_script("arguments[0].click();", history_btn); time.sleep(3)
                                    clean_temp_files(download_folder, prefixes=("AIS_", "TIS_", "20", name_from_header))
                                    hist_click_time = time.time()
                                    final_dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "(//img[@alt='Download'])[1]")))
                                    driver.execute_script("arguments[0].click();", final_dl_icon)
                                    self.log("        ✅ AIS Export Triggered from History.")
                                    saved_path = wait_and_rename_file(download_folder, year, self.log, prefix="AIS_", start_time=hist_click_time-2, taxpayer_name=name_from_header)
                                    if saved_path:
                                        ais_count += 1; unlock_pdf(saved_path, user_id, dob, self.log); break
                            except Exception as e:
                                self.log(f"       ⚠️ AIS Attempt {pdf_attempt} Failed: {str(e)[:30]}")
                                if pdf_attempt == 3: self.log("       ❌ AIS PDF Export failed after 3 attempts.")

                    try:
                        ais_m = wait.until(EC.element_to_be_clickable((By.XPATH, AIS_MENU_XPATH)))
                        driver.execute_script("arguments[0].click();", ais_m); time.sleep(2)
                    except: pass

                    # TIS download
                    tis_year_ok = False
                    for year_attempt in range(1, 4):
                        try:
                            if year_attempt > 1:
                                self.log(f"     -> TIS Retry {year_attempt}/3 for {year}...")
                                driver.refresh(); time.sleep(2)
                                ais_m = wait.until(EC.element_to_be_clickable((By.XPATH, AIS_MENU_XPATH)))
                                driver.execute_script("arguments[0].click();", ais_m); time.sleep(2)
                            else:
                                self.log(f"     -> TIS: {year}...")
                            try:
                                dt = driver.find_element(By.ID, "dropdownMenuButton")
                                driver.execute_script("arguments[0].click();", dt); time.sleep(0.5)
                            except: pass
                            yr_btn = wait.until(EC.presence_of_element_located((By.XPATH, f"//button[contains(@class,'dropdown-item') and contains(text(),'{year}')]")))
                            driver.execute_script("arguments[0].click();", yr_btn); time.sleep(2)
                            tis_dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[contains(@title,'Download TIS related documents') or contains(@alt,'Download TIS related documents')]")))
                            driver.execute_script("arguments[0].click();", tis_dl_icon); time.sleep(1)
                            tis_year_ok = True; break
                        except Exception as e:
                            if year_attempt == 3: self.log(f"        ⚠️ TIS failed for {year} after 3 attempts")

                    if tis_year_ok:
                        for pdf_attempt in range(1, 4):
                            try:
                                if pdf_attempt > 1:
                                    self.log(f"        ⚠️ TIS PDF Retry {pdf_attempt}/3..."); time.sleep(2)
                                clean_temp_files(download_folder, prefixes=("AIS_", "TIS_", "20", name_from_header))
                                wait.until(EC.element_to_be_clickable((By.XPATH, "//p[contains(text(), 'Taxpayer Information Summary')]/ancestor::app-download//button[contains(@class, 'btn-outline-primary')]")))
                                modal_dl_btn = driver.find_element(By.XPATH, "//p[contains(text(), 'Taxpayer Information Summary')]/ancestor::app-download//button[contains(@class, 'btn-outline-primary')]")
                                modal_click_time = time.time()
                                driver.execute_script("arguments[0].click();", modal_dl_btn)
                                self.log("        Generating TIS Document...")
                                direct_download = False
                                for _ in range(8):
                                    time.sleep(1)
                                    for f in os.listdir(download_folder):
                                        fp = os.path.join(download_folder, f)
                                        if os.path.isfile(fp) and os.path.getmtime(fp) >= modal_click_time - 2:
                                            if f.endswith(".crdownload") or f.endswith(".pdf"):
                                                direct_download = True; break
                                    if direct_download: break
                                if direct_download:
                                    self.log("        ✅ TIS Direct download detected.")
                                    saved_path = wait_and_rename_file(download_folder, year, self.log, prefix="TIS_", start_time=modal_click_time-2, taxpayer_name=name_from_header)
                                    if saved_path:
                                        tis_count += 1; unlock_pdf(saved_path, user_id, dob, self.log); break
                                    else:
                                        self.log("        ❌ TIS file capture failed.")
                                    try:
                                        close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                        driver.execute_script("arguments[0].click();", close_btn)
                                    except: pass
                                else:
                                    self.log("        ℹ️ No TIS direct download. Checking Activity History...")
                                    try:
                                        close_btn = driver.find_element(By.XPATH, "//button[contains(translate(text(), 'CLOSE', 'close'), 'close')]")
                                        driver.execute_script("arguments[0].click();", close_btn); time.sleep(0.5)
                                    except: pass
                                    history_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Go To Activity History']")))
                                    driver.execute_script("arguments[0].click();", history_btn); time.sleep(3)
                                    clean_temp_files(download_folder, prefixes=("AIS_", "TIS_", "20", name_from_header))
                                    hist_click_time = time.time()
                                    final_dl_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "(//img[@alt='Download'])[1]")))
                                    driver.execute_script("arguments[0].click();", final_dl_icon)
                                    self.log("        ✅ TIS Export Triggered from History.")
                                    saved_path = wait_and_rename_file(download_folder, year, self.log, prefix="TIS_", start_time=hist_click_time-2, taxpayer_name=name_from_header)
                                    if saved_path:
                                        tis_count += 1; unlock_pdf(saved_path, user_id, dob, self.log); break
                            except Exception as e:
                                self.log(f"       ⚠️ TIS Attempt {pdf_attempt} Failed: {str(e)[:30]}")
                                if pdf_attempt == 3: self.log("       ❌ TIS PDF Export failed after 3 attempts.")

                    try:
                        ais_m = wait.until(EC.element_to_be_clickable((By.XPATH, AIS_MENU_XPATH)))
                        driver.execute_script("arguments[0].click();", ais_m); time.sleep(2)
                    except: pass

                return "Success", f"AIS: {ais_count} files, TIS: {tis_count} files", download_folder

            except Exception as e:
                return "Failed", f"Portal Error: {str(e)[:20]}", download_folder

        except Exception as e:
            return "Failed", "Browser Crash", download_folder
        finally:
            if driver: driver.quit()

# ============================================================
#  MAIN APP GUI
# ============================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Suite Pro")
        self.geometry("900x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.worker = None

        # --- Header ---
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 5))
        self.title_label = ctk.CTkLabel(self.header_frame, text="INCOME TAX AUTOMATION SUITE", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(side="left")
        
        # --- Main Tab View ---
        self.tabview = ctk.CTkTabview(self, width=860)
        self.tabview.add("26AS")
        self.tabview.add("AIS & TIS")
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=20, pady=5)

        self.tab_26as = self.tabview.tab("26AS")
        self.tab_aistis = self.tabview.tab("AIS & TIS")

        self.manual_credentials = []
        self._build_26as_ui()
        self._build_aistis_ui()

    # --- UI BUILDERS ---
    def _build_26as_ui(self):
        self.tab_26as.grid_columnconfigure(0, weight=1)
        self.tab_26as.grid_rowconfigure(0, weight=1)

        # SCROLLABLE CONTAINER FOR 26AS
        self.scroll_26as = ctk.CTkScrollableFrame(self.tab_26as, fg_color="transparent")
        self.scroll_26as.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll_26as.grid_columnconfigure(0, weight=1)
        self.tab_26as.grid_rowconfigure(0, weight=1)
        self.tab_26as.grid_rowconfigure(1, weight=0)

        self.config_26as = ctk.CTkFrame(self.scroll_26as)
        self.config_26as.grid(row=0, column=0, sticky="ew", padx=10, pady=(2, 5))

        ctk.CTkLabel(self.config_26as, text="1. CREDENTIALS SOURCE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        f_frame = ctk.CTkFrame(self.config_26as, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file_26as = ctk.CTkEntry(f_frame, placeholder_text="Add PAN, Password, DOB manually...")
        self.entry_file_26as.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_actions = ctk.CTkFrame(f_frame, fg_color="transparent")
        btn_actions.pack(side="right")
        # Add ID first
        ctk.CTkButton(btn_actions, text="➕ Add ID Password", command=lambda: self.add_id_password("26as"), width=150, fg_color="#059669", hover_color="#047857", font=("Segoe UI", 12, "bold")).pack(side="left")
        # View and Delete next
        self.btn_view_26as = ctk.CTkButton(btn_actions, text="👁 View ID", command=lambda: self.view_saved_user("26as"), width=95, fg_color="#475569", hover_color="#334155", font=("Segoe UI", 11, "bold"))
        self.btn_view_26as.pack(side="left", padx=(5, 0))
        self.btn_delete_26as = ctk.CTkButton(btn_actions, text="🗑 Delete ID", command=lambda: self.delete_saved_user("26as"), width=105, fg_color="#7C3AED", hover_color="#6D28D9", font=("Segoe UI", 11, "bold"))
        self.btn_delete_26as.pack(side="left", padx=(5, 0))
        # Demo last
        ctk.CTkButton(btn_actions, text="▶ Demo", command=self.open_demo_link, width=80, fg_color="#DC2626", hover_color="#B91C1C", font=("Segoe UI", 12, "bold")).pack(side="left", padx=(5, 0))
        self.btn_view_26as.configure(state="disabled")
        self.btn_delete_26as.configure(state="disabled")

        pref_frame = ctk.CTkFrame(self.config_26as, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(pref_frame, text="Assessment Year:", text_color="gray").pack(side="left", padx=(0, 10))
        self.combo_years_26as = ctk.CTkComboBox(pref_frame, values=ASSESSMENT_YEAR_OPTIONS, width=250, state="readonly")
        self.combo_years_26as.set(ASSESSMENT_YEAR_OPTIONS[0])
        self.combo_years_26as.pack(side="left")

        self.log_frame_26as = ctk.CTkFrame(self.scroll_26as)
        self.log_frame_26as.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 5))
        self.log_frame_26as.grid_rowconfigure(1, weight=1)
        self.log_frame_26as.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.log_frame_26as, text="3. LIVE LOG", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(2, 2))
        self.log_box_26as = ctk.CTkTextbox(self.log_frame_26as, font=("Consolas", 12), activate_scrollbars=True, height=100)
        self.log_box_26as.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box_26as.configure(state="disabled")
        
        self.progress_26as = ctk.CTkProgressBar(self.log_frame_26as, mode="determinate")
        self.progress_26as.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_26as.set(0)

        btn_footer_26as = ctk.CTkFrame(self.tab_26as, fg_color="transparent")
        btn_footer_26as.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))
        self.btn_start_26as = ctk.CTkButton(btn_footer_26as, text="START 26AS DOWNLOAD", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=lambda: self.start_process("26as"))
        self.btn_start_26as.pack(side="left", expand=True, fill="x")
        self.btn_stop_26as = ctk.CTkButton(btn_footer_26as, text="⏹ STOP", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#DC2626", hover_color="#B91C1C", command=lambda: self.stop_process("26as"), width=150)
        self.btn_stop_26as.pack(side="left", padx=(10, 0))
        self.btn_stop_26as.pack_forget()
        self.btn_open_folder_26as = ctk.CTkButton(btn_footer_26as, text="📂 OPEN FOLDER", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#2563EB", hover_color="#1D4ED8", command=lambda: self.open_download_folder("26as"), width=180)
        self.btn_open_folder_26as.pack(side="left", padx=(10, 0))
        self.btn_open_folder_26as.pack_forget()

    def _build_aistis_ui(self):
        self.tab_aistis.grid_columnconfigure(0, weight=1)
        self.tab_aistis.grid_rowconfigure(0, weight=1)

        # SCROLLABLE CONTAINER FOR AIS/TIS
        self.scroll_aistis = ctk.CTkScrollableFrame(self.tab_aistis, fg_color="transparent")
        self.scroll_aistis.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll_aistis.grid_columnconfigure(0, weight=1)
        self.tab_aistis.grid_rowconfigure(0, weight=1)
        self.tab_aistis.grid_rowconfigure(1, weight=0)

        self.config_aistis = ctk.CTkFrame(self.scroll_aistis)
        self.config_aistis.grid(row=0, column=0, sticky="ew", padx=10, pady=(2, 5))

        ctk.CTkLabel(self.config_aistis, text="1. CREDENTIALS SOURCE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        f_frame = ctk.CTkFrame(self.config_aistis, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file_aistis = ctk.CTkEntry(f_frame, placeholder_text="Add PAN, Password, DOB manually...")
        self.entry_file_aistis.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_actions = ctk.CTkFrame(f_frame, fg_color="transparent")
        btn_actions.pack(side="right")
        ctk.CTkButton(btn_actions, text="➕ Add ID Password", command=lambda: self.add_id_password("aistis"), width=150, fg_color="#059669", hover_color="#047857", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.btn_view_aistis = ctk.CTkButton(btn_actions, text="👁 View ID", command=lambda: self.view_saved_user("aistis"), width=95, fg_color="#475569", hover_color="#334155", font=("Segoe UI", 11, "bold"))
        self.btn_view_aistis.pack(side="left", padx=(5, 0))
        self.btn_delete_aistis = ctk.CTkButton(btn_actions, text="🗑 Delete ID", command=lambda: self.delete_saved_user("aistis"), width=105, fg_color="#7C3AED", hover_color="#6D28D9", font=("Segoe UI", 11, "bold"))
        self.btn_delete_aistis.pack(side="left", padx=(5, 0))
        ctk.CTkButton(btn_actions, text="▶ Demo", command=self.open_demo_link, width=80, fg_color="#DC2626", hover_color="#B91C1C", font=("Segoe UI", 12, "bold")).pack(side="left", padx=(5, 0))
        self.btn_view_aistis.configure(state="disabled")
        self.btn_delete_aistis.configure(state="disabled")

        pref_frame = ctk.CTkFrame(self.config_aistis, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(pref_frame, text="Assessment Year:", text_color="gray").pack(side="left", padx=(0, 10))
        self.combo_years_aistis = ctk.CTkComboBox(pref_frame, values=ASSESSMENT_YEAR_OPTIONS, width=250, state="readonly")
        self.combo_years_aistis.set(ASSESSMENT_YEAR_OPTIONS[0])
        self.combo_years_aistis.pack(side="left")

        self.log_frame_aistis = ctk.CTkFrame(self.scroll_aistis)
        self.log_frame_aistis.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 5))
        self.log_frame_aistis.grid_rowconfigure(1, weight=1)
        self.log_frame_aistis.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.log_frame_aistis, text="3. LIVE LOG", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(2, 2))
        self.log_box_aistis = ctk.CTkTextbox(self.log_frame_aistis, font=("Consolas", 12), activate_scrollbars=True, height=100)
        self.log_box_aistis.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box_aistis.configure(state="disabled")

        self.progress_aistis = ctk.CTkProgressBar(self.log_frame_aistis, mode="determinate")
        self.progress_aistis.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_aistis.set(0)

        btn_footer_aistis = ctk.CTkFrame(self.tab_aistis, fg_color="transparent")
        btn_footer_aistis.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))
        self.btn_start_aistis = ctk.CTkButton(btn_footer_aistis, text="START AIS & TIS DOWNLOAD", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=lambda: self.start_process("aistis"))
        self.btn_start_aistis.pack(side="left", expand=True, fill="x")
        self.btn_stop_aistis = ctk.CTkButton(btn_footer_aistis, text="⏹ STOP", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#DC2626", hover_color="#B91C1C", command=lambda: self.stop_process("aistis"), width=150)
        self.btn_stop_aistis.pack(side="left", padx=(10, 0))
        self.btn_stop_aistis.pack_forget()
        self.btn_open_folder_aistis = ctk.CTkButton(btn_footer_aistis, text="📂 OPEN FOLDER", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#2563EB", hover_color="#1D4ED8", command=lambda: self.open_download_folder("aistis"), width=180)
        self.btn_open_folder_aistis.pack(side="left", padx=(10, 0))
        self.btn_open_folder_aistis.pack_forget()

    # --- GUI Handlers ---
    def trigger_year_selection(self, years_list, user_id, callback):
        self.after(0, lambda: self._show_popup(years_list, user_id, callback))

    def _show_popup(self, years_list, user_id, callback):
        YearSelectionPopup(self, years_list, user_id, callback)
    def download_sample(self):
        import shutil
        import os
        from tkinter import messagebox
        sample_path = os.path.join(os.path.dirname(__file__), "Income Tax Sample File.xlsx")
        if not os.path.exists(sample_path):
            messagebox.showerror("Download Error", f"Sample file not found: {sample_path}")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Income Tax Sample File.xlsx", filetypes=[("Excel", "*.xlsx")])
        if save_path:
            try:
                shutil.copy2(sample_path, save_path)
                messagebox.showinfo("Success", f"Sample downloaded to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download: {e}")

    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://youtu.be/byMvFynIJuo")

    def browse_file(self, mode):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.manual_credentials = []
            self._refresh_manual_controls()
            if mode == "26as":
                self.excel_file_path_26as = filename
                self.entry_file_26as.delete(0, "end")
                self.entry_file_26as.insert(0, filename)
                self.log_to_gui_26as(f"File Loaded: {os.path.basename(filename)}")

    def _get_saved_user_id(self):
        if not self.manual_credentials:
            return ""
        return str(self.manual_credentials[0].get("PAN", "")).strip()

    def _refresh_manual_controls(self):
        has_manual = bool(self.manual_credentials)
        for btn_attr in ["btn_view_26as", "btn_delete_26as", "btn_view_aistis", "btn_delete_aistis"]:
            btn = getattr(self, btn_attr, None)
            if btn is not None:
                btn.configure(state="normal" if has_manual else "disabled")

        if has_manual:
            user_id = self._get_saved_user_id()
            manual_text = f"Selected ID: {user_id}"
            for entry_attr in ["entry_file_26as", "entry_file_aistis"]:
                entry = getattr(self, entry_attr, None)
                if entry is not None:
                    entry.delete(0, "end")
                    entry.insert(0, manual_text)

    def view_saved_user(self, mode=None):
        user_id = self._get_saved_user_id()
        if not user_id:
            messagebox.showinfo("Info", "No saved ID found.")
            return
        messagebox.showinfo("Saved User ID", f"Current ID: {user_id}")

    def delete_saved_user(self, mode=None):
        user_id = self._get_saved_user_id()
        if not user_id:
            messagebox.showinfo("Info", "No saved ID found.")
            return
        if not messagebox.askyesno("Delete ID", f"Delete saved ID {user_id}?"):
            return
        self.manual_credentials = []
        for entry_attr in ["entry_file_26as", "entry_file_aistis"]:
            entry = getattr(self, entry_attr, None)
            if entry is not None:
                entry.delete(0, "end")
        self._refresh_manual_controls()
        messagebox.showinfo("Deleted", "Saved ID deleted successfully.")

    def add_id_password(self, mode):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add ID Password")
        dialog.geometry("430x300")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        card = ctk.CTkFrame(dialog, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(card, text="PAN / User ID").pack(anchor="w")
        ent_user = ctk.CTkEntry(card, placeholder_text="Enter PAN / User ID")
        ent_user.pack(fill="x", pady=(4, 10))

        ctk.CTkLabel(card, text="Password").pack(anchor="w")
        pass_frm = ctk.CTkFrame(card, fg_color="transparent")
        pass_frm.pack(fill="x", pady=(4, 10))
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

        ctk.CTkLabel(card, text="DOB (optional)").pack(anchor="w")
        ent_dob = ctk.CTkEntry(card, placeholder_text="DD/MM/YYYY")
        ent_dob.pack(fill="x", pady=(4, 14))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x")

        def _save():
            user_id = (ent_user.get() or "").strip()
            password = (ent_pass.get() or "").strip()
            dob = (ent_dob.get() or "").strip()
            if not user_id or not password:
                messagebox.showerror("Missing Data", "Please enter PAN/User ID and Password", parent=dialog)
                return

            existing_user = self._get_saved_user_id()
            if existing_user and not messagebox.askyesno(
                "Overwrite ID",
                "Your previous ID will be overwritten with this.",
                parent=dialog
            ):
                return

            self.manual_credentials = [{"PAN": user_id, "Password": password, "DOB": dob}]
            self.excel_file_path_26as = ""
            self.excel_file_path_aistis = ""

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
            user_id = str(item.get("PAN", "")).strip()
            password = str(item.get("Password", "")).strip()
            dob = str(item.get("DOB", "")).strip()
            if user_id and password:
                rows.append({"PAN": user_id, "Password": password, "DOB": dob})

        if not rows:
            return ""

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="it_26as_manual_") as tmp:
            temp_excel = tmp.name
        pd.DataFrame(rows, columns=["PAN", "Password", "DOB"]).to_excel(temp_excel, index=False)
        return temp_excel

    def _resolve_excel_path(self, mode):
        if mode == "26as":
            path = self.excel_file_path_26as
        elif mode == "aistis":
            path = self.excel_file_path_aistis
        else:
            path = ""

        if path:
            return path

        if self.manual_credentials:
            return self._create_manual_excel()

        return ""

    def _get_selected_year_mode(self, mode):
        combo_map = {
            "26as": "combo_years_26as",
            "aistis": "combo_years_aistis",
        }
        combo_attr = combo_map.get(mode)
        if not combo_attr:
            return ASSESSMENT_YEAR_OPTIONS[0]

        combo = getattr(self, combo_attr, None)
        try:
            selected = (combo.get() or "").strip() if combo is not None else ""
        except Exception:
            selected = ""

        return selected if selected in ASSESSMENT_YEAR_OPTIONS else ASSESSMENT_YEAR_OPTIONS[0]

    def start_process(self, mode):
        if mode == "26as":
            excel_path = self._resolve_excel_path("26as")
            if not excel_path: return messagebox.showwarning("Error", "Select file or add ID/Password first")
            year_mode = self._get_selected_year_mode("26as")
            self.btn_start_26as.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            self.btn_stop_26as.pack(side="left", padx=(10, 0))
            self.btn_open_folder_26as.pack_forget()
            self.progress_26as.set(0)
            self.worker = Tax26ASWorker(self, excel_path, year_mode)
            threading.Thread(target=self.worker.run, daemon=True).start()
        elif mode == "aistis":
            excel_path = self._resolve_excel_path("aistis")
            if not excel_path: return messagebox.showwarning("Error", "Select file or add ID/Password first")
            year_mode = self._get_selected_year_mode("aistis")
            self.btn_start_aistis.configure(state="disabled", text="PROCESSING...", fg_color="gray")
            self.btn_stop_aistis.pack(side="left", padx=(10, 0))
            self.btn_open_folder_aistis.pack_forget()
            self.progress_aistis.set(0)
            self.worker = AISTISWorker(self, excel_path, year_mode)
            threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self, mode=None):
        if self.worker:
            self.worker.keep_running = False
        if mode == "26as" or mode is None:
            try: self.btn_stop_26as.configure(state="disabled", text="Stopping...")
            except: pass
        if mode == "aistis" or mode is None:
            try: self.btn_stop_aistis.configure(state="disabled", text="Stopping...")
            except: pass

    # --- 26AS SAFE UPDATERS ---
    def log_to_gui_26as(self, msg):
        self.log_box_26as.configure(state="normal")
        self.log_box_26as.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box_26as.see("end")
        self.log_box_26as.configure(state="disabled")

    def update_log_safe_26as(self, msg): self.after(0, lambda: self.log_to_gui_26as(msg))
    def update_progress_safe_26as(self, val): self.after(0, lambda: self.progress_26as.set(val))
    def process_finished_safe_26as(self, msg):
        def _finish():
            self.log_to_gui_26as(f"\nSTATUS: {msg}")
            self.btn_start_26as.configure(state="normal", text="START 26AS DOWNLOAD", fg_color="#2563EB")
            self.btn_stop_26as.configure(state="normal", text="⏹ STOP")
            self.btn_stop_26as.pack_forget()
            self.btn_open_folder_26as.pack(side="left", padx=(10, 0))
            messagebox.showinfo("Done", msg)
        self.after(0, _finish)

    # --- AIS & TIS SAFE UPDATERS ---
    def log_to_gui_aistis(self, msg):
        self.log_box_aistis.configure(state="normal")
        self.log_box_aistis.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box_aistis.see("end")
        self.log_box_aistis.configure(state="disabled")

    def update_log_safe_aistis(self, msg): self.after(0, lambda: self.log_to_gui_aistis(msg))
    def update_progress_safe_aistis(self, val): self.after(0, lambda: self.progress_aistis.set(val))
    def process_finished_safe_aistis(self, msg):
        def _finish():
            self.log_to_gui_aistis(f"\nSTATUS: {msg}")
            self.btn_start_aistis.configure(state="normal", text="START AIS & TIS DOWNLOAD", fg_color="#2563EB")
            self.btn_stop_aistis.configure(state="normal", text="⏹ STOP")
            self.btn_stop_aistis.pack_forget()
            self.btn_open_folder_aistis.pack(side="left", padx=(10, 0))
            messagebox.showinfo("Done", msg)
        self.after(0, _finish)

    def open_download_folder(self, mode):
        folder_map = {
            "26as":   os.path.join("Income Tax Downloaded", "26 AS"),
            "aistis": os.path.join("Income Tax Downloaded", "AIS-TIS"),
        }
        try:
            folder = folder_map.get(mode, os.path.join("Income Tax Downloaded", "26 AS"))
            target = os.path.join(os.getcwd(), folder)
            if not os.path.exists(target):
                target = os.path.join(os.getcwd(), "Income Tax Downloaded")
            os.startfile(target)
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Error", f"Failed to open folder: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()