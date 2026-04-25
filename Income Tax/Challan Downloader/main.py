import re
import glob
import shutil
import threading
import time
import os
import tempfile
import pandas as pd
import customtkinter as ctk
from datetime import datetime
from tkinter import filedialog, messagebox

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ============================================================
#  BASE HELPER FUNCTIONS
# ============================================================
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

# ============================================================
#  WORKER CLASS: CHALLAN DOWNLOADER
# ============================================================
class ChallanWorker:
    def __init__(self, app_instance, excel_path, year_mode):
        self.app = app_instance
        self.excel_path = excel_path
        self.year_mode = year_mode
        self.keep_running = True
        self.report_data = []

    def log(self, message):
        self.app.update_log_safe(message)

    def run(self):
        self.log("🚀 INITIALIZING CHALLAN DOWNLOAD ENGINE...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")
        
        try:
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)
            
            if not user_col or not pass_col:
                self.log(f"❌ ERROR: Headers missing. Need 'PAN' and 'Password'.")
                self.app.process_finished_safe("Failed: Column Header Error")
                return

            self.log(f"✅ Mapped Columns -> ID: '{user_col}', Pass: '{pass_col}'")
            total_users = len(df)
            
            for index, row in df.iterrows():
                if not self.keep_running: 
                    self.log("🛑 Process Stopped by User.")
                    break
                
                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None
                
                self.app.update_progress_safe((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "Income Tax Downloaded", "Challan Downloader")

                status, reason, final_path = self.process_single_user(user_id, password, dob, download_root)

                entry = {
                    "PAN": user_id, "Status": status, "Details": reason,
                    "Folder Saved": os.path.basename(final_path) if final_path else user_id,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                self.report_data.append(entry)
                self._save_user_report(entry, final_path)
                self.log("-" * 40)

            self.generate_report()
            self.app.update_progress_safe(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe("Critical Error Occurred")



    def _save_user_report(self, entry, folder_path):
        try:
            if not folder_path or not os.path.exists(folder_path):
                return
            pan = entry.get("PAN", "unknown")
            report_path = os.path.join(folder_path, f"Report_{pan}.xlsx")
            pd.DataFrame([entry]).to_excel(report_path, index=False)
            self.log(f"   📄 Report saved: Report_{pan}.xlsx")
        except Exception as e:
            self.log(f"   ⚠️ Failed to save user report: {e}")

    def generate_report(self):
        try:
            if not self.report_data: return
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Challan_Report_{timestamp}.xlsx"
            report_dir = os.path.join(os.getcwd(), "Income Tax Downloaded", "Challan Downloader", "reports")
            os.makedirs(report_dir, exist_ok=True)
            report_path = os.path.join(report_dir, filename)
            df_report = pd.DataFrame(self.report_data)
            df_report.to_excel(report_path, index=False)
            self.log(f"📄 Summary report saved: {report_path}")
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
                "safebrowsing.enabled": True,
            }
            options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            driver.implicitly_wait(10)
            wait = WebDriverWait(driver, 20)

            # --- LOGIN ---
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

                    pan_entered = False
                    for pan_retry in range(3):
                        try:
                            pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                            pan_field.clear()
                            pan_field.send_keys(user_id)
                            pan_entered = True
                            break
                        except:
                            if pan_retry == 2: raise
                            time.sleep(1)

                    if not pan_entered: continue
                    time.sleep(0.5)

                    for cont_retry in range(3):
                        try:
                            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                            driver.execute_script("arguments[0].click();", continue_btn)
                            break
                        except:
                            if cont_retry == 2: raise
                            time.sleep(1)

                    time.sleep(1.5)
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN", download_folder

                    for pwd_retry in range(3):
                        try:
                            pwd_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                            pwd_field.clear()
                            pwd_field.send_keys(password)
                            break
                        except:
                            if pwd_retry == 2: raise
                            time.sleep(1)

                    try:
                        driver.execute_script("document.getElementById('passwordCheckBox-input').click();")
                        time.sleep(0.3)
                    except: pass

                    self.log("   ⏳ Waiting for security check (3s)...")
                    time.sleep(3.5)

                    for submit_retry in range(3):
                        try:
                            driver.execute_script("document.querySelector('button.large-button-primary').click();")
                            break
                        except:
                            if submit_retry == 2: raise
                            time.sleep(1)

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

            # Extract taxpayer name and create NAME_PAN folder
            name_from_header = get_taxpayer_name(driver, fallback=user_id)
            if name_from_header != user_id:
                self.log(f"   👤 Taxpayer Name: {name_from_header}")
            else:
                self.log("   ⚠️ Name not found in header; using PAN as folder name.")

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


            # ==========================================================
            # STEP A: Click "e-File" from the top navigation menu
            # ==========================================================
            self.log("   🔹 Step A: Clicking 'e-File' menu...")
            efile_clicked = False
            for attempt in range(3):
                try:
                    efile_btn = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//span[contains(@class,'mdc-button__label') and normalize-space(text())='e-File']"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", efile_btn)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", efile_btn)
                    self.log("   ✅ 'e-File' clicked.")
                    efile_clicked = True
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Failed to click 'e-File': {str(e)[:60]}")
                        return "Failed", "Could not click e-File"
                    self.log(f"   ⚠️ Retry {attempt+1}/3 for 'e-File'...")
                    time.sleep(1.5)

            # ==========================================================
            # STEP B: Click "e-Pay Tax" from the dropdown
            # ==========================================================
            self.log("   🔹 Step B: Clicking 'e-Pay Tax' menu item...")
            epay_clicked = False
            for attempt in range(3):
                try:
                    epay_item = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//span[contains(@class,'mat-mdc-menu-item-text')]//span[normalize-space(text())='e-Pay Tax']"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", epay_item)
                    time.sleep(0.4)
                    driver.execute_script("arguments[0].click();", epay_item)
                    self.log("   ✅ 'e-Pay Tax' clicked.")
                    epay_clicked = True
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Failed to click 'e-Pay Tax': {str(e)[:60]}")
                        return "Failed", "Could not click e-Pay Tax"
                    self.log(f"   ⚠️ Retry {attempt+1}/3 for 'e-Pay Tax'...")
                    time.sleep(1.5)

            # ==========================================================
            # STEP B-2: Select Applicable Income Tax Act (New portal logic)
            # ==========================================================
            try:
                time.sleep(2)
                if driver.find_elements(By.XPATH, "//*[contains(text(), 'Select Applicable Income Tax Act')]"):
                    self.log("   🔹 Step B-2: Income Tax Act selection screen detected.")
                    try:
                        # Find both radio buttons (2025 logic for current year, 1961 for history)
                        # Currently we select "Income-tax Act, 1961" which includes historical and previous FY
                        # as that aligns with standard "Challan Download" behavior for all History.
                        # Using dynamic text match for robust selection.
                        act_1961 = driver.find_element(By.XPATH, "//div[contains(text(), 'Income-tax Act, 1961')]/ancestor::label")
                        driver.execute_script("arguments[0].click();", act_1961)
                        self.log("   ✅ Selected 'Income-tax Act, 1961' (Historical & Previous FY).")
                        time.sleep(0.5)
                        
                        continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary.iconsAfter.nextIcon")))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", continue_btn)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", continue_btn)
                        self.log("   ✅ Proceeded past Act selection (Clicked Continue).")
                        time.sleep(2)
                    except Exception as e:
                        self.log(f"   ⚠️ Failed to select Act, might be pre-selected. {str(e)[:40]}")
            except Exception as ignore:
                pass

            # ==========================================================
            # STEP C: Click "Payment History" tab
            # ==========================================================
            self.log("   🔹 Step C: Clicking 'Payment History' tab...")
            time.sleep(2)  # Allow the e-Pay Tax page to load
            for attempt in range(3):
                try:
                    pay_hist_tab = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//span[contains(@class,'mdc-tab__text-label') and contains(normalize-space(text()),'Payment History')]"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", pay_hist_tab)
                    time.sleep(0.4)
                    driver.execute_script("arguments[0].click();", pay_hist_tab)
                    self.log("   ✅ 'Payment History' tab clicked.")
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Failed to click 'Payment History': {str(e)[:60]}")
                        return "Failed", "Could not click Payment History tab"
                    self.log(f"   ⚠️ Retry {attempt+1}/3 for 'Payment History'...")
                    time.sleep(1.5)

            # Wait for Payment History AG Grid table to load
            time.sleep(3)
            self.log("   ✅ Payment History page loaded.")

            # ==========================================================
            # STEP D: Fetch all Assessment Years from the AG Grid table
            # ==========================================================
            self.log("   🔹 Step D: Reading Assessment Years from table...")
            
            # Fast empty table check
            try:
                empty_grid_check = driver.find_elements(By.CSS_SELECTOR, "div.ag-center-cols-container")
                if empty_grid_check and empty_grid_check[0].get_attribute("style") and "height: 1px" in empty_grid_check[0].get_attribute("style"):
                    self.log("   ⚠️ Table is completely empty. No challan records found.")
                    return "Success", "No Challan Records Found"
            except:
                pass

            all_years = []
            for fetch_attempt in range(3):
                try:
                    # Wait for at least one assessmentYear cell to appear
                    wait.until(EC.presence_of_element_located((
                        By.CSS_SELECTOR, "[col-id='assessmentYear'].ag-cell-value"
                    )))
                    time.sleep(1)  # Let all rows render
                    year_cells = driver.find_elements(
                        By.CSS_SELECTOR, "[col-id='assessmentYear'].ag-cell-value"
                    )
                    raw_years = [c.text.strip() for c in year_cells if c.text.strip()]
                    # Deduplicate while preserving order
                    seen = set()
                    for y in raw_years:
                        if y and y not in seen:
                            seen.add(y)
                            all_years.append(y)
                    if all_years:
                        break
                    time.sleep(1.5)
                except Exception as e:
                    if fetch_attempt == 2:
                        self.log(f"   ❌ Could not read year table: {str(e)[:60]}")
                        return "Failed", "Could not read Assessment Year table"
                    self.log(f"   ⚠️ Retry {fetch_attempt+1}/3 reading year table...")
                    time.sleep(2)

            if not all_years:
                self.log("   ⚠️ No challan records found in Payment History.")
                return "Success", "No Challan Records Found"

            # Sort years descending (e.g. ["2026-27","2025-26","2024-25"])
            def year_sort_key(y):
                try: return int(y.split("-")[0])
                except: return 0
            all_years.sort(key=year_sort_key, reverse=True)
            self.log(f"   📋 All Years Found: {', '.join(all_years)}")

            # ==========================================================
            # STEP E: Filter years based on selected Download Filter
            # Supports selecting a specific Financial Year (e.g. '2026-2027')
            # ==========================================================
            selected_years = []
            mode = (self.year_mode or "").strip()

            # If a specific FY like '2026-2027' or '2026-27' was chosen,
            # match available entries by the starting year (left side of dash).
            m = re.match(r'^(\d{4})\s*-\s*(\d{2,4})$', mode)
            if m:
                try:
                    start_year = int(m.group(1))
                    for y in all_years:
                        try:
                            if int(str(y).split("-")[0]) == start_year:
                                selected_years.append(y)
                        except:
                            continue
                except:
                    selected_years = []
            elif mode == "Last 2 Years":
                selected_years = all_years[:2]
            elif mode == "All History":
                selected_years = all_years
            else:
                # Default to most recent year if nothing matches
                selected_years = all_years[:1]

            if not selected_years:
                if re.match(r'^(\d{4})\s*-\s*(\d{2,4})$', mode):
                    self.log(f"   ⚠️ No data found for Assessment Year: {mode}")
                    self.log(f"   📋 Available Years: {', '.join(all_years)}")
                    return "Success", f"No data found for Assessment Year {mode}", download_folder
                selected_years = all_years[:1]

            self.log(f"   🎯 Selected Years to Download ({mode}): {', '.join(selected_years)}")

            # ==========================================================
            # STEP F: For each selected year — click ⋮ (more_vert) → Download
            # Angular Material menus render in a global overlay (outside the row),
            # so we must locate the row first, click its icon button, then find
            # the Download item in the CDK overlay container.
            # ==========================================================
            downloaded_years = []
            failed_years = []

            for year in selected_years:
                self.log(f"   🔹 Processing Year: {year}")
                
                # Create subfolder for the year (financial year wise)
                safe_year = year.replace('/', '-').replace(' ', '_').strip()
                year_folder = os.path.join(download_folder, safe_year)
                os.makedirs(year_folder, exist_ok=True)
                
                # Update download behavior for this specific year folder
                try:
                    driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                        'behavior': 'allow',
                        'downloadPath': year_folder
                    })
                except: pass

                for dl_attempt in range(3):
                    try:
                        # Dismiss any leftover open menu
                        try:
                            driver.execute_script("document.body.click();")
                            time.sleep(0.5)
                        except: pass

                        # --- F1: Find the AG Grid row whose assessmentYear cell = year ---
                        year_cell_xpath = (
                            f"//div[@col-id='assessmentYear' and "
                            f"contains(@class,'ag-cell') and "
                            f"normalize-space(text())='{year}']"
                        )
                        year_cell = wait.until(EC.presence_of_element_located((By.XPATH, year_cell_xpath)))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", year_cell)
                        time.sleep(0.5)

                        # --- F2: From that cell climb to role="row", find the more_vert button ---
                        # The action button contains a <mat-icon> with text "more_vert"
                        row_el = year_cell.find_element(By.XPATH, "./ancestor::div[@role='row'][1]")
                        more_btn = row_el.find_element(
                            By.XPATH,
                            ".//button[contains(@class,'mat-mdc-icon-button')]"
                            "[.//mat-icon[normalize-space(text())='more_vert']]"
                        )
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", more_btn)
                        time.sleep(0.3)
                        driver.execute_script("arguments[0].click();", more_btn)
                        self.log(f"      ✅ ⋮ menu opened for {year}.")

                        # --- F3: Wait for Angular Material overlay, then click Download ---
                        # Mat menus render inside .cdk-overlay-container at the document level
                        download_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((
                                By.XPATH,
                                "//div[contains(@class,'cdk-overlay-container')]"
                                "//button[contains(@class,'mat-mdc-menu-item')]"
                                "[.//span[normalize-space(text())='Download']]"
                            ))
                        )
                        driver.execute_script("arguments[0].click();", download_btn)
                        self.log(f"      ✅ 'Download' clicked for {year}.")
                        time.sleep(4)  # Wait for file to appear

                        # Rename the downloaded PDF to NAME-PAN-YEAR.pdf
                        pdf_files = sorted(glob.glob(os.path.join(year_folder, '*.pdf')), key=os.path.getmtime, reverse=True)
                        if pdf_files:
                            latest_pdf = pdf_files[0]
                            new_name = f"{name_from_header}-{user_id}-{year}.pdf"
                            new_path = os.path.join(year_folder, new_name)
                            try:
                                shutil.move(latest_pdf, new_path)
                                self.log(f"      📄 Renamed to: {new_name}")
                            except Exception as e:
                                self.log(f"      ⚠️ Could not rename PDF: {str(e)[:60]}")

                        downloaded_years.append(year)
                        break

                    except Exception as e:
                        # Dismiss menu before retry
                        try: driver.execute_script("document.body.click();")
                        except: pass
                        time.sleep(1)

                        if dl_attempt == 2:
                            self.log(f"      ❌ Download failed for {year}: {str(e)[:80]}")
                            failed_years.append(year)
                        else:
                            self.log(f"      ⚠️ Retry {dl_attempt+1}/3 for year {year}...")
                            time.sleep(2)

            summary = f"Downloaded: {', '.join(downloaded_years) or 'None'}"
            if failed_years:
                summary += f" | Failed: {', '.join(failed_years)}"
            self.log(f"   📦 {summary}")
            return ("Success" if downloaded_years else "Failed"), summary, download_folder

        except Exception as e: 
            return "Failed", f"Browser Error: {str(e)[:20]}", download_folder
        finally:
            if driver: driver.quit()


# ============================================================
#  WORKER CLASS: DEMAND CHECKER
# ============================================================
class DemandCheckerWorker:
    def __init__(self, app_instance, excel_path):
        self.app = app_instance
        self.excel_path = excel_path
        self.keep_running = True
        self.report_data = []

    def log(self, message):
        self.app.update_log_safe_demand(message)

    def run(self):
        self.log("🚀 INITIALIZING DEMAND CHECKER ENGINE...")
        self.log(f"📂 Reading Credentials: {os.path.basename(self.excel_path)}")

        try:
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)

            if not user_col or not pass_col:
                self.log("❌ ERROR: Headers missing. Need 'PAN' and 'Password'.")
                self.app.process_finished_safe_demand("Failed: Column Header Error")
                return

            self.log(f"✅ Mapped Columns -> ID: '{user_col}', Pass: '{pass_col}'")
            total_users = len(df)

            for index, row in df.iterrows():
                if not self.keep_running:
                    self.log("🛑 Process Stopped by User.")
                    break

                user_id = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                dob = row[dob_col] if dob_col and pd.notna(row[dob_col]) else None

                self.app.update_progress_safe_demand((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                status, reason = self.process_single_user(user_id, password, dob)

                if isinstance(reason, dict):
                    entry = {
                        "PAN":                      user_id,
                        "Status":                   status,
                        "Worklist_Status":          reason.get("Worklist_Status", ""),
                        "Worklist_Items":           reason.get("Worklist_Items", ""),
                        "Outstanding_Demand_Status": reason.get("Outstanding_Demand_Status", ""),
                        "Outstanding_Demand_Items":  reason.get("Outstanding_Demand_Items", ""),
                        "Timestamp":                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                else:
                    entry = {
                        "PAN": user_id, "Status": status, "Details": reason,
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                
                # Create user folder and save individual report
                name = reason.get("TaxpayerName", user_id) if isinstance(reason, dict) else user_id
                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "Income Tax Downloaded", "Demand Checker")
                folder_name = f"{user_id}_{name}"
                final_path = create_unique_folder(download_root, folder_name)
                
                self.report_data.append(entry)
                self._save_user_report(entry, final_path)
                self.log("-" * 40)

            self.generate_report()
            self.app.update_progress_safe_demand(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe_demand("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe_demand("Critical Error Occurred")

    def _save_user_report(self, entry, folder_path):
        try:
            if not folder_path or not os.path.exists(folder_path): return
            pan = entry.get("PAN", "unknown")
            report_path = os.path.join(folder_path, f"Report_{pan}.xlsx")
            pd.DataFrame([entry]).to_excel(report_path, index=False)
            self.log(f"   📄 User Report saved: Report_{pan}.xlsx")
        except Exception as e:
            self.log(f"   ⚠️ Failed to save user report: {e}")

    def generate_report(self):
        try:
            if not self.report_data: return
            df_report = pd.DataFrame(self.report_data)
            # Ensure column order
            col_order = [
                "PAN", "Status",
                "Worklist_Status", "Worklist_Items",
                "Outstanding_Demand_Status", "Outstanding_Demand_Items",
                "Timestamp"
            ]
            for c in col_order:
                if c not in df_report.columns:
                    df_report[c] = ""
            df_report = df_report[[c for c in col_order if c in df_report.columns]]
            filename = f"Demand_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_dir = os.path.join(os.getcwd(), "Income Tax Downloaded", "Demand Checker", "reports")
            os.makedirs(report_dir, exist_ok=True)
            report_path = os.path.join(report_dir, filename)
            df_report.to_excel(report_path, index=False)
            self.log(f"📄 Report saved: {report_path}")
        except: pass

    def process_single_user(self, user_id, password, dob):
        driver = None
        try:
            # --- CHROME OPTIONS & ANTI-BOT CONFIG ---
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument("--disable-blink-features=AutomationControlled")
            prefs = {
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
            }
            options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            driver.implicitly_wait(10)

            wait = WebDriverWait(driver, 20)

            # --- LOGIN WITH COMPREHENSIVE RETRY ---
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

                    # Step 1: Enter PAN
                    pan_entered = False
                    for pan_retry in range(3):
                        try:
                            pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                            pan_field.clear()
                            pan_field.send_keys(user_id)
                            pan_entered = True
                            break
                        except:
                            if pan_retry == 2: raise
                            time.sleep(1)

                    if not pan_entered: continue
                    time.sleep(0.5)

                    # Step 2: Click Continue
                    for cont_retry in range(3):
                        try:
                            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                            driver.execute_script("arguments[0].click();", continue_btn)
                            break
                        except:
                            if cont_retry == 2: raise
                            time.sleep(1)

                    time.sleep(1.5)
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN"

                    # Step 3: Enter Password
                    for pwd_retry in range(3):
                        try:
                            pwd_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                            pwd_field.clear()
                            pwd_field.send_keys(password)
                            break
                        except:
                            if pwd_retry == 2: raise
                            time.sleep(1)

                    try:
                        driver.execute_script("document.getElementById('passwordCheckBox-input').click();")
                        time.sleep(0.3)
                    except: pass

                    self.log("   ⏳ Waiting for security check (3s)...")
                    time.sleep(3.5)

                    # Step 4: Submit login
                    for submit_retry in range(3):
                        try:
                            driver.execute_script("document.querySelector('button.large-button-primary').click();")
                            break
                        except:
                            if submit_retry == 2: raise
                            time.sleep(1)

                    # Step 5: Wait for dashboard
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

            name_from_header = get_taxpayer_name(driver, fallback=user_id)
            if name_from_header != user_id:
                self.log(f"   👤 Taxpayer Name: {name_from_header}")

            self.log("   ✅ Login Successful! Reached Dashboard.")

            worklist_data        = "No items in worklist"
            outstanding_data     = "No outstanding demand"
            worklist_raw         = []
            outstanding_raw      = []

            # ──────────────────────────────────────────────────────────
            # SECTION 1 ▸ Pending Actions → Worklist
            # ──────────────────────────────────────────────────────────
            def click_pending_actions():
                for attempt in range(3):
                    try:
                        pa_btn = wait.until(EC.element_to_be_clickable((
                            By.XPATH,
                            "//span[contains(@class,'mdc-button__label') "
                            "and normalize-space(text())='Pending Actions']"
                        )))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", pa_btn)
                        time.sleep(0.4)
                        driver.execute_script("arguments[0].click();", pa_btn)
                        self.log("   ✅ 'Pending Actions' menu opened.")
                        return True
                    except Exception as e:
                        if attempt == 2:
                            self.log(f"   ❌ Could not open 'Pending Actions': {str(e)[:60]}")
                            return False
                        self.log(f"   ⚠️ Retry {attempt+1}/3 for 'Pending Actions'...")
                        time.sleep(1.5)
                return False

            # --- Step 1: Click Pending Actions ---
            self.log("   🔹 Step 1: Clicking 'Pending Actions'...")
            if not click_pending_actions():
                return "Failed", "Could not open Pending Actions"

            # --- Step 2: Click Worklist ---
            self.log("   🔹 Step 2: Clicking 'Worklist'...")
            worklist_clicked = False
            for attempt in range(3):
                try:
                    wl_item = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class,'cdk-overlay-container')]"
                        "//span[normalize-space(text())='Worklist'] | "
                        "//span[normalize-space(text())='Worklist']"
                    )))
                    driver.execute_script("arguments[0].click();", wl_item)
                    self.log("   ✅ 'Worklist' clicked.")
                    worklist_clicked = True
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Could not click 'Worklist': {str(e)[:60]}")
                    else:
                        self.log(f"   ⚠️ Retry {attempt+1}/3 for 'Worklist'...")
                        time.sleep(1.5)

            if worklist_clicked:
                time.sleep(3)  # Let the page fully render

                # --- Step 3: Check Worklist content ---
                self.log("   🔹 Step 3: Reading Worklist...")
                try:
                    # Check for "no items" message
                    no_items_els = driver.find_elements(
                        By.XPATH,
                        "//h4[contains(normalize-space(text()),'There is no item in worklist')]"
                    )
                    if no_items_els:
                        worklist_data = "No items in worklist"
                        self.log("   ℹ️ Worklist: No items found.")
                    else:
                        # Scrape rows — try table rows first, then list items / cards
                        rows = driver.find_elements(By.XPATH, "//table//tr[td]")
                        if not rows:
                            rows = driver.find_elements(By.XPATH,
                                "//div[contains(@class,'worklist') or contains(@class,'list-item') "
                                "or contains(@class,'card') or contains(@class,'action-item')]")

                        scraped = []
                        for r in rows:
                            txt = r.text.strip()
                            if txt:
                                scraped.append(txt)

                        if scraped:
                            worklist_raw = scraped
                            worklist_data = f"{len(scraped)} item(s) found"
                            self.log(f"   📋 Worklist: {len(scraped)} item(s) scraped.")
                            for i, item in enumerate(scraped[:5], 1):
                                self.log(f"      [{i}] {item[:100]}")
                        else:
                            worklist_data = "Items may exist but could not be scraped"
                            self.log("   ⚠️ Worklist: Page loaded but rows could not be read.")
                except Exception as e:
                    worklist_data = f"Error reading worklist: {str(e)[:60]}"
                    self.log(f"   ⚠️ Worklist scrape error: {str(e)[:60]}")

            # ──────────────────────────────────────────────────────────
            # SECTION 2 ▸ Pending Actions → Response to Outstanding Demand
            # ──────────────────────────────────────────────────────────

            # --- Step 4: Re-open Pending Actions ---
            self.log("   🔹 Step 4: Re-opening 'Pending Actions'...")
            time.sleep(1)
            if not click_pending_actions():
                return "Partial", f"Worklist: {worklist_data} | Outstanding: Could not re-open Pending Actions"

            # --- Step 5: Click Response to Outstanding Demand ---
            self.log("   🔹 Step 5: Clicking 'Response to Outstanding Demand'...")
            outstanding_clicked = False
            for attempt in range(3):
                try:
                    od_item = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class,'cdk-overlay-container')]"
                        "//span[contains(normalize-space(text()),'Response to Outstanding Demand')] | "
                        "//span[contains(normalize-space(text()),'Response to Outstanding Demand')]"
                    )))
                    driver.execute_script("arguments[0].click();", od_item)
                    self.log("   ✅ 'Response to Outstanding Demand' clicked.")
                    outstanding_clicked = True
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Could not click 'Response to Outstanding Demand': {str(e)[:60]}")
                    else:
                        self.log(f"   ⚠️ Retry {attempt+1}/3 for 'Outstanding Demand'...")
                        time.sleep(1.5)

            if outstanding_clicked:
                time.sleep(3)

                # --- Step 6: Check Outstanding Demand content ---
                self.log("   🔹 Step 6: Reading Outstanding Demand...")
                try:
                    no_demand_signals = [
                        "//h4[contains(normalize-space(text()),'no outstanding demand')]",
                        "//p[contains(normalize-space(text()),'no outstanding demand')]",
                        "//div[contains(normalize-space(text()),'No records found')]",
                        "//span[contains(normalize-space(text()),'No records found')]",
                    ]
                    found_empty = any(driver.find_elements(By.XPATH, xp) for xp in no_demand_signals)

                    if found_empty:
                        outstanding_data = "No outstanding demand"
                        self.log("   ℹ️ Outstanding Demand: Nothing found.")
                    else:
                        # Each demand card starts with div.innerBoxHeader
                        # We climb to its parent to access both header + stepper
                        card_headers = driver.find_elements(
                            By.CSS_SELECTOR, "div.innerBoxHeader"
                        )

                        if not card_headers:
                            outstanding_data = "No outstanding demand"
                            self.log("   ℹ️ Outstanding Demand: No cards found.")
                        else:
                            records = []
                            for card in card_headers:
                                try:
                                    # ── Demand Reference No ──────────────────
                                    try:
                                        ref_el = card.find_element(
                                            By.CSS_SELECTOR, "span.heading5.mNoWrap"
                                        )
                                        demand_ref = ref_el.text.strip()
                                    except:
                                        demand_ref = "N/A"

                                    # ── Assessment Year ──────────────────────
                                    try:
                                        ay_el = card.find_element(
                                            By.CSS_SELECTOR, "div.ass_yr_spacing span.heading5"
                                        )
                                        assessment_year = ay_el.text.strip()
                                    except:
                                        assessment_year = "N/A"

                                    # ── Stepper data (status + dates) ────────
                                    # The stepper lives in a sibling section.pipeline
                                    # inside the same parent row as the card header
                                    current_status      = "N/A"
                                    response_submitted  = "N/A"
                                    date_demand_raised  = "N/A"

                                    try:
                                        parent = card.find_element(By.XPATH, "./ancestor::div[@class and contains(@class,'row')][1]")
                                        stepper = parent.find_element(
                                            By.CSS_SELECTOR, "mat-vertical-stepper"
                                        )
                                        step_headers = stepper.find_elements(
                                            By.CSS_SELECTOR, "mat-step-header"
                                        )
                                        for step_hdr in step_headers:
                                            # heading label for this step
                                            try:
                                                heading = step_hdr.find_element(
                                                    By.CSS_SELECTOR, "section.dataHeading"
                                                ).text.strip().lower()
                                            except:
                                                heading = ""

                                            if "current status" in heading:
                                                try:
                                                    current_status = step_hdr.find_element(
                                                        By.CSS_SELECTOR, "mat-label.statusValue"
                                                    ).text.strip()
                                                except: pass

                                            elif "response submitted" in heading:
                                                try:
                                                    response_submitted = step_hdr.find_element(
                                                        By.CSS_SELECTOR, "section.subtitle2"
                                                    ).text.strip()
                                                except: pass

                                            elif "date of demand raised" in heading:
                                                try:
                                                    date_demand_raised = step_hdr.find_element(
                                                        By.CSS_SELECTOR, "section.subtitle2"
                                                    ).text.strip()
                                                except: pass
                                    except: pass

                                    rec = {
                                        "Demand_Ref_No":        demand_ref,
                                        "Assessment_Year":      assessment_year,
                                        "Current_Status":       current_status,
                                        "Response_Submitted":   response_submitted,
                                        "Date_Demand_Raised":   date_demand_raised,
                                    }
                                    records.append(rec)
                                    self.log(
                                        f"      📌 Ref: {demand_ref} | AY: {assessment_year} "
                                        f"| Status: {current_status} | Resp: {response_submitted} "
                                        f"| Raised: {date_demand_raised}"
                                    )
                                except Exception as ce:
                                    self.log(f"      ⚠️ Card parse error: {str(ce)[:60]}")

                            if records:
                                outstanding_raw = [
                                    f"Ref:{r['Demand_Ref_No']} AY:{r['Assessment_Year']} "
                                    f"Status:{r['Current_Status']} RespDate:{r['Response_Submitted']} "
                                    f"Raised:{r['Date_Demand_Raised']}"
                                    for r in records
                                ]
                                outstanding_data = f"{len(records)} demand(s) found"
                                self.log(f"   📋 Outstanding Demand: {len(records)} record(s) extracted.")
                            else:
                                outstanding_data = "No outstanding demand"
                                self.log("   ℹ️ Outstanding Demand: Cards found but data unreadable.")

                except Exception as e:
                    outstanding_data = f"Error reading demand: {str(e)[:60]}"
                    self.log(f"   ⚠️ Outstanding Demand scrape error: {str(e)[:60]}")

            # ──────────────────────────────────────────────────────────
            # Compose final result dict (returned to run() → report)
            # ──────────────────────────────────────────────────────────
            self.log(f"   📦 Worklist: {worklist_data}")
            self.log(f"   📦 Outstanding Demand: {outstanding_data}")

            return "Success", {
                "Worklist_Status":          worklist_data,
                "Worklist_Items":           " | ".join(worklist_raw) if worklist_raw else "",
                "Outstanding_Demand_Status": outstanding_data,
                "Outstanding_Demand_Items":  " | ".join(outstanding_raw) if outstanding_raw else "",
                "TaxpayerName":             name_from_header
            }

        except Exception as e:
            return "Failed", f"Browser Error: {str(e)[:40]}"
        finally:
            if driver: driver.quit()


# ============================================================
#  MAIN APP GUI
# ============================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Suite Pro - Challan Downloader")
        self.geometry("900x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.worker = None
        self.manual_credentials = []

        # --- Header ---
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 5))
        self.title_label = ctk.CTkLabel(self.header_frame, text="AUTOMATION SUITE PRO", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(side="left")

        # --- Main content area ---
        self.tab_challan = ctk.CTkFrame(self, fg_color="transparent")
        self.tab_challan.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.tab_challan.grid_columnconfigure(0, weight=1)
        self.tab_challan.grid_rowconfigure(0, weight=1)

        # SCROLLABLE CONTAINER
        self.scroll_container = ctk.CTkScrollableFrame(self.tab_challan, fg_color="transparent")
        self.scroll_container.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll_container.grid_columnconfigure(0, weight=1)
        self.tab_challan.grid_rowconfigure(0, weight=1)
        self.tab_challan.grid_rowconfigure(1, weight=0)

        self._build_challan_ui()

    def _build_challan_ui(self):
        self.excel_file_path = ""
        self.config_frame = ctk.CTkFrame(self.scroll_container)
        self.config_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))

        ctk.CTkLabel(self.config_frame, text="1. CREDENTIALS SOURCE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        
        f_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file = ctk.CTkEntry(f_frame, placeholder_text="Add PAN, Password, DOB manually...")
        self.entry_file.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_actions = ctk.CTkFrame(f_frame, fg_color="transparent")
        btn_actions.pack(side="right")
        # Add ID first
        ctk.CTkButton(btn_actions, text="➕ Add ID Password", command=self.add_id_password, width=150, fg_color="#059669", hover_color="#047857", font=("Segoe UI", 12, "bold")).pack(side="left")
        # View and Delete next
        self.btn_view_id = ctk.CTkButton(btn_actions, text="👁 View ID", command=self.view_saved_user, width=95, fg_color="#475569", hover_color="#334155", font=("Segoe UI", 11, "bold"))
        self.btn_view_id.pack(side="left", padx=(5, 0))
        self.btn_delete_id = ctk.CTkButton(btn_actions, text="🗑 Delete ID", command=self.delete_saved_user, width=105, fg_color="#7C3AED", hover_color="#6D28D9", font=("Segoe UI", 11, "bold"))
        self.btn_delete_id.pack(side="left", padx=(5, 0))
        # Demo last
        ctk.CTkButton(btn_actions, text="▶ Demo", command=self.open_demo_link, width=80, fg_color="#DC2626", hover_color="#B91C1C", font=("Segoe UI", 12, "bold")).pack(side="left", padx=(5, 0))
        self.btn_view_id.configure(state="disabled")
        self.btn_delete_id.configure(state="disabled")

        pref_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(pref_frame, text="Assessment Year:", text_color="gray").pack(side="left", padx=(0, 10))
        # Provide last 5 financial years (newest first). Change these values if you want a different range.
        fy_values = ["2027-2028", "2026-2027", "2025-2026", "2024-2025", "2023-2024", "2022-2023"]
        self.combo_filter = ctk.CTkComboBox(pref_frame, values=fy_values, width=250, state="readonly")
        self.combo_filter.set(fy_values[0])
        self.combo_filter.pack(side="left")

        # Log UI
        self.log_frame = ctk.CTkFrame(self.scroll_container)
        self.log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 5))
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self.log_frame, text="2. LIVE LOG", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(2, 2))
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), activate_scrollbars=True, height=100)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box.configure(state="disabled")
        
        self.progress = ctk.CTkProgressBar(self.log_frame, mode="determinate")
        self.progress.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress.set(0)

        btn_row_footer = ctk.CTkFrame(self.tab_challan, fg_color="transparent")
        btn_row_footer.grid(row=1, column=0, sticky="ew", padx=20, pady=(5, 20))
        self.btn_start = ctk.CTkButton(btn_row_footer, text="START CHALLAN DOWNLOAD", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=self.start_process)
        self.btn_start.pack(side="left", expand=True, fill="x")
        self.btn_stop = ctk.CTkButton(btn_row_footer, text="⏹ STOP", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#DC2626", hover_color="#B91C1C", command=self.stop_process, width=150)
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_stop.pack_forget()
        self.btn_open_folder = ctk.CTkButton(btn_row_footer, text="📂 OPEN FOLDER", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#2563EB", hover_color="#1D4ED8", command=self.open_output_folder, width=180)
        self.btn_open_folder.pack(side="left", padx=(10, 0))
        self.btn_open_folder.pack_forget()

    # --- GUI Handlers ---
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
        webbrowser.open_new_tab("https://youtu.be/doam0_V3zFc")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.excel_file_path = filename
            self.manual_credentials = []
            self._refresh_manual_controls()
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, filename)
            self.log_to_gui(f"File Loaded: {os.path.basename(filename)}")

    def _get_saved_user_id(self):
        if not self.manual_credentials:
            return ""
        return str(self.manual_credentials[0].get("PAN", "")).strip()

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
            user_id = str(item.get("PAN", "")).strip()
            password = str(item.get("Password", "")).strip()
            dob = str(item.get("DOB", "")).strip()
            if user_id and password:
                rows.append({"PAN": user_id, "Password": password, "DOB": dob})

        if not rows:
            return ""

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="it_challan_manual_") as tmp:
            temp_excel = tmp.name
        pd.DataFrame(rows, columns=["PAN", "Password", "DOB"]).to_excel(temp_excel, index=False)
        return temp_excel

    def start_process(self):
        excel_path = self.excel_file_path
        if not excel_path and self.manual_credentials:
            excel_path = self._create_manual_excel()

        if not excel_path:
            return messagebox.showwarning("Error", "Add ID/Password first")

        self.btn_start.configure(state="disabled", text="PROCESSING...", fg_color="gray")
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_open_folder.pack_forget()
        self.progress.set(0)
        self.worker = ChallanWorker(self, excel_path, self.combo_filter.get())
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if self.worker:
            self.worker.keep_running = False
        self.btn_stop.configure(state="disabled", text="Stopping...")

    # --- CHALLAN UI Safe Updaters ---
    def log_to_gui(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def update_log_safe(self, msg): 
        self.after(0, lambda: self.log_to_gui(msg))
        
    def update_progress_safe(self, val): 
        self.after(0, lambda: self.progress.set(val))
        
    def process_finished_safe(self, msg):
        def _finish():
            self.log_to_gui(f"\nSTATUS: {msg}")
            self.btn_start.configure(state="normal", text="START CHALLAN DOWNLOAD", fg_color="#2563EB")
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.btn_stop.pack_forget()
            self.btn_open_folder.pack(side="left", padx=(10, 0))
            messagebox.showinfo("Done", msg)
        self.after(0, _finish)

    def open_output_folder(self):
        try:
            target = os.path.join(os.getcwd(), "Income Tax Downloaded", "Challan Downloader")
            if not os.path.exists(target):
                target = os.path.join(os.getcwd(), "Income Tax Downloaded")
            os.startfile(target)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()