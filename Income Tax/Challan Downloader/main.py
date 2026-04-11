import threading
import time
import os
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
                download_root = os.path.join(base_dir, "Challan_Downloads")
                final_path = create_unique_folder(download_root, user_id)

                status, reason = self.process_single_user(user_id, password, dob, final_path)
                
                self.report_data.append({
                    "PAN": user_id, "Status": status, "Details": reason,
                    "Folder Saved": os.path.basename(final_path),
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                self.log("-" * 40)
            
            self.generate_report()
            self.app.update_progress_safe(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe("Critical Error Occurred")

    def generate_report(self):
        try:
            if not self.report_data: return
            df_report = pd.DataFrame(self.report_data)
            filename = f"Challan_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df_report.to_excel(filename, index=False)
            self.log(f"📄 Report saved: {filename}")
        except: pass

    def process_single_user(self, user_id, password, dob, download_folder):
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
            
            # Set aggressive timeouts
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            driver.implicitly_wait(10)
            
            wait = WebDriverWait(driver, 20)
            actions = ActionChains(driver)

            # --- 1. LOGIN WITH COMPREHENSIVE RETRY ---
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
            # ==========================================================
            if self.year_mode == "Current Year":
                selected_years = all_years[:1]
            elif self.year_mode == "Last 2 Years":
                selected_years = all_years[:2]
            else:  # "All History"
                selected_years = all_years

            self.log(f"   🎯 Selected Years to Download ({self.year_mode}): {', '.join(selected_years)}")

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
                        time.sleep(4)  # Wait for file download to initiate

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
            return ("Success" if downloaded_years else "Failed"), summary

        except Exception as e: 
            return "Failed", f"Browser Error: {str(e)[:20]}"
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
                        "PAN":                      user_id,
                        "Status":                   status,
                        "Worklist_Status":          reason,
                        "Worklist_Items":           "",
                        "Outstanding_Demand_Status": "",
                        "Outstanding_Demand_Items":  "",
                        "Timestamp":                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                self.report_data.append(entry)
                self.log("-" * 40)

            self.generate_report()
            self.app.update_progress_safe_demand(1.0)
            self.log("\n✅ BATCH COMPLETED!")
            self.app.process_finished_safe_demand("All Tasks Completed.")

        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}")
            self.app.process_finished_safe_demand("Critical Error Occurred")

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
            df_report.to_excel(filename, index=False)
            self.log(f"📄 Report saved: {filename}")
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

        # --- Header ---
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.title_label = ctk.CTkLabel(self.header_frame, text="AUTOMATION SUITE PRO", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(side="left")

        # --- Main content area ---
        self.tab_challan = ctk.CTkFrame(self, fg_color="transparent")
        self.tab_challan.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.tab_challan.grid_columnconfigure(0, weight=1)
        self.tab_challan.grid_rowconfigure(1, weight=1)

        self._build_challan_ui()

    def _build_challan_ui(self):
        self.excel_file_path = ""
        self.config_frame = ctk.CTkFrame(self.tab_challan)
        self.config_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))

        ctk.CTkLabel(self.config_frame, text="1. CREDENTIALS SOURCE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        
        f_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file = ctk.CTkEntry(f_frame, placeholder_text="Excel File (Headers: PAN, Password, DOB)...")
        self.entry_file.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_frame, text="BROWSE", command=self.browse_file, width=80).pack(side="right")
        ctk.CTkButton(f_frame, text="▶ Demo", command=self.open_demo_link, width=80, fg_color="#e53935", hover_color="#b71c1c", font=("Arial", 12, "bold")).pack(side="right", padx=(0, 5))
        ctk.CTkButton(f_frame, text="📥 Sample", command=self.download_sample, width=100, fg_color="#43a047", hover_color="#2e7d32", font=("Arial", 12, "bold")).pack(side="right", padx=(0, 5))

        pref_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(pref_frame, text="Download Filter:", text_color="gray").pack(side="left", padx=(0, 10))
        self.combo_filter = ctk.CTkComboBox(pref_frame, values=["Current Year", "Last 2 Years", "All History"], width=250, state="readonly")
        self.combo_filter.set("Current Year")
        self.combo_filter.pack(side="left")

        # Log UI
        self.log_frame = ctk.CTkFrame(self.tab_challan)
        self.log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 5))
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self.log_frame, text="2. LIVE LOG", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(5, 5))
        self.log_box = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), activate_scrollbars=True)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box.configure(state="disabled")
        
        self.progress = ctk.CTkProgressBar(self.log_frame, mode="determinate")
        self.progress.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress.set(0)

        btn_row_footer = ctk.CTkFrame(self.tab_challan, fg_color="transparent")
        btn_row_footer.grid(row=2, column=0, sticky="ew", padx=20, pady=(5, 20))
        self.btn_start = ctk.CTkButton(btn_row_footer, text="START CHALLAN DOWNLOAD", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=self.start_process)
        self.btn_start.pack(side="left", expand=True, fill="x")
        self.btn_stop = ctk.CTkButton(btn_row_footer, text="⏹ STOP", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#c62828", hover_color="#8e0000", command=self.stop_process, width=150)
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.btn_stop.pack_forget()

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
        webbrowser.open_new_tab("https://www.youtube.com/watch?v=XXXXXXXXXX")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.excel_file_path = filename
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, filename)
            self.log_to_gui(f"File Loaded: {os.path.basename(filename)}")

    def start_process(self):
        if not self.excel_file_path:
            return messagebox.showwarning("Error", "Select an Excel file first")
        self.btn_start.configure(state="disabled", text="PROCESSING...", fg_color="gray")
        self.btn_stop.pack(side="left", padx=(10, 0))
        self.progress.set(0)
        self.worker = ChallanWorker(self, self.excel_file_path, self.combo_filter.get())
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
            self.btn_start.configure(state="normal", text="START CHALLAN DOWNLOAD", fg_color="#1f538d")
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.btn_stop.pack_forget()
            messagebox.showinfo("Done", msg)
        self.after(0, _finish)

if __name__ == "__main__":
    app = App()
    app.mainloop()