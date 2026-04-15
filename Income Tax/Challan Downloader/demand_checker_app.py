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
from selenium.webdriver.support.ui import WebDriverWait
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
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                    window.navigator.chrome = { runtime: {} };
                    Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
                    Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
                """
            })
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            driver.implicitly_wait(10)

            wait = WebDriverWait(driver, 20)

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
                    if "does not exist" in driver.page_source: return "Failed", "Invalid PAN"

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

            worklist_data    = "No items in worklist"
            outstanding_data = "No outstanding demand"
            worklist_raw     = []
            outstanding_raw  = []

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

            self.log("   🔹 Step 1: Clicking 'Pending Actions'...")
            if not click_pending_actions():
                return "Failed", "Could not open Pending Actions"

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
                time.sleep(3)
                self.log("   🔹 Step 3: Reading Worklist...")
                try:
                    no_items_els = driver.find_elements(
                        By.XPATH,
                        "//h4[contains(normalize-space(text()),'There is no item in worklist')]"
                    )
                    if no_items_els:
                        worklist_data = "No items in worklist"
                        self.log("   ℹ️ Worklist: No items found.")
                    else:
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

            self.log("   🔹 Step 4: Re-opening 'Pending Actions'...")
            time.sleep(1)
            if not click_pending_actions():
                return "Partial", f"Worklist: {worklist_data} | Outstanding: Could not re-open Pending Actions"

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
                        card_headers = driver.find_elements(By.CSS_SELECTOR, "div.innerBoxHeader")

                        if not card_headers:
                            outstanding_data = "No outstanding demand"
                            self.log("   ℹ️ Outstanding Demand: No cards found.")
                        else:
                            records = []
                            for card in card_headers:
                                try:
                                    try:
                                        ref_el = card.find_element(By.CSS_SELECTOR, "span.heading5.mNoWrap")
                                        demand_ref = ref_el.text.strip()
                                    except:
                                        demand_ref = "N/A"

                                    try:
                                        ay_el = card.find_element(By.CSS_SELECTOR, "div.ass_yr_spacing span.heading5")
                                        assessment_year = ay_el.text.strip()
                                    except:
                                        assessment_year = "N/A"

                                    current_status     = "N/A"
                                    response_submitted = "N/A"
                                    date_demand_raised = "N/A"

                                    try:
                                        parent = card.find_element(By.XPATH, "./ancestor::div[@class and contains(@class,'row')][1]")
                                        stepper = parent.find_element(By.CSS_SELECTOR, "mat-vertical-stepper")
                                        step_headers = stepper.find_elements(By.CSS_SELECTOR, "mat-step-header")
                                        for step_hdr in step_headers:
                                            try:
                                                heading = step_hdr.find_element(By.CSS_SELECTOR, "section.dataHeading").text.strip().lower()
                                            except:
                                                heading = ""

                                            if "current status" in heading:
                                                try:
                                                    current_status = step_hdr.find_element(By.CSS_SELECTOR, "mat-label.statusValue").text.strip()
                                                except: pass

                                            elif "response submitted" in heading:
                                                try:
                                                    response_submitted = step_hdr.find_element(By.CSS_SELECTOR, "section.subtitle2").text.strip()
                                                except: pass

                                            elif "date of demand raised" in heading:
                                                try:
                                                    date_demand_raised = step_hdr.find_element(By.CSS_SELECTOR, "section.subtitle2").text.strip()
                                                except: pass
                                    except: pass

                                    rec = {
                                        "Demand_Ref_No":      demand_ref,
                                        "Assessment_Year":    assessment_year,
                                        "Current_Status":     current_status,
                                        "Response_Submitted": response_submitted,
                                        "Date_Demand_Raised": date_demand_raised,
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

            self.log(f"   📦 Worklist: {worklist_data}")
            self.log(f"   📦 Outstanding Demand: {outstanding_data}")

            return "Success", {
                "Worklist_Status":           worklist_data,
                "Worklist_Items":            " | ".join(worklist_raw) if worklist_raw else "",
                "Outstanding_Demand_Status": outstanding_data,
                "Outstanding_Demand_Items":  " | ".join(outstanding_raw) if outstanding_raw else "",
            }

        except Exception as e:
            return "Failed", f"Browser Error: {str(e)[:40]}"
        finally:
            if driver: driver.quit()


# ============================================================
#  MAIN APP GUI — DEMAND CHECKER (standalone)
# ============================================================
class DemandCheckerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Suite Pro - Demand Checker")
        self.geometry("900x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.demand_worker = None
        self.manual_credentials = []

        # --- Header ---
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        ctk.CTkLabel(header_frame, text="AUTOMATION SUITE PRO",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(side="left")

        # --- Content area ---
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(1, weight=1)

        self._build_demand_checker_ui()

    def _build_demand_checker_ui(self):
        self.excel_file_path_demand = ""

        config_frame = ctk.CTkFrame(self.content)
        config_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))

        ctk.CTkLabel(config_frame, text="1. CREDENTIALS SOURCE",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))

        f_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file_demand = ctk.CTkEntry(f_frame, placeholder_text="Add PAN, Password, DOB manually...")
        self.entry_file_demand.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_actions = ctk.CTkFrame(f_frame, fg_color="transparent")
        btn_actions.pack(side="right")
        ctk.CTkButton(btn_actions, text="▶ Demo", command=self.open_demo_link, width=80, fg_color="#DC2626", hover_color="#B91C1C", font=("Segoe UI", 12, "bold")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(btn_actions, text="➕ Add ID Password", command=self.add_id_password, width=150, fg_color="#059669", hover_color="#047857", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.btn_view_id = ctk.CTkButton(btn_actions, text="👁 View ID", command=self.view_saved_user, width=95, fg_color="#475569", hover_color="#334155", font=("Segoe UI", 11, "bold"))
        self.btn_view_id.pack(side="left", padx=(5, 0))
        self.btn_delete_id = ctk.CTkButton(btn_actions, text="🗑 Delete ID", command=self.delete_saved_user, width=105, fg_color="#7C3AED", hover_color="#6D28D9", font=("Segoe UI", 11, "bold"))
        self.btn_delete_id.pack(side="left", padx=(5, 0))
        self.btn_view_id.configure(state="disabled")
        self.btn_delete_id.configure(state="disabled")

        # Log UI
        log_frame = ctk.CTkFrame(self.content)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 5))
        log_frame.grid_rowconfigure(1, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(log_frame, text="2. LIVE LOG",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(5, 5))
        self.log_box_demand = ctk.CTkTextbox(log_frame, font=("Consolas", 12), activate_scrollbars=True)
        self.log_box_demand.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box_demand.configure(state="disabled")

        self.progress_demand = ctk.CTkProgressBar(log_frame, mode="determinate")
        self.progress_demand.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_demand.set(0)

        btn_footer = ctk.CTkFrame(self.content, fg_color="transparent")
        btn_footer.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))
        btn_footer.grid_columnconfigure(0, weight=1)
        self.btn_start_demand = ctk.CTkButton(btn_footer, text="START DEMAND CHECKER",
                                              font=ctk.CTkFont(size=16, weight="bold"), height=50,
                                              command=self.start_process_demand)
        self.btn_start_demand.grid(row=0, column=0, sticky="ew")
        self.btn_stop = ctk.CTkButton(btn_footer, text="⏹ STOP", font=ctk.CTkFont(size=16, weight="bold"),
                                      height=50, fg_color="#DC2626", hover_color="#B91C1C",
                                      command=self.stop_process, width=150)
        self.btn_stop.grid(row=0, column=1, padx=(10, 0))
        self.btn_stop.grid_remove()
        self.btn_open_folder_demand = ctk.CTkButton(btn_footer, text="📂 OPEN FOLDER", font=ctk.CTkFont(size=16, weight="bold"),
                                      height=50, fg_color="#2563EB", hover_color="#1D4ED8",
                                      command=self.open_output_folder_demand, width=180)
        self.btn_open_folder_demand.grid(row=0, column=2, padx=(10, 0))
        self.btn_open_folder_demand.grid_remove()

    # --- GUI Handlers ---
    def download_sample(self):
        import shutil
        import os
        from tkinter import messagebox, filedialog
        sample_path = os.path.join(os.path.dirname(__file__), "Income Tax Sample File.xlsx")
        
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

    def browse_file_demand(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.excel_file_path_demand = filename
            self.manual_credentials = []
            self._refresh_manual_controls()
            self.entry_file_demand.delete(0, "end")
            self.entry_file_demand.insert(0, filename)
            self.log_to_gui_demand(f"File Loaded: {os.path.basename(filename)}")

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
            self.entry_file_demand.delete(0, "end")
            self.entry_file_demand.insert(0, f"Selected ID: {user_id}")

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
        self.entry_file_demand.delete(0, "end")
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
        ent_pass = ctk.CTkEntry(card, placeholder_text="Enter Password", show="*")
        ent_pass.pack(fill="x", pady=(4, 10))

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
            self.excel_file_path_demand = ""
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

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="it_demand_manual_") as tmp:
            temp_excel = tmp.name
        pd.DataFrame(rows, columns=["PAN", "Password", "DOB"]).to_excel(temp_excel, index=False)
        return temp_excel

    def start_process_demand(self):
        excel_path = self.excel_file_path_demand
        if not excel_path and self.manual_credentials:
            excel_path = self._create_manual_excel()

        if not excel_path:
            return messagebox.showwarning("Error", "Add ID/Password first")

        self.btn_start_demand.configure(state="disabled", text="PROCESSING...", fg_color="gray")
        self.btn_stop.grid()
        self.btn_open_folder_demand.grid_remove()
        self.progress_demand.set(0)
        self.demand_worker = DemandCheckerWorker(self, excel_path)
        threading.Thread(target=self.demand_worker.run, daemon=True).start()

    def stop_process(self):
        if self.demand_worker:
            self.demand_worker.keep_running = False
        self.btn_stop.configure(state="disabled", text="Stopping...")

    # --- Safe Updaters ---
    def log_to_gui_demand(self, msg):
        self.log_box_demand.configure(state="normal")
        self.log_box_demand.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box_demand.see("end")
        self.log_box_demand.configure(state="disabled")

    def update_log_safe_demand(self, msg):
        self.after(0, lambda: self.log_to_gui_demand(msg))

    def update_progress_safe_demand(self, val):
        self.after(0, lambda: self.progress_demand.set(val))

    def process_finished_safe_demand(self, msg):
        def _finish():
            self.log_to_gui_demand(f"\nSTATUS: {msg}")
            self.btn_start_demand.configure(state="normal", text="START DEMAND CHECKER", fg_color="#2563EB")
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.btn_stop.grid_remove()
            self.btn_open_folder_demand.grid(row=0, column=2, padx=(10, 0))
            messagebox.showinfo("Done", msg)
        self.after(0, _finish)

    def open_output_folder_demand(self):
        try:
            target = os.getcwd()
            os.startfile(target)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {e}")


if __name__ == "__main__":
    app = DemandCheckerApp()
    app.mainloop()
