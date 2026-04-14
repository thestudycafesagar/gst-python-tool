import threading
import time
import os
import tempfile
import pandas as pd
import customtkinter as ctk
from datetime import datetime
from tkinter import filedialog, messagebox
from functools import wraps

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
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

        ctk.CTkLabel(self, text=f"⚠️ User: {user_id}\nSelect Years to Download:",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=20)

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

        ctk.CTkButton(self, text="CONFIRM & DOWNLOAD", command=self.on_confirm,
                      fg_color="green", hover_color="darkgreen", height=40).pack(pady=20)

    def on_confirm(self):
        selected = [year for year, var in self.check_vars.items() if var.get() != "off"]
        if not selected:
            messagebox.showwarning("Warning", "Please select at least one year!")
            return
        self.callback(selected)
        self.grab_release()
        self.destroy()


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
#  WORKER: FILED RETURN / REFUND CHECKER
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
            df = pd.read_excel(self.excel_path)
            user_col, pass_col, dob_col = normalize_columns(df)

            if not user_col or not pass_col:
                self.log("❌ ERROR: Headers missing. Need 'PAN' and 'Password'.")
                self.app.process_finished_safe_filed("Failed: Column Header Error")
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

                self.app.update_progress_safe_filed((index) / total_users)
                self.log(f"🔹 [{index+1}/{total_users}] PROCESSING USER: {user_id}")

                status, reason = self.process_single_user(user_id, password, dob)

                self.log(f"   📊 Result: {status} - {reason}")
                self.log("-" * 40)

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
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            driver.implicitly_wait(10)

            wait = WebDriverWait(driver, 20)
            actions = ActionChains(driver)

            # LOGIN WITH COMPREHENSIVE RETRY
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
                        except Exception as e:
                            if pan_retry == 2:
                                self.log(f"   ⚠️ Failed to enter PAN after 3 tries")
                                raise
                            time.sleep(1)

                    if not pan_entered: continue
                    time.sleep(0.5)

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
                        except Exception as e:
                            if submit_retry == 2:
                                self.log(f"   ⚠️ Failed to submit login after 3 tries")
                                raise
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

            # NAVIGATE TO VIEW FILED RETURNS
            self.log("   🚀 Navigating to View Filed Returns...")
            nav_success = False
            for nav_attempt in range(1, 4):
                try:
                    if nav_attempt > 1:
                        self.log(f"   ⚠️ Navigation Retry {nav_attempt}/3...")
                        time.sleep(2)

                    self.log("   📂 Clicking e-File menu...")
                    efile = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[@class='mdc-button__label' and contains(text(), 'e-File')]"))
                    )
                    driver.execute_script("arguments[0].click();", efile)
                    time.sleep(1)

                    self.log("   📋 Hovering over Income Tax Returns...")
                    itr_menu = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//button[contains(@class, 'mat-mdc-menu-item')]//span[contains(text(), 'Income Tax Returns')]"))
                    )
                    actions.move_to_element(itr_menu).perform()
                    time.sleep(0.5)

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

            # DATA EXTRACTION
            self.log("   📊 Starting data extraction...")
            time.sleep(2)

            try:
                try:
                    filing_count_elem = driver.find_element(By.CLASS_NAME, "filingCount")
                    self.log(f"   📄 {filing_count_elem.text}")
                except:
                    self.log("   ⚠️ Could not find filing count")

                self.log("   🔽 Setting pagination to show all records...")
                try:
                    pagination_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "paginatorselect")))
                    pagination_select = pagination_field.find_element(By.TAG_NAME, "mat-select")
                    driver.execute_script("arguments[0].click();", pagination_select)
                    time.sleep(1)
                    options = driver.find_elements(By.XPATH, "//mat-option[@role='option']")
                    if options:
                        driver.execute_script("arguments[0].click();", options[-1])
                        self.log("   ✅ Set pagination to maximum records")
                        time.sleep(2)
                except Exception as e:
                    self.log(f"   ⚠️ Pagination selection skipped: {str(e)[:30]}")

                self.log("   🔍 Extracting Assessment Years...")
                year_elements = driver.find_elements(By.XPATH, "//mat-label[@class='contentHeadingText']")
                available_years = [elem.text.strip() for elem in year_elements if "A.Y." in elem.text.strip()]

                if not available_years:
                    self.log("   ⚠️ No filed returns found")
                    return "Success", "No Filed Returns Found"

                self.log(f"   📋 Found {len(available_years)} Assessment Years: {', '.join(available_years)}")

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

                self.log("   📥 Extracting return details...")
                cards = driver.find_elements(By.XPATH, "//mat-card[contains(@class, 'contextBox')]")

                extracted_data = []
                for idx, card in enumerate(cards):
                    try:
                        ay = card.find_element(By.CLASS_NAME, "contentHeadingText").text.strip()
                        if ay not in years_to_extract:
                            continue

                        filing_type = card.find_element(By.CLASS_NAME, "leftSideVal").text.strip()

                        first_status = "N/A"
                        first_date = "N/A"
                        try:
                            status_divs = card.find_elements(By.CLASS_NAME, "matStepStatus")
                            date_divs = card.find_elements(By.CLASS_NAME, "matStepDate")
                            if status_divs: first_status = status_divs[0].text.strip()
                            if date_divs: first_date = date_divs[0].text.strip()
                        except: pass

                        itr_type = ack_no = filed_by = filing_date = filing_section = "N/A"
                        try:
                            right_labels = card.find_elements(By.CLASS_NAME, "rightsideLabel")
                            right_values = card.find_elements(By.CLASS_NAME, "fieldVal")
                            for i, label_elem in enumerate(right_labels):
                                label = label_elem.text.strip().lower()
                                if i < len(right_values):
                                    value = right_values[i].text.strip()
                                    if "itr" in label: itr_type = value
                                    elif "acknowledgement" in label: ack_no = value
                                    elif "filed by" in label: filed_by = value
                                    elif "filing date" in label: filing_date = value
                                    elif "filing section" in label: filing_section = value
                        except: pass

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
            if driver: driver.quit()

    def generate_report(self):
        try:
            if not self.report_data:
                self.log("⚠️ No data to generate report")
                return

            df_report = pd.DataFrame(self.report_data)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Filed_Return_Report_{timestamp}.xlsx"
            self.log(f"📝 Generating report: {filename}")

            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df_report.to_excel(writer, index=False, sheet_name='Filed Returns')
                worksheet = writer.sheets['Filed Returns']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except: pass
                    worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

            self.log(f"✅ Report saved successfully: {filename}")
            self.log(f"📊 Total records in report: {len(df_report)}")

        except Exception as e:
            self.log(f"❌ Report generation error: {str(e)}")


# ============================================================
#  MAIN APP GUI — REFUND CHECKER (standalone)
# ============================================================
class RefundCheckerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Suite Pro - Refund Checker")
        self.geometry("900x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.worker = None
        self.manual_credentials = []

        # --- Header ---
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        ctk.CTkLabel(header_frame, text="AUTOMATION SUITE PRO",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(side="left")

        # --- Content area ---
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=1, column=0, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(1, weight=1)

        self._build_ui()

    def _build_ui(self):
        self.excel_file_path_filed = ""

        config_frame = ctk.CTkFrame(self.content)
        config_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))

        ctk.CTkLabel(config_frame, text="1. CREDENTIALS SOURCE",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))

        f_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        f_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.entry_file_filed = ctk.CTkEntry(f_frame, placeholder_text="Add PAN and Assessment Year manually...")
        self.entry_file_filed.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_actions = ctk.CTkFrame(f_frame, fg_color="transparent")
        btn_actions.pack(side="right")
        ctk.CTkButton(btn_actions, text="▶ Demo", command=self.open_demo_link, width=80, fg_color="#e53935", hover_color="#b71c1c", font=("Arial", 12, "bold")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(btn_actions, text="➕ Add ID Password", command=self.add_id_password, width=150, fg_color="#43a047", hover_color="#2e7d32", font=("Arial", 12, "bold")).pack(side="left")
        self.btn_view_id = ctk.CTkButton(btn_actions, text="👁 View ID", command=self.view_saved_user, width=95, fg_color="#546e7a", hover_color="#37474f", font=("Arial", 11, "bold"))
        self.btn_view_id.pack(side="left", padx=(5, 0))
        self.btn_delete_id = ctk.CTkButton(btn_actions, text="🗑 Delete ID", command=self.delete_saved_user, width=105, fg_color="#8e24aa", hover_color="#6a1b9a", font=("Arial", 11, "bold"))
        self.btn_delete_id.pack(side="left", padx=(5, 0))
        self.btn_view_id.configure(state="disabled")
        self.btn_delete_id.configure(state="disabled")

        pref_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        pref_frame.pack(fill="x", padx=15, pady=(5, 15))
        ctk.CTkLabel(pref_frame, text="Extract Data for:", text_color="gray").pack(side="left", padx=(0, 10))
        self.combo_years_filed = ctk.CTkComboBox(
            pref_frame,
            values=["Current Year", "Current and Last Year", "Current and Last 2 Years", "Manual Selection (Popup)"],
            width=250, state="readonly"
        )
        self.combo_years_filed.set("Current Year")
        self.combo_years_filed.pack(side="left")

        # Log UI
        log_frame = ctk.CTkFrame(self.content)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 5))
        log_frame.grid_rowconfigure(1, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(log_frame, text="2. LIVE LOG",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(5, 5))
        self.log_box_filed = ctk.CTkTextbox(log_frame, font=("Consolas", 12), activate_scrollbars=True)
        self.log_box_filed.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 10))
        self.log_box_filed.configure(state="disabled")

        self.progress_filed = ctk.CTkProgressBar(log_frame, mode="determinate")
        self.progress_filed.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
        self.progress_filed.set(0)

        btn_footer = ctk.CTkFrame(self.content, fg_color="transparent")
        btn_footer.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))
        btn_footer.grid_columnconfigure(0, weight=1)
        self.btn_start_filed = ctk.CTkButton(btn_footer, text="GENERATE REPORT",
                                             font=ctk.CTkFont(size=16, weight="bold"), height=50,
                                             command=self.start_process)
        self.btn_start_filed.grid(row=0, column=0, sticky="ew")
        self.btn_stop = ctk.CTkButton(btn_footer, text="⏹ STOP", font=ctk.CTkFont(size=16, weight="bold"),
                                      height=50, fg_color="#c62828", hover_color="#8e0000",
                                      command=self.stop_process, width=150)
        self.btn_stop.grid(row=0, column=1, padx=(10, 0))
        self.btn_stop.grid_remove()

    # --- GUI Handlers ---
    def download_sample(self):
        sample_path = os.path.join(os.path.dirname(__file__), "Income Tax Sample File.xlsx")
        if not os.path.exists(sample_path):
            messagebox.showerror("Error", f"Sample file not found:\n{sample_path}")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Income Tax Sample File.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All Files", "*.*")],
        )
        if not save_path:
            return

        try:
            import shutil

            shutil.copy2(sample_path, save_path)
            messagebox.showinfo("Success", f"Sample downloaded to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download sample:\n{e}")

    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://www.youtube.com/watch?v=XXXXXXXXXX")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.excel_file_path_filed = filename
            self.manual_credentials = []
            self._refresh_manual_controls()
            self.entry_file_filed.delete(0, "end")
            self.entry_file_filed.insert(0, filename)
            self.log_to_gui_filed(f"File Loaded: {os.path.basename(filename)}")

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
            self.entry_file_filed.delete(0, "end")
            self.entry_file_filed.insert(0, f"Selected ID: {user_id}")

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
        self.entry_file_filed.delete(0, "end")
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
            self.excel_file_path_filed = ""
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

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="it_refund_manual_") as tmp:
            temp_excel = tmp.name
        pd.DataFrame(rows, columns=["PAN", "Password", "DOB"]).to_excel(temp_excel, index=False)
        return temp_excel

    def start_process(self):
        excel_path = self.excel_file_path_filed
        if not excel_path and self.manual_credentials:
            excel_path = self._create_manual_excel()

        if not excel_path:
            return messagebox.showwarning("Error", "Select file or add ID/Password first")

        self.btn_start_filed.configure(state="disabled", text="GENERATING...", fg_color="gray")
        self.btn_stop.grid()
        self.progress_filed.set(0)
        self.worker = FiledReturnWorker(self, excel_path, self.combo_years_filed.get())
        threading.Thread(target=self.worker.run, daemon=True).start()

    def stop_process(self):
        if self.worker:
            self.worker.keep_running = False
        self.btn_stop.configure(state="disabled", text="Stopping...")

    def trigger_year_selection(self, years_list, user_id, callback):
        self.after(0, lambda: YearSelectionPopup(self, years_list, user_id, callback))

    # --- Safe Updaters ---
    def log_to_gui_filed(self, msg):
        self.log_box_filed.configure(state="normal")
        self.log_box_filed.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box_filed.see("end")
        self.log_box_filed.configure(state="disabled")

    def update_log_safe_filed(self, msg):
        self.after(0, lambda: self.log_to_gui_filed(msg))

    def update_progress_safe_filed(self, val):
        self.after(0, lambda: self.progress_filed.set(val))

    def process_finished_safe_filed(self, msg):
        def _finish():
            self.log_to_gui_filed(f"\nSTATUS: {msg}")
            self.btn_start_filed.configure(state="normal", text="GENERATE REPORT", fg_color="#1f538d")
            self.btn_stop.configure(state="normal", text="⏹ STOP")
            self.btn_stop.grid_remove()
            messagebox.showinfo("Done", msg)
        self.after(0, _finish)


if __name__ == "__main__":
    app = RefundCheckerApp()
    app.mainloop()
