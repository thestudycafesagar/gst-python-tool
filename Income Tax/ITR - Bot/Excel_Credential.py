import sys
import time
import os
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QPushButton, QLabel, QLineEdit, QGroupBox, QTextEdit, QFileDialog)
from PySide6.QtCore import QThread, Signal, Qt

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (UnexpectedAlertPresentException, NoAlertPresentException, 
                                        TimeoutException)
from webdriver_manager.chrome import ChromeDriverManager

# --- WORKER THREAD (The Engine) ---
class IncomeTaxWorker(QThread):
    log_signal = Signal(str)           
    progress_signal = Signal(int)
    finished_signal = Signal(str)      

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self.keep_running = True

    def run(self):
        self.log_signal.emit("🚀 Reading Excel File...")
        
        try:
            # 1. READ EXCEL
            df = pd.read_excel(self.excel_path)
            df.columns = df.columns.str.strip() 
            
            if 'User ID' not in df.columns or 'Password' not in df.columns:
                self.log_signal.emit("❌ Error: Excel must have 'User ID' and 'Password' columns.")
                self.finished_signal.emit("Error: Bad Excel Format")
                return

            total_users = len(df)
            self.log_signal.emit(f"📂 Found {total_users} users. Starting Batch Process...")

            # 2. LOOP THROUGH USERS
            for index, row in df.iterrows():
                if not self.keep_running: break
                
                user_id = str(row['User ID']).strip()
                password = str(row['Password']).strip()
                
                self.log_signal.emit(f"\n==========================================")
                self.log_signal.emit(f"👤 Processing User ({index+1}/{total_users}): {user_id}")
                self.log_signal.emit(f"==========================================")

                # --- NEW FOLDER CREATION LOGIC (Unique Folders) ---
                base_dir = os.getcwd()
                download_root = os.path.join(base_dir, "Downloads")
                
                # Ensure main Downloads folder exists
                if not os.path.exists(download_root):
                    os.makedirs(download_root)

                # Determine unique folder name
                folder_name = user_id
                counter = 1
                final_path = os.path.join(download_root, folder_name)
                
                # Loop to check if folder exists, if so, append (1), (2), etc.
                while os.path.exists(final_path):
                    folder_name = f"{user_id}({counter})"
                    final_path = os.path.join(download_root, folder_name)
                    counter += 1
                
                # Create the unique folder
                os.makedirs(final_path)
                user_folder = final_path
                self.log_signal.emit(f"📂 New unique folder created: {folder_name}")

                # START BROWSER FOR THIS USER
                self.process_single_user(user_id, password, user_folder)
                
            self.log_signal.emit("\n✅ Batch Processing Completed!")
            self.finished_signal.emit("All Done")

        except Exception as e:
            self.log_signal.emit(f"❌ Critical Error: {str(e)}")

    def process_single_user(self, user_id, password, download_folder):
        driver = None
        try:
            # --- BROWSER CONFIG ---
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            options.add_argument("--disable-blink-features=AutomationControlled")
            
            prefs = {
                "download.default_directory": download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_setting_values.automatic_downloads": 1
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
            wait = WebDriverWait(driver, 15)
            actions = ActionChains(driver)

            # ============================================================
            # OUTER RETRY LOOP (Try the whole login process up to 3 times)
            # ============================================================
            login_success = False
            
            for login_attempt in range(1, 4): # Try 1, 2, 3
                if login_success: break
                
                if login_attempt > 1:
                    self.log_signal.emit(f"⚠️ Login Attempt {login_attempt}/3: Refreshing and Retrying...")
                    driver.delete_all_cookies()
                    driver.refresh()
                    time.sleep(3)

                try:
                    # --- STEP 1: LOGIN PAGE ---
                    self.log_signal.emit("🌐 Navigating to Portal...")
                    driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")

                    try:
                        time.sleep(1)
                        driver.switch_to.alert.accept()
                    except: pass

                    # Enter User ID
                    self.log_signal.emit("🔑 Entering User ID...")
                    pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                    pan_field.clear()
                    pan_field.send_keys(user_id)
                    
                    btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                    btn.click()

                    # Check for Invalid PAN immediately
                    time.sleep(1.5)
                    if "does not exist" in driver.page_source:
                        self.log_signal.emit("❌ Error: Invalid PAN. Stopping User.")
                        return # Stop everything for this user

                    # Enter Password
                    self.log_signal.emit("🔑 Entering Password...")
                    pass_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                    pass_field.clear()
                    pass_field.send_keys(password)
                    
                    try:
                        cb = driver.find_element(By.ID, "passwordCheckBox-input")
                        driver.execute_script("arguments[0].click();", cb)
                    except: pass
                    
                    # Human-like pause before submitting
                    self.log_signal.emit("⏳ Waiting 4s before submitting...")
                    time.sleep(4) 
                    
                    # Click Login
                    login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
                    driver.execute_script("arguments[0].click();", login_btn)

                    # --- STEP 2: VALIDATION LOOP (30s) ---
                    self.log_signal.emit("⏳ Verifying Login Status...")
                    
                    for _ in range(30):
                        time.sleep(1)
                        
                        # A. Success
                        if driver.find_elements(By.ID, "e-File"):
                            self.log_signal.emit("✅ Login Successful!")
                            login_success = True
                            break
                        
                        # B. Wrong Password (Don't retry)
                        if "Invalid Password" in driver.page_source:
                            self.log_signal.emit("❌ Error: Wrong Password.")
                            return 
                        
                        # C. Dual Login (Fix it)
                        try:
                            dual_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual_btn and dual_btn[0].is_displayed():
                                self.log_signal.emit("⚠️ Dual Session. Taking Control...")
                                driver.execute_script("arguments[0].click();", dual_btn[0])
                                time.sleep(3)
                        except: pass

                        # D. Retry "Request Not Authenticated"
                        try:
                            if "Request is not authenticated" in driver.page_source:
                                 # Only retry clicking occasionally
                                 login_btns = driver.find_elements(By.CSS_SELECTOR, "button.large-button-primary")
                                 if login_btns:
                                     driver.execute_script("arguments[0].click();", login_btns[0])
                        except: pass

                    if login_success: break # Break outer retry loop

                except Exception as e:
                    self.log_signal.emit(f"⚠️ Attempt {login_attempt} failed: {str(e)}")
            
            # --- END OF RETRY LOOP ---

            if not login_success:
                self.log_signal.emit("❌ Failed to Login after 3 attempts. Skipping User.")
                return

            # --- STEP 3: NAVIGATION ---
            self.log_signal.emit("🚀 Navigating to Returns...")
            try:
                # Extra wait for dashboard to settle
                time.sleep(2)
                e_file = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "e-File")))
                e_file.click()
                
                submenu = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Income Tax Returns')]")))
                actions.move_to_element(submenu).perform()
                
                view_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'View Filed Returns')]")))
                view_btn.click()
            except Exception as nav_e:
                self.log_signal.emit(f"❌ Navigation Failed: {nav_e}")
                return

            # --- STEP 4: DOWNLOAD ---
            self.log_signal.emit("⬇️ Downloading Files...")
            try:
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "contextBox")))
                time.sleep(2)
                cards = driver.find_elements(By.CLASS_NAME, "contextBox")
                
                count = min(len(cards), 3)
                
                for i in range(count):
                    cards = driver.find_elements(By.CLASS_NAME, "contextBox") # Refresh
                    card = cards[i]
                    try:
                        year = card.find_element(By.CLASS_NAME, "contentHeadingText").text
                    except: year = f"Year {i+1}"
                    
                    self.log_signal.emit(f"   📄 {year}")
                    
                    def click_dl(cls, name):
                        try:
                            btn = card.find_element(By.CSS_SELECTOR, f".{cls}")
                            driver.execute_script("arguments[0].click();", btn)
                            self.log_signal.emit(f"      -> {name}")
                            time.sleep(0.5)
                        except: pass

                    click_dl("dformback", "Form")
                    click_dl("drecback", "Receipt")
                    click_dl("dxmlback", "JSON")
                
                self.log_signal.emit("✅ Downloads Initiated.")
                time.sleep(5) 
                
            except Exception as e:
                self.log_signal.emit(f"❌ Download Error: {e}")

        except Exception as e:
            self.log_signal.emit(f"❌ Browser Error: {e}")
        finally:
            if driver:
                driver.quit()
                self.log_signal.emit("🔄 Browser Closed.")

    def stop(self):
        self.keep_running = False

# --- MAIN GUI WINDOW ---
class IncomeTaxApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Income Tax Bulk Downloader Pro")
        self.resize(600, 700)
        self.worker = None

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.setStyleSheet("""
            QGroupBox { font-weight: bold; font-size: 14px; }
            QLineEdit { padding: 8px; font-size: 14px; }
            QPushButton { padding: 10px; font-size: 14px; font-weight: bold; }
            QTextEdit { font-family: Consolas; font-size: 12px; }
        """)

        # 1. File Upload
        file_group = QGroupBox("1. Load Credentials")
        file_layout = QVBoxLayout()
        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("Select Excel file with 'User ID' and 'Password' columns...")
        self.btn_browse = QPushButton("Browse Excel")
        self.btn_browse.clicked.connect(self.browse_file)
        self.btn_browse.setStyleSheet("background-color: #555; color: white;")
        
        file_layout.addWidget(self.file_input)
        file_layout.addWidget(self.btn_browse)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 2. Log
        log_group = QGroupBox("2. Process Log")
        log_layout = QVBoxLayout()
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("background-color: #f0f0f0; border: 1px solid #ccc;")
        log_layout.addWidget(self.log_box)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 3. Start Button
        self.btn_start = QPushButton("START BATCH DOWNLOAD")
        self.btn_start.clicked.connect(self.start_process)
        self.btn_start.setStyleSheet("background-color: #2196F3; color: white;")
        layout.addWidget(self.btn_start)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.file_input.setText(file_path)
            self.log(f"File Selected: {file_path}")

    def log(self, text):
        timestamp = time.strftime("[%H:%M:%S] ")
        self.log_box.append(timestamp + text)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    def start_process(self):
        excel_path = self.file_input.text().strip()
        if not excel_path:
            self.log("ERROR: Please select an Excel file first.")
            return

        self.btn_start.setEnabled(False)
        self.log("🚀 Starting Worker Thread...")
        
        self.worker = IncomeTaxWorker(excel_path)
        self.worker.log_signal.connect(self.log)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.start()

    def process_finished(self, msg):
        self.log(f"\nDONE: {msg}")
        self.btn_start.setEnabled(True)

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.terminate()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = IncomeTaxApp()
    window.show()
    sys.exit(app.exec())