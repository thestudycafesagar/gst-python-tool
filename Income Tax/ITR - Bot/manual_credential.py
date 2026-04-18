import sys
import time
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QPushButton, QLabel, QLineEdit, QGroupBox, QTextEdit)
from PySide6.QtCore import QThread, Signal, Qt

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (UnexpectedAlertPresentException, NoAlertPresentException, 
                                        TimeoutException, StaleElementReferenceException)
from webdriver_manager.chrome import ChromeDriverManager

# --- WORKER THREAD (The Engine) ---
class LoginWorker(QThread):
    log_signal = Signal(str)           
    finished_signal = Signal(str)      

    def __init__(self, user_id, password):
        super().__init__()
        self.user_id = user_id
        self.password = password

    def run(self):
        self.log_signal.emit("Initializing Browser...")
        
        # --- BROWSER CONFIGURATION ---
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        options.add_experimental_option("detach", True) 

        # Bypass Download Prompts
        prefs = {
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        options.add_experimental_option("prefs", prefs)

        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            wait = WebDriverWait(driver, 20)
            actions = ActionChains(driver) 
        except Exception as e:
            self.log_signal.emit(f"Failed to start browser: {e}")
            return

        try:
            # ============================================================
            # STEP 1: ENTER USER ID (PAN)
            # ============================================================
            self.log_signal.emit("Navigating to Portal...")
            driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")

            try:
                time.sleep(1)
                driver.switch_to.alert.accept()
            except NoAlertPresentException:
                pass 

            self.log_signal.emit(f"Entering User ID: {self.user_id}")
            try:
                pan_field = wait.until(EC.visibility_of_element_located((By.ID, "panAdhaarUserId")))
                pan_field.clear()
                pan_field.send_keys(self.user_id)
            except UnexpectedAlertPresentException:
                driver.switch_to.alert.accept()
                pan_field = driver.find_element(By.ID, "panAdhaarUserId")
                pan_field.send_keys(self.user_id)

            continue_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
            continue_btn.click()

            # --- CHECK FOR INVALID PAN ERROR ---
            time.sleep(1.5)
            try:
                pan_error = driver.find_elements(By.XPATH, "//*[contains(text(), 'PAN does not exist')]")
                if pan_error and pan_error[0].is_displayed():
                    self.log_signal.emit("❌ ERROR: Invalid PAN Card Number. Stopping.")
                    self.finished_signal.emit("Failed: Invalid PAN")
                    return 
            except: pass

            # ============================================================
            # STEP 2: ENTER PASSWORD
            # ============================================================
            self.log_signal.emit("Entering Password...")
            try:
                password_field = wait.until(EC.visibility_of_element_located((By.ID, "loginPasswordField")))
                password_field.clear()
                password_field.send_keys(self.password)
            except TimeoutException:
                self.log_signal.emit("❌ Password field not found. (Check if PAN is valid)")
                return

            try:
                checkbox = wait.until(EC.presence_of_element_located((By.ID, "passwordCheckBox-input")))
                driver.execute_script("arguments[0].click();", checkbox)
            except:
                self.log_signal.emit("Checkbox click issue (ignoring)")

            self.log_signal.emit("Clicking Login...")
            login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.large-button-primary")))
            time.sleep(1.0) 
            driver.execute_script("arguments[0].click();", login_btn)

            # ============================================================
            # STEP 3: LOGIN VALIDATION (STRICT CHECK)
            # ============================================================
            self.log_signal.emit("Verifying Login Status...")
            
            # We wait up to 5 seconds specifically looking for the error OR a successful login indicator
            for _ in range(5):
                time.sleep(1)
                
                # CHECK 1: INVALID PASSWORD
                try:
                    # Using a more general text search for "Invalid Password" to be safe
                    pass_error = driver.find_elements(By.XPATH, "//*[contains(text(), 'Invalid Password')]")
                    if pass_error and pass_error[0].is_displayed():
                        self.log_signal.emit("❌ ERROR: Wrong Password detected!")
                        self.finished_signal.emit("Failed: Wrong Password")
                        return # HARD STOP
                except: pass

                # CHECK 2: REQUEST NOT AUTHENTICATED (RETRY)
                try:
                    error_msg = driver.find_elements(By.XPATH, "//span[contains(text(), 'Request is not authenticated')]")
                    if error_msg and error_msg[0].is_displayed():
                        self.log_signal.emit(">> 'Request not authenticated' detected. Retrying Click...")
                        driver.execute_script("arguments[0].click();", login_btn)
                except: pass
                
                # CHECK 3: DUAL LOGIN
                try:
                    login_here_btn = driver.find_elements(By.XPATH, "//button[contains(@class, 'primaryButton') and contains(text(), 'Login Here')]")
                    if login_here_btn and login_here_btn[0].is_displayed():
                        self.log_signal.emit(">> 'Dual Login' detected. Clicking Login Here...")
                        driver.execute_script("arguments[0].click();", login_here_btn[0])
                        time.sleep(2)
                except: pass
                
                # CHECK 4: SUCCESS (Dashboard Element)
                try:
                    # If "e-File" menu appears, we are successfully logged in
                    if driver.find_elements(By.ID, "e-File"):
                        self.log_signal.emit("✅ Login Successful!")
                        break # Break the validation loop and proceed
                except: pass

            # FINAL SAFETY CHECK: If we still see the password field, login failed.
            if driver.find_elements(By.ID, "loginPasswordField"):
                 # Check for error one last time
                try:
                    pass_error = driver.find_elements(By.XPATH, "//*[contains(text(), 'Invalid Password')]")
                    if pass_error and pass_error[0].is_displayed():
                        self.log_signal.emit("❌ ERROR: Wrong Password detected (Final Check).")
                        self.finished_signal.emit("Failed: Wrong Password")
                        return 
                except: pass
                
                # If we are here, we are likely stuck on the password page
                self.log_signal.emit("⚠️ Stuck on Password Page. Stopping to prevent crash.")
                self.finished_signal.emit("Failed: Login Stuck")
                return

            # ============================================================
            # STEP 4: NAVIGATE TO VIEW FILED RETURNS
            # ============================================================
            self.log_signal.emit("Navigating to 'View Filed Returns'...")
            try:
                # Wait for Dashboard
                e_file_menu = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "e-File")))
                e_file_menu.click()
                time.sleep(1) 

                # Hover Income Tax Returns
                itr_submenu = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Income Tax Returns')]")))
                actions.move_to_element(itr_submenu).perform()
                time.sleep(1) 

                # Click View Filed Returns
                view_filed_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'View Filed Returns')]")))
                view_filed_btn.click()
                
            except Exception as nav_e:
                self.log_signal.emit(f"Navigation Error (Login might have failed): {nav_e}")
                return

            # ============================================================
            # STEP 5: DOWNLOAD FILES
            # ============================================================
            self.log_signal.emit("Scanning for Filed Returns...")
            
            try:
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "contextBox")))
                time.sleep(2) 
                
                cards = driver.find_elements(By.CLASS_NAME, "contextBox")
                self.log_signal.emit(f"Found {len(cards)} total filings.")

                process_count = min(len(cards), 3)

                for i in range(process_count):
                    cards = driver.find_elements(By.CLASS_NAME, "contextBox")
                    current_card = cards[i]
                    
                    try:
                        year_text = current_card.find_element(By.CLASS_NAME, "contentHeadingText").text
                    except:
                        year_text = f"Year {i+1}"

                    self.log_signal.emit(f"--- Processing: {year_text} ---")

                    def download_file(classname, name):
                        try:
                            btn = current_card.find_element(By.CSS_SELECTOR, f".{classname}")
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                            time.sleep(0.5)
                            driver.execute_script("arguments[0].click();", btn)
                            self.log_signal.emit(f"   -> Started {name} Download")
                            time.sleep(0.5) 
                        except:
                            self.log_signal.emit(f"   -> {name} not available")

                    download_file("dformback", "Form")
                    download_file("drecback", "Receipt")
                    download_file("dxmlback", "JSON")
                    
                    self.log_signal.emit("-----------------------------")

                self.log_signal.emit("✅ SUCCESS: All downloads initiated.")
                self.finished_signal.emit("Task Completed Successfully")

            except Exception as e:
                self.log_signal.emit(f"Error during download phase: {str(e)}")

        except Exception as e:
            self.log_signal.emit(f"Critical Error: {str(e)}")

# --- MAIN GUI WINDOW ---
class IncomeTaxApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Income Tax Downloader")
        self.resize(550, 650)
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

        # Inputs
        cred_group = QGroupBox("1. Enter Credentials")
        cred_layout = QVBoxLayout()
        self.lbl_user = QLabel("User ID / PAN:")
        self.input_user = QLineEdit()
        self.lbl_pass = QLabel("Password:")
        self.input_pass = QLineEdit()
        self.input_pass.setEchoMode(QLineEdit.Password) 
        cred_layout.addWidget(self.lbl_user)
        cred_layout.addWidget(self.input_user)
        cred_layout.addWidget(self.lbl_pass)
        cred_layout.addWidget(self.input_pass)
        cred_group.setLayout(cred_layout)
        layout.addWidget(cred_group)

        # Log
        log_group = QGroupBox("2. Status Log")
        log_layout = QVBoxLayout()
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("background-color: #f0f0f0; border: 1px solid #ccc;")
        log_layout.addWidget(self.log_box)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # Start Button
        self.btn_start = QPushButton("START DOWNLOAD")
        self.btn_start.clicked.connect(self.start_login)
        self.btn_start.setStyleSheet("background-color: #2196F3; color: white;")
        layout.addWidget(self.btn_start)

    def log(self, text):
        timestamp = time.strftime("[%H:%M:%S] ")
        self.log_box.append(timestamp + text)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    def start_login(self):
        user_id = self.input_user.text().strip()
        password = self.input_pass.text().strip()
        if not user_id or not password:
            self.log("ERROR: Please enter both User ID and Password.")
            return

        self.btn_start.setEnabled(False)
        self.worker = LoginWorker(user_id, password)
        self.worker.log_signal.connect(self.log)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.start()

    def process_finished(self, msg):
        self.log(f"DONE: {msg}")
        self.btn_start.setEnabled(True)

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = IncomeTaxApp()
    window.show()
    sys.exit(app.exec())