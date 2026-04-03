import sys
import time
import pandas as pd
from io import BytesIO
from PIL import Image

# GUI Imports
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                               QTextEdit, QLineEdit, QGroupBox, QProgressBar)
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtGui import QPixmap, QImage

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException

# --- WORKER THREAD (The Engine) ---
class GSTWorker(QThread):
    log_signal = Signal(str)           
    captcha_signal = Signal(bytes, str) 
    finished_signal = Signal(str)      
    progress_signal = Signal(int)      

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.user_captcha_response = None
        self.is_waiting_for_captcha = False
        self.keep_running = True

    def run(self):
        self.log_signal.emit("Initializing Browser...")
        
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized") 
        options.add_argument("--disable-blink-features=AutomationControlled") 
        
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        except Exception as e:
            self.log_signal.emit(f"Failed to start browser: {e}")
            return

        try:
            # Load Excel
            df = pd.read_excel(self.file_path)
            df.columns = df.columns.str.strip() 
            
            if 'GSTIN' not in df.columns:
                self.log_signal.emit("Error: Excel must have a column named 'GSTIN'")
                driver.quit()
                return
            
            gstin_list = df['GSTIN'].astype(str).unique().tolist()
            results = []
            total = len(gstin_list)

            self.log_signal.emit(f"Found {total} unique GSTINs. Starting process...")

            for index, gstin in enumerate(gstin_list):
                if not self.keep_running: break
                
                self.log_signal.emit(f"Processing ({index+1}/{total}): {gstin}")
                self.progress_signal.emit(int(((index) / total) * 100))

                try:
                    # Navigate
                    driver.get("https://services.gst.gov.in/services/searchtp")
                    
                    # 1. Enter GSTIN
                    try:
                        input_box = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.ID, "for_gstin"))
                        )
                        input_box.clear()
                        input_box.send_keys(gstin)
                    except TimeoutException:
                        self.log_signal.emit("Error: Site took too long to load.")
                        continue

                    # --- CAPTCHA RETRY LOOP ---
                    while self.keep_running:
                        
                        # A. Fetch Captcha
                        self.log_signal.emit("Fetching Captcha...")
                        try:
                            # Wait for overlay to disappear
                            try:
                                WebDriverWait(driver, 2).until(
                                    EC.invisibility_of_element_located((By.CLASS_NAME, "dimmer-holder"))
                                )
                            except: pass

                            captcha_img = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "imgCaptcha"))
                            )
                        except:
                            # Fallback
                            captcha_img = driver.find_element(By.XPATH, "//img[contains(@src, 'captcha')]")

                        driver.execute_script("arguments[0].scrollIntoView();", captcha_img)
                        time.sleep(0.5) 
                        png_data = captcha_img.screenshot_as_png
                        
                        # B. Ask User (Popup)
                        self.user_captcha_response = None
                        self.is_waiting_for_captcha = True
                        self.captcha_signal.emit(png_data, gstin) 
                        
                        while self.user_captcha_response is None and self.keep_running:
                            time.sleep(0.1)
                        
                        self.is_waiting_for_captcha = False
                        if not self.keep_running: break

                        # C. Submit Captcha
                        self.log_signal.emit("Submitting...")
                        try:
                            captcha_box = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.ID, "fo-captcha"))
                            )
                            captcha_box.clear()
                            captcha_box.send_keys(self.user_captcha_response)
                            
                            # Force Click Search
                            search_btn = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.ID, "lotsearch"))
                            )
                            driver.execute_script("arguments[0].click();", search_btn)
                            
                        except Exception as e:
                            self.log_signal.emit(f"Error submitting: {e}")
                            break # Break retry loop

                        # D. STRICT VERIFICATION
                        self.log_signal.emit("Verifying...")
                        try:
                            # Wait for EITHER:
                            # 1. Success (Legal Name text)
                            # 2. Error (Invalid chars text)
                            # 3. Invalid GSTIN (Invalid text)
                            WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, 
                                    "//strong[contains(text(),'Legal Name')] | " +
                                    "//p[contains(text(),'Legal Name')] | " +
                                    "//div[contains(text(),'Enter valid characters')] | " +
                                    "//span[contains(text(),'Invalid')]"))
                            )
                        except:
                            # Timeout means nothing happened -> Likely silent failure -> Retry
                            self.log_signal.emit(" >> No response detected. Retrying...")
                            continue

                        page_source = driver.page_source

                        # CASE 1: SUCCESS (Must verify "Legal Name" exists)
                        if "Legal Name" in page_source:
                            self.log_signal.emit(" -> Captcha Correct! Extracting Data...")
                            
                            # --- DATA EXTRACTION ---
                            row_data = {"GSTIN": gstin}
                            
                            # Helper
                            def get_text(label):
                                xpaths = [
                                    f"//p[contains(text(),'{label}')]/following-sibling::p",
                                    f"//strong[contains(text(),'{label}')]/../following-sibling::p",
                                    f"//*[contains(text(),'{label}')]/following::p[1]" 
                                ]
                                for xpath in xpaths:
                                    try:
                                        el = driver.find_element(By.XPATH, xpath)
                                        if el.text.strip(): return el.text.strip()
                                    except: continue
                                return "N/A"

                            row_data["Legal Name"] = get_text("Legal Name of Business")
                            row_data["Trade Name"] = get_text("Trade Name")
                            row_data["Effective Date"] = get_text("Effective Date of registration")
                            row_data["Constitution"] = get_text("Constitution of Business")
                            row_data["Status"] = get_text("GSTIN / UIN")
                            row_data["Taxpayer Type"] = get_text("Taxpayer Type")
                            row_data["Address"] = get_text("Principal Place of Business")
                            
                            # Lists
                            def get_list_text(xpath):
                                try:
                                    items = driver.find_elements(By.XPATH, xpath)
                                    return ", ".join([x.text.strip() for x in items if x.text.strip()])
                                except: return "N/A"

                            row_data["Admin Office"] = get_list_text("//strong[contains(text(),'Administrative Office')]/parent::p/following-sibling::ul//li")
                            row_data["Other Office"] = get_list_text("//strong[contains(text(),'Other Office')]/parent::p/following-sibling::ul//li")
                            
                            row_data["Core Business"] = get_text("Nature Of Core Business Activity")
                            row_data["Business Activities"] = get_list_text("//p[contains(text(),'Nature of Business Activities')]/ancestor::div[@class='panel-heading']/following-sibling::div//li")

                            # HSN
                            try:
                                hsn_rows = []
                                rows = driver.find_elements(By.XPATH, "//div[contains(@class,'table-responsive')]//table[contains(@class,'tbl')]//tbody//tr")
                                for row in rows:
                                    cols = row.find_elements(By.TAG_NAME, "td")
                                    if len(cols) >= 2:
                                        t1, t2 = cols[0].text.strip(), cols[1].text.strip()
                                        if t1 or t2: hsn_rows.append(f"{t1}-{t2}")
                                    if len(cols) >= 4:
                                        t3, t4 = cols[2].text.strip(), cols[3].text.strip()
                                        if t3 or t4: hsn_rows.append(f"{t3}-{t4}")
                                row_data["Goods & Services"] = " | ".join(hsn_rows)
                            except:
                                row_data["Goods & Services"] = "N/A"

                            # Filing Tables
                            self.log_signal.emit(" -> Extracting Filing Data...")
                            try:
                                show_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "filingTable")))
                                driver.execute_script("arguments[0].click();", show_btn)
                                
                                search_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary.srchbtn")))
                                driver.execute_script("arguments[0].click();", search_btn)
                                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//h4[contains(text(),'GSTR3B')]")))
                                time.sleep(1)

                                def get_filing_history(header_text):
                                    history = []
                                    try:
                                        xpath = f"//h4[contains(text(),'{header_text}')]/ancestor::div[@class='table-responsive']//tbody/tr"
                                        rows = driver.find_elements(By.XPATH, xpath)
                                        for row in rows[:5]:
                                            cols = row.find_elements(By.TAG_NAME, "td")
                                            if len(cols) >= 4:
                                                history.append(f"[{cols[1].text}-{cols[0].text}: {cols[3].text} on {cols[2].text}]")
                                        return " | ".join(history)
                                    except: return "Not Found"

                                row_data["GSTR-3B History"] = get_filing_history("GSTR3B")
                                row_data["GSTR-1 History"] = get_filing_history("GSTR-1")
                            except:
                                row_data["GSTR-3B History"] = "Hidden/Error"
                                row_data["GSTR-1 History"] = "Hidden/Error"

                            # Frequency
                            self.log_signal.emit(" -> Extracting Return Frequency...")
                            try:
                                freq_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "profileTable")))
                                driver.execute_script("arguments[0].click();", freq_btn)
                                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'exp')]")))
                                time.sleep(1)
                                freq_data = []
                                freq_rows = driver.find_elements(By.XPATH, "//table[contains(@class,'exp')]//tbody//tr")
                                for row in freq_rows:
                                    cols = row.find_elements(By.TAG_NAME, "td")
                                    if len(cols) >= 9:
                                        yr = cols[0].text.strip()
                                        q1, f1 = cols[1].text.strip(), cols[2].text.strip()
                                        q2, f2 = cols[3].text.strip(), cols[4].text.strip()
                                        q3, f3 = cols[5].text.strip(), cols[6].text.strip()
                                        q4, f4 = cols[7].text.strip(), cols[8].text.strip()
                                        freq_data.append(f"[{yr}: {q1}({f1}), {q2}({f2}), {q3}({f3}), {q4}({f4})]")
                                row_data["Return Frequency"] = " | ".join(freq_data)
                            except:
                                row_data["Return Frequency"] = "Not Available"

                            self.log_signal.emit(f" -> Success: {row_data['Legal Name']}")
                            results.append(row_data)
                            break # Exit Retry Loop (Success)
                        
                        # CASE 2: INVALID GSTIN -> Fail
                        elif "Invalid GSTIN" in page_source or "No Records Found" in page_source:
                            self.log_signal.emit(" -> Result: Invalid Number")
                            results.append({"GSTIN": gstin, "Status": "Invalid GSTIN"})
                            break 

                        # CASE 3: WRONG CAPTCHA or RELOAD -> Retry
                        else:
                            self.log_signal.emit(" >> WRONG CAPTCHA or Page Reloaded! Please try again.")
                            time.sleep(1.5) 
                            continue 

                except Exception as e:
                    self.log_signal.emit(f"Error processing {gstin}: {str(e)}")
                    results.append({"GSTIN": gstin, "Status": "Error"})
            
            # Export with Unique Filename
            self.progress_signal.emit(100)
            
            # Generate Timestamp: e.g., GST_Report_2023-10-27_14-30-00.xlsx
            timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
            output_file = f"GST_Report_{timestamp}.xlsx"
            
            try:
                pd.DataFrame(results).to_excel(output_file, index=False)
                self.finished_signal.emit(f"Completed! Saved as {output_file}")
            except Exception as e:
                self.log_signal.emit(f"CRITICAL ERROR: Could not save report. {e}")
                self.finished_signal.emit("Failed to save")
            
        except Exception as e:
            self.log_signal.emit(f"Critical Error: {e}")
        finally:
            driver.quit()

    def receive_captcha_input(self, text):
        self.user_captcha_response = text

    def stop(self):
        self.keep_running = False


# --- MAIN GUI WINDOW ---
class GSTApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GST Bulk Verification Tool")
        self.resize(650, 750)
        self.file_path = None
        self.worker = None

        # Main Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Style
        self.setStyleSheet("""
            QGroupBox { font-weight: bold; font-size: 14px; }
            QPushButton { padding: 8px; font-size: 13px; }
            QLabel { font-size: 12px; }
        """)

        # 1. File Selection
        file_group = QGroupBox("1. Upload Data")
        file_layout = QHBoxLayout()
        self.path_label = QLineEdit()
        self.path_label.setPlaceholderText("Select Excel file with 'GSTIN' column...")
        self.path_label.setReadOnly(True)
        btn_browse = QPushButton("Browse File")
        btn_browse.clicked.connect(self.browse_file)
        file_layout.addWidget(self.path_label)
        file_layout.addWidget(btn_browse)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 2. Status Log
        log_group = QGroupBox("2. Process Log")
        log_layout = QVBoxLayout()
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("background-color: #f9f9f9; border: 1px solid #ccc;")
        log_layout.addWidget(self.log_box)
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        log_layout.addWidget(self.progress_bar)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 3. Captcha Section
        captcha_group = QGroupBox("3. CAPTCHA Action Required")
        captcha_layout = QVBoxLayout()
        
        self.lbl_instruction = QLabel("Waiting for process to start...")
        self.lbl_instruction.setAlignment(Qt.AlignCenter)
        self.lbl_instruction.setStyleSheet("font-weight: bold; color: #333;")
        
        self.captcha_image_label = QLabel()
        self.captcha_image_label.setAlignment(Qt.AlignCenter)
        self.captcha_image_label.setFixedSize(250, 100)
        self.captcha_image_label.setStyleSheet("border: 2px dashed #aaa; background: #e0e0e0;")
        
        self.captcha_input = QLineEdit()
        self.captcha_input.setPlaceholderText("Type Code & Press Enter")
        self.captcha_input.setStyleSheet("font-size: 16px; padding: 5px;")
        self.captcha_input.returnPressed.connect(self.submit_captcha)
        self.captcha_input.setEnabled(False)

        self.btn_submit_captcha = QPushButton("Submit Captcha")
        self.btn_submit_captcha.clicked.connect(self.submit_captcha)
        self.btn_submit_captcha.setEnabled(False)
        self.btn_submit_captcha.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")

        captcha_layout.addWidget(self.lbl_instruction)
        captcha_layout.addWidget(self.captcha_image_label, alignment=Qt.AlignCenter)
        captcha_layout.addWidget(self.captcha_input)
        captcha_layout.addWidget(self.btn_submit_captcha)
        captcha_group.setLayout(captcha_layout)
        layout.addWidget(captcha_group)

        # 4. Start Button
        self.btn_start = QPushButton("START PROCESSING")
        self.btn_start.setFixedHeight(50)
        self.btn_start.clicked.connect(self.start_process)
        self.btn_start.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px; font-weight: bold;")
        layout.addWidget(self.btn_start)

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.file_path = path
            self.path_label.setText(path)
            self.log("File loaded. Ready to start.")

    def log(self, text):
        self.log_box.append(text)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    def start_process(self):
        if not self.file_path:
            self.log("Please select a file first!")
            return

        self.btn_start.setEnabled(False)
        self.log("Starting Automation Thread...")
        
        self.worker = GSTWorker(self.file_path)
        self.worker.log_signal.connect(self.log)
        self.worker.captcha_signal.connect(self.display_captcha)
        self.worker.progress_signal.connect(self.progress_bar.setValue)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.start()

    def display_captcha(self, image_data, gstin):
        # Convert bytes to QImage
        img = QImage.fromData(image_data)
        pixmap = QPixmap.fromImage(img)
        
        # Update UI
        self.lbl_instruction.setText(f"Enter CAPTCHA for: {gstin}")
        self.lbl_instruction.setStyleSheet("font-weight: bold; color: #d32f2f; font-size: 14px;")
        
        self.captcha_image_label.setPixmap(pixmap.scaled(self.captcha_image_label.size(), Qt.KeepAspectRatio))
        self.captcha_input.setEnabled(True)
        self.btn_submit_captcha.setEnabled(True)
        self.captcha_input.clear()
        self.captcha_input.setFocus()
        
        # Bring window to front
        self.activateWindow()

    def submit_captcha(self):
        text = self.captcha_input.text()
        if not text: return
        
        if self.worker and self.worker.is_waiting_for_captcha:
            self.worker.receive_captcha_input(text)
            self.captcha_input.setEnabled(False)
            self.btn_submit_captcha.setEnabled(False)
            self.lbl_instruction.setText("Verifying...")
            self.lbl_instruction.setStyleSheet("color: #388E3C;")

    def process_finished(self, msg):
        self.log(f"\nDONE: {msg}")
        self.btn_start.setEnabled(True)
        self.lbl_instruction.setText("Process Completed")

    def closeEvent(self, event):
        if self.worker:
            self.worker.stop()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GSTApp()
    window.show()
    sys.exit(app.exec())