import threading
import time
import os
import sys
import random
import glob
import base64
import pandas as pd
import customtkinter as ctk
from PIL import Image
from datetime import datetime
from tkinter import filedialog, messagebox
import sqlite3

# Selenium Imports  
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select

# Shared Stealth Driver Import
import sys
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _ROOT not in sys.path: sys.path.insert(0, _ROOT)
from stealth_driver import create_chrome_driver, build_chrome_options, show_browser_alert

# --- UI CONFIGURATION ---
ctk.set_appearance_mode("System")
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
        self.report_data = [] 

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
        self.log("🚀 INITIALIZING GSTR-1 JSON ENGINE...")
        try:
            if self.credentials:
                df = pd.DataFrame(self.credentials)
                user_col, pass_col = "Username", "Password"
            else:
                if not self.excel_path:
                    self.app.process_finished_safe("No credentials.")
                    return
                df = pd.read_excel(self.excel_path)
                clean_cols = {c.lower().strip(): c for c in df.columns}
                user_col = next((clean_cols[c] for c in clean_cols if 'user' in c), "Username")
                pass_col = next((clean_cols[c] for c in clean_cols if 'pass' in c), "Password")

            total = len(df)
            base_dir = os.path.join(os.getcwd(), "GST Downloaded", "GSTR1 JSON")
            os.makedirs(base_dir, exist_ok=True)

            for index, row in df.iterrows():
                if not self.keep_running: break
                username = str(row[user_col]).strip()
                password = str(row[pass_col]).strip()
                self.app.update_progress_safe(index / total)
                self.log(f"\n🔹 Processing: {username}")
                
                user_root = os.path.join(base_dir, username)
                os.makedirs(user_root, exist_ok=True)
                
                status, reason = self.process_single_user(username, password, user_root)
                self.report_data.append({"Username": username, "Status": status, "Details": reason})
                if not self.keep_running: break

            self.app.update_progress_safe(1.0)
            self.log("✅ TASKS COMPLETED.")
            self.app.process_finished_safe("Process Finished")
        except Exception as e:
            self.log(f"❌ Error: {e}")
            self.app.process_finished_safe("Error")

    def process_single_user(self, username, password, user_root):
        try:
            self.driver = create_chrome_driver(build_chrome_options(user_root))
            wait = WebDriverWait(self.driver, 20)
            
            self.driver.get("https://services.gst.gov.in/services/login")
            try:
                usr = wait.until(EC.presence_of_element_located((By.ID, "username")))
                self.type_like_human(usr, username)
                pwd = self.driver.find_element(By.ID, "user_pass")
                self.type_like_human(pwd, password)
            except: pass
            
            # If captcha present, show a browser banner
            try:
                if self.driver.find_elements(By.ID, "imgCaptcha"):
                    show_browser_alert(self.driver, "Please enter captcha in the browser and submit to continue.")
                    self.log("   🟨 Captcha detected — please complete it in the browser.")
            except Exception:
                pass

            self.log("   👉 Complete login and click Return Dashboard.")
            while self.keep_running:
                url = self.driver.current_url.lower()
                if any(k in url for k in ("dashboard", "auth/home")):
                    if not self.driver.find_elements(By.ID, "imgCaptcha"): break
                time.sleep(2)
            
            if not self.keep_running: return "Stopped", "User stopped"
            time.sleep(2)
            
            try:
                dash_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Return Dashboard')]")))
                self.driver.execute_script("arguments[0].click();", dash_btn)
            except: pass
            
            time.sleep(3)
            fin_year = self.app.cb_year.get()
            q_text = self.app.cb_qtr.get()
            m_text = self.app.cb_month.get()
            
            Select(wait.until(EC.presence_of_element_located((By.NAME, "fin")))).select_by_visible_text(fin_year)
            time.sleep(2)
            Select(self.driver.find_element(By.NAME, "quarter")).select_by_visible_text(q_text)
            time.sleep(2)
            Select(self.driver.find_element(By.NAME, "mon")).select_by_visible_text(m_text)
            time.sleep(1)
            self.driver.find_element(By.XPATH, "//button[contains(text(), 'Search')]").click()
            time.sleep(4)
            
            # GSTR-1 JSON Logic
            view_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//p[contains(text(),'GSTR1')]/ancestor::div[contains(@class,'col-')]//button[contains(normalize-space(),'PREPARE OFFLINE')]")))
            self.driver.execute_script("arguments[0].click();", view_btn)
            time.sleep(4)
            
            # Download Tab
            dl_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Download')]")))
            self.driver.execute_script("arguments[0].click();", dl_tab)
            time.sleep(3)
            
            # Generate JSON
            gen_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'GENERATE JSON FILE')]")))
            self.driver.execute_script("arguments[0].click();", gen_btn)
            time.sleep(10)
            
            return "Success", "Processed"
        except Exception as e:
            return "Error", str(e)[:50]
        finally:
            if self.driver: self.driver.quit()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GSTR-1 JSON Downloader")
        self.geometry("800x800")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.manual_credentials = []
        
        # Header
        self.head = ctk.CTkFrame(self, fg_color="#1D4ED8", height=70)
        self.head.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(self.head, text="GSTR-1 JSON DOWNLOADER", font=("Segoe UI", 22, "bold"), text_color="white").pack(side="left", padx=20)

        # Content
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll.grid_columnconfigure(0, weight=1)
        
        card = ctk.CTkFrame(self.scroll, border_color="#334155", border_width=1)
        card.pack(fill="x", padx=10, pady=10)
        self.ent_file = ctk.CTkEntry(card, placeholder_text="Credentials...")
        self.ent_file.pack(fill="x", padx=15, pady=10)
        ctk.CTkButton(card, text="📂 Load ID Pass", command=self.load_id_pass, fg_color="#4338ca").pack(fill="x", padx=15, pady=5)

        # Settings
        card2 = ctk.CTkFrame(self.scroll)
        card2.pack(fill="x", padx=10, pady=10)
        self.cb_year = ctk.CTkComboBox(card2, values=["2023-24", "2024-25", "2025-26"])
        self.cb_year.pack(pady=5)
        self.cb_qtr = ctk.CTkComboBox(card2, values=["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)", "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"])
        self.cb_qtr.pack(pady=5)
        self.cb_month = ctk.CTkComboBox(card2, values=["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"])
        self.cb_month.pack(pady=5)

        self.log_box = ctk.CTkTextbox(self.scroll, height=300)
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

        self.prog_bar = ctk.CTkProgressBar(self)
        self.prog_bar.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        self.prog_bar.set(0)

        self.btn_start = ctk.CTkButton(self, text="START DOWNLOAD", height=50, fg_color="#059669", command=self.start_process)
        self.btn_start.grid(row=3, column=0, sticky="ew", padx=20, pady=20)

    def load_id_pass(self):
        import sqlite3 as _sq, os as _os
        db_path = _os.path.join(_os.environ.get("APPDATA", _os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
        if not _os.path.exists(db_path):
            db_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "..", "suite_profiles.db")
        try:
            conn = _sq.connect(db_path)
            rows = conn.execute("SELECT username, password FROM gst_profiles ORDER BY username").fetchall()
            conn.close()
        except Exception:
            rows = []
        if not rows:
            messagebox.showinfo("No Profiles", "No saved profiles found. Add via GST Suite -> Manage ID/Pass.", parent=self)
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Load ID Password")
        dialog.geometry("400x460")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        ctk.CTkLabel(dialog, text="Select Profiles to Load", font=("Segoe UI", 14, "bold")).pack(pady=(16, 8))
        sel_all_var = ctk.BooleanVar()
        vars_ = {}
        def _toggle_all():
            state = sel_all_var.get()
            for v in vars_.values():
                v.set(state)
        ctk.CTkCheckBox(dialog, text="Select All", variable=sel_all_var, command=_toggle_all,
                        font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
        scroll = ctk.CTkScrollableFrame(dialog, height=300)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        for u, p in rows:
            v = ctk.BooleanVar()
            ctk.CTkCheckBox(scroll, text=u, variable=v).pack(anchor="w", padx=10, pady=3)
            vars_[(u, p)] = v
        def _load():
            selected = [{"Username": u, "Password": p} for (u, p), v in vars_.items() if v.get()]
            if not selected:
                return
            self.manual_credentials = selected
            n = len(selected)
            label = selected[0]["Username"] if n == 1 else f"Loaded {n} profiles"
            self.ent_file.delete(0, "end")
            self.ent_file.insert(0, label)
            dialog.destroy()
        ctk.CTkButton(dialog, text="Load Selected", fg_color="#4338ca", command=_load).pack(pady=8)

    def add_id_password(self):
        dialog = ctk.CTkToplevel(self)
        u_ent = ctk.CTkEntry(dialog, placeholder_text="User")
        u_ent.pack()
        p_ent = ctk.CTkEntry(dialog, placeholder_text="Pass", show="*")
        p_ent.pack()
        def save():
            self.manual_credentials = [{"Username": u_ent.get(), "Password": p_ent.get()}]
            self.ent_file.delete(0, "end")
            self.ent_file.insert(0, f"Manual: {u_ent.get()}")
            dialog.destroy()
        ctk.CTkButton(dialog, text="Save", command=save).pack()

    def load_master_profiles(self):
        MasterProfilePicker(self, self.on_master_profiles_selected)

    def on_master_profiles_selected(self, profiles):
        if not profiles: return
        self.manual_credentials = profiles
        self.ent_file.delete(0, "end")
        self.ent_file.insert(0, f"Loaded {len(profiles)} profiles")

    def update_log_safe(self, msg):
        self.after(0, lambda: self.log_box.insert("end", f"{msg}\n"))

    def update_progress_safe(self, val):
        self.after(0, lambda: self.prog_bar.set(val))

    def process_finished_safe(self, msg):
        self.after(0, lambda: messagebox.showinfo("Info", msg))
        self.after(0, lambda: self.btn_start.configure(state="normal"))

    def start_process(self):
        self.btn_start.configure(state="disabled")
        self.worker = GSTWorker(self, None, {}, credentials=self.manual_credentials)
        threading.Thread(target=self.worker.run, daemon=True).start()

if __name__ == "__main__":
    app = App()
    app.mainloop()

class MasterProfilePicker(ctk.CTkToplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.title("Select Profiles")
        self.geometry("400x500")
        self.callback = callback
        self.vars = {}
        self.grab_set()
        self.attributes("-topmost", True)
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True)
        try:
            db_path = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
            if not os.path.exists(db_path):
                db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "suite_profiles.db")
            conn = sqlite3.connect(db_path)
            rows = conn.execute("SELECT username, password FROM gst_profiles ORDER BY username").fetchall()
            conn.close()
            for u, p in rows:
                v = ctk.BooleanVar()
                cb = ctk.CTkCheckBox(self.scroll, text=u, variable=v)
                cb.pack(padx=10, pady=2)
                self.vars[(u, p)] = v
        except: pass
        ctk.CTkButton(self, text="Load Selected", command=self._submit).pack(pady=10)

    def _submit(self):
        selected = [{"Username": u, "Password": p} for (u, p), v in self.vars.items() if v.get()]
        self.callback(selected)
        self.destroy()
