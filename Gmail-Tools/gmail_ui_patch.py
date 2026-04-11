import os
import re

TARGET = r"Gmail-Tools\main.py"

with open(TARGET, "r", encoding="utf-8") as f:
    code = f.read()

# 1. Remove app_password and sender_email from config_fields
code = re.sub(r'\(\"Sender Gmail Address:\",\s*\"sender_email\"\),\s*', '', code)
code = re.sub(r'\(\"App Password:\",\s*\"app_password\"\),\s*', '', code)
code = re.sub(r'\"sender_email\":\s*\"your_email@gmail\.com\",\s*', '', code)
code = re.sub(r'\"app_password\":\s*\"\",\s*', '', code)

# 2. Add import webbrowser if not present
if "import webbrowser" not in code:
    code = code.replace("import tkinter as tk", "import tkinter as tk\nimport webbrowser\nimport json\nimport os")

# 3. Add header buttons
header_search = r'''        self._tmpl_cb  = None

        if not self._hide_switcher:'''

header_replace = r'''        self._tmpl_cb  = None

        # Add Email Settings & Setup link to header
        link_lbl = tk.Label(hdr, text="How to setup email 🔗", bg=HDR_BG, fg="#4da8da", font=("Segoe UI", 10, "underline"), cursor="hand2")
        link_lbl.pack(side="right", padx=(0, 14), pady=(12, 0), anchor="n")
        link_lbl.bind("<Button-1>", lambda e: webbrowser.open("https://www.youtube.com/watch?v=1YRVZcqxxNQ"))

        settings_btn = ttk.Button(hdr, text="⚙ Email Settings", style="Import.TButton", command=self._open_email_settings)
        settings_btn.pack(side="right", padx=(0, 14), pady=(12, 0), anchor="n")

        if not self._hide_switcher:'''

code = code.replace(header_search, header_replace)

# 4. Add _open_email_settings
settings_method = r'''
    def _open_email_settings(self):
        cred_path = os.path.join(os.path.dirname(__file__), "gmail_credentials.json")
        email_val = ""
        pass_val = ""
        if os.path.exists(cred_path):
            try:
                with open(cred_path, "r") as f:
                    data = json.load(f)
                    email_val = data.get("email", "")
                    pass_val = data.get("password", "")
            except:
                pass
            
        top = tk.Toplevel(self)
        top.title("⚙ Email Configuration")
        top.geometry("400x250")
        top.configure(bg=BG)
        top.transient(self)
        top.grab_set()
        
        tk.Label(top, text="Gmail Details", font=("Segoe UI", 12, "bold"), bg=BG, fg=TEXT).pack(pady=(15, 5))
        
        frame = tk.Frame(top, bg=BG)
        frame.pack(padx=20, pady=10, fill="x")
        
        tk.Label(frame, text="Gmail Address:", bg=BG, fg=TEXT, font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=5)
        e_email = ttk.Entry(frame, width=30)
        e_email.grid(row=0, column=1, sticky="w", pady=5, padx=5)
        e_email.insert(0, email_val)
        
        tk.Label(frame, text="App Password:", bg=BG, fg=TEXT, font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=5)
        e_pass = ttk.Entry(frame, width=30, show="*")
        e_pass.grid(row=1, column=1, sticky="w", pady=5, padx=5)
        e_pass.insert(0, pass_val)
        
        def save():
            with open(cred_path, "w") as f:
                json.dump({"email": e_email.get().strip(), "password": e_pass.get().strip()}, f)
            top.destroy()
            messagebox.showinfo("Saved", "Settings saved successfully!", parent=self)
            
        ttk.Button(top, text="Save Settings", style="Send.TButton", command=save).pack(pady=(10, 0))

    # ── Notebook ────────────────────────────────────────────────'''

notebook_search = r'''    # ── Notebook ────────────────────────────────────────────────'''

code = code.replace(notebook_search, settings_method)

# 5. Fix send_emails

send_search = r'''def send_emails(template, cfg, recipients, attachment_folder, log_cb, done_cb):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    import os

    sender_email = str(cfg.get("sender_email", "")).strip()
    app_password = str(cfg.get("app_password", "")).strip()

    if not sender_email or not app_password:'''

send_replace = r'''def send_emails(template, cfg, recipients, attachment_folder, log_cb, done_cb):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    import os
    import json

    cred_path = os.path.join(os.path.dirname(__file__), "gmail_credentials.json")
    sender_email = ""
    app_password = ""
    try:
        if os.path.exists(cred_path):
            with open(cred_path, "r") as f:
                creds = json.load(f)
                sender_email = creds.get("email", "")
                app_password = creds.get("password", "")
    except Exception:
        pass

    if not sender_email or not app_password:'''
    
code = code.replace(send_search, send_replace)

# 6. Also fix show_preview missing app password error check if applicable
preview_search = r'''        sender_email = str(self._cfg_vars.get("sender_email", tk.StringVar()).get())
        if not sender_email:
            sender_email = "Your Email"'''
preview_replace = r'''        cred_path = os.path.join(os.path.dirname(__file__), "gmail_credentials.json")
        sender_email = "Your Email"
        try:
            if os.path.exists(cred_path):
                with open(cred_path, "r") as f:
                    sender_email = json.load(f).get("email", "Your Email")
        except:
            pass'''
code = code.replace(preview_search, preview_replace)


with open(TARGET, "w", encoding="utf-8") as f:
    f.write(code)

print("PATCHED")
