import re
with open("Gmail-Tools/main.py", "r", encoding="utf-8") as f:
    text = f.read()

new_settings_method = """
    def _open_email_settings(self):
        import os
        import json
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

        top = ctk.CTkToplevel(self)
        top.title("⚙ Email Configuration")
        top.geometry("440x280")
        top.transient(self)
        top.grab_set()

        title_fr = ctk.CTkFrame(top, fg_color="transparent")
        title_fr.pack(fill="x", pady=(20, 10))
        ctk.CTkLabel(title_fr, text="⚙", font=("Segoe UI Emoji", 16), text_color="#4da8da").pack(side="left", padx=(30, 10))
        ctk.CTkLabel(title_fr, text="Gmail Details", font=("Segoe UI", 14, "bold")).pack(side="left")

        frame = ctk.CTkFrame(top, corner_radius=8, border_width=1)
        frame.pack(padx=25, fill="x", pady=10)

        ctk.CTkLabel(frame, text="Gmail Address:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5, padx=10)
        e_email = ctk.CTkEntry(frame, width=200)
        e_email.grid(row=0, column=1, sticky="w", pady=5, padx=10)
        e_email.insert(0, email_val)

        ctk.CTkLabel(frame, text="App Password:", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", pady=(15, 5), padx=10)
        e_pass = ctk.CTkEntry(frame, width=200, show="*")
        e_pass.grid(row=1, column=1, sticky="w", pady=(15, 5), padx=10)
        e_pass.insert(0, pass_val)

        def save():
            with open(cred_path, "w") as f:
                json.dump({"email": e_email.get().strip(), "password": e_pass.get().strip()}, f)
            top.destroy()
            from tkinter import messagebox
            messagebox.showinfo("Saved", "Settings saved successfully!", parent=self)

        ctk.CTkButton(top, text="Save Settings", command=save).pack(pady=(10, 0))

    def _rebuild_config_tab(self):"""

text = text.replace("    def _rebuild_config_tab(self):", new_settings_method)

settings_btn = """
        cb = ctk.CTkComboBox(cb_frame, variable=self._tmpl_var, values=names, state="readonly", width=300, command=self._on_template_change)
        cb.pack(side="left")
        self._tmpl_cb = cb

        settings_btn = ctk.CTkButton(cb_frame, text="⚙ Email Settings", command=self._open_email_settings)
        settings_btn.pack(side="right", padx=(0, 0) if self._hide_switcher else (14, 0))
        
        import webbrowser
        link = ctk.CTkLabel(cb_frame, text="How to setup email 🔗", font=("Segoe UI", 10, "underline"), cursor="hand2", text_color="#4da8da")
        link.pack(side="right", padx=(14, 14))
        link.bind("<Button-1>", lambda e: webbrowser.open("https://www.youtube.com/watch?v=MkLX85XU5rU"))
"""
text = re.sub(r'cb = ctk\.CTkComboBox\(cb_frame.+?self\._tmpl_cb = cb', settings_btn, text, flags=re.DOTALL)

text = text.replace('self.title("Bulk Outlook Mailer")', 'self.title("Bulk Gmail Mailer")')

with open("Gmail-Tools/main.py", "w", encoding="utf-8") as f:
    f.write(text)
