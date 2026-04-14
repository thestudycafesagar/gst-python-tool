import os
import re

def fix_ui(path):
    if not os.path.exists(path): return
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # 1. Fix _download_sample parent argument
    content = content.replace("parent=self.winfo_toplevel()", "parent=self")
    
    # 2. Fix corrupted unicode in the buttons
    content = content.replace("ï¿½  Download Sample", "📥  Download Sample")
    content = content.replace("ï¿½ðŸ—‘  Clear All", "🗑️  Clear All")
    
    # 3. Pull Settings configuration out of hide_switcher
    # Find the current _rebuild_config_tab injection
    search = """        self._row_i = 0
        if not self._hide_switcher:
            lbl = tk.Label(form_card, text="Email Template", anchor="w", bg=CARD, fg=TEXT, font=("Segoe UI", 11, "bold"))
            lbl.grid(row=self._row_i, column=0, sticky="w", pady=12, padx=(0, 28))

            cb_frame = tk.Frame(form_card, bg=CARD)
            cb_frame.grid(row=self._row_i, column=1, sticky="w", pady=12)

            self._tmpl_var = tk.StringVar()
            self._tmpl_var.set(f"{t.icon}  {t.name}")
            names = [f"{tmpl.icon}  {tmpl.name}" for tmpl in ALL_TEMPLATES]
            import tkinter.ttk as ttk
            cb = ttk.Combobox(cb_frame, textvariable=self._tmpl_var, values=names, state="readonly", width=36, style="TCombobox")
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: self._on_template_change())
            self._tmpl_cb = cb

            if hasattr(self, "_open_email_settings"):
                settings_btn = ttk.Button(cb_frame, text="⚙ Email Settings", style="Import.TButton", command=self._open_email_settings)
                settings_btn.pack(side="right", padx=(14, 0))
                import webbrowser
                link = tk.Label(cb_frame, text="How to setup email 🔗", bg=CARD, fg="#4da8da", font=("Segoe UI", 10, "underline"), cursor="hand2")
                link.pack(side="right", padx=(14, 14))
                link.bind("<Button-1>", lambda e: webbrowser.open("https://www.youtube.com/watch?v=MkLX85XU5rU"))

            self._row_i += 1
            sep = tk.Frame(form_card, bg=BORDER, height=1)
            sep.grid(row=self._row_i, column=0, columnspan=2, sticky="ew", pady=(0, 16))
            self._row_i += 1"""
    
    replace = """        self._row_i = 0
        # Settings Bar always visible!
        lbl_text = "Email Template" if not self._hide_switcher else f"{t.icon}  {t.name} Config"
        lbl = tk.Label(form_card, text=lbl_text, anchor="w", bg=CARD, fg=TEXT, font=("Segoe UI", 13, "bold"))
        lbl.grid(row=self._row_i, column=0, sticky="w", pady=12, padx=(0, 28))

        cb_frame = tk.Frame(form_card, bg=CARD)
        cb_frame.grid(row=self._row_i, column=1, sticky="e" if self._hide_switcher else "w", pady=12)

        if not self._hide_switcher:
            self._tmpl_var = tk.StringVar()
            self._tmpl_var.set(f"{t.icon}  {t.name}")
            names = [f"{tmpl.icon}  {tmpl.name}" for tmpl in ALL_TEMPLATES]
            import tkinter.ttk as ttk
            cb = ttk.Combobox(cb_frame, textvariable=self._tmpl_var, values=names, state="readonly", width=36, style="TCombobox")
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: self._on_template_change())
            self._tmpl_cb = cb

        if hasattr(self, "_open_email_settings"):
            import tkinter.ttk as ttk
            settings_btn = ttk.Button(cb_frame, text="⚙ Email Settings", style="Import.TButton", command=self._open_email_settings)
            settings_btn.pack(side="right", padx=(0, 0) if self._hide_switcher else (14, 0))
            import webbrowser
            link = tk.Label(cb_frame, text="How to setup email 🔗", bg=CARD, fg="#4da8da", font=("Segoe UI", 10, "underline"), cursor="hand2")
            link.pack(side="right", padx=(14, 14))
            link.bind("<Button-1>", lambda e: webbrowser.open("https://www.youtube.com/watch?v=MkLX85XU5rU"))

        self._row_i += 1
        sep = tk.Frame(form_card, bg=BORDER, height=1)
        sep.grid(row=self._row_i, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        self._row_i += 1"""
        
    content = content.replace(search, replace)
    
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

fix_ui("Gmail-Tools/main.py")
fix_ui("Outlook Email Tools/main.py")
print("Done patching UI fields and fixing download buttons")
