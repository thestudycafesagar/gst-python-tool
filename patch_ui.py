import os
import re

def patch_file(path):
    if not os.path.exists(path): return
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    # 1. Remove _build_header completely
    start = content.find("def _build_header(self):")
    if start != -1:
        if "def _open_email_settings(self):" in content[start:]:
            end = content.find("def _open_email_settings(self):", start)
        else:
            end = content.find("def _build_notebook(self):", start)
            
        if end != -1:
            content = content[:start] + "def _build_header(self):\n        pass\n\n    " + content[end:]

    # 2. Patch _rebuild_config_tab
    search = """        t = self._active_template

        outer = tk.Frame(self._tab_config, bg=BG)
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 12))

        # Form card
        form_border = tk.Frame(outer, bg=BORDER, padx=1, pady=1)
        form_border.pack(fill="x")
        form_card = tk.Frame(form_border, bg=CARD, padx=32, pady=24)
        form_card.pack(fill="both", expand=True)
        form_card.columnconfigure(1, weight=1)

        def validate_phone(P):"""
        
    replace = """        t = self._active_template

        outer = tk.Frame(self._tab_config, bg=BG)
        outer.pack(fill="both", expand=True, padx=24, pady=(12, 12))

        # Form card
        form_border = tk.Frame(outer, bg=BORDER, padx=1, pady=1)
        form_border.pack(fill="x")
        form_card = tk.Frame(form_border, bg=CARD, padx=32, pady=24)
        form_card.pack(fill="both", expand=True)
        form_card.columnconfigure(1, weight=1)

        # Template Selection inside card instead of header
        self._row_i = 0
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
            self._row_i += 1

        def validate_phone(P):"""
        
    content = content.replace(search, replace)
    
    # 3. Patch grid coordinates for dynamic row placement
    content = content.replace('.grid(row=i, column=0, sticky="w",', '.grid(row=getattr(self, "_row_i", 0)+i, column=0, sticky="w",')
    content = content.replace('ef.grid(row=i, column=1, sticky="ew", pady=12)', 'ef.grid(row=getattr(self, "_row_i", 0)+i, column=1, sticky="ew", pady=12)')

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

patch_file("Gmail-Tools/main.py")
patch_file("Outlook Email Tools/main.py")
print("Done patching UI")
