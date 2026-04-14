import os

def fix_ui(filepath):
    if not os.path.exists(filepath):
        return
    with open(filepath, "r", encoding="utf-8") as f:
        text = f.read()

    # Fix the extra spaces vertically for the tab container
    text = text.replace('form_card.pack(fill="both", expand=True, padx=30, pady=20)', 'form_card.pack(fill="x", anchor="n", padx=30, pady=20)')

    # Also fix the `hide_switcher` indentation bug in Gmail-Tools
    bad_cb = '''        if not self._hide_switcher:
            self._tmpl_var = ctk.StringVar(value=f"{t.icon}  {t.name}")
            names = [f"{tmpl.icon}  {tmpl.name}" for tmpl in ALL_TEMPLATES]

        cb = ctk.CTkComboBox(cb_frame, variable=self._tmpl_var, values=names, state="readonly", width=300, command=self._on_template_change)
        cb.pack(side="left")
        self._tmpl_cb = cb'''

    good_cb = '''        if not self._hide_switcher:
            self._tmpl_var = ctk.StringVar(value=f"{t.icon}  {t.name}")
            names = [f"{tmpl.icon}  {tmpl.name}" for tmpl in ALL_TEMPLATES]
            cb = ctk.CTkComboBox(cb_frame, variable=self._tmpl_var, values=names, state="readonly", width=300, command=self._on_template_change)
            cb.pack(side="left")
            self._tmpl_cb = cb'''

    text = text.replace(bad_cb, good_cb)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(text)

fix_ui('Outlook Email Tools/main.py')
fix_ui('Gmail-Tools/main.py')
print("Fixed!")