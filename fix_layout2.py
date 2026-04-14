import os

def fix_ui(filepath):
    if not os.path.exists(filepath):
        return
    with open(filepath, "r", encoding="utf-8") as f:
        text = f.read()

    # Find the bad block in Gmail-Tools/main.py
    import re
    
    # We want to change the unindented 'cb =' block to be indented inside the if statement
    pattern = r'(if not self\._hide_switcher:\n\s+self\._tmpl_var = ctk\.StringVar\(value=f"\{t\.icon\}  \{t\.name\}"\)\n\s+names = \[f"\{tmpl\.icon\}  \{tmpl\.name\}" for tmpl in ALL_TEMPLATES\])\n\n\s+cb = ctk\.CTkComboBox\(cb_frame, variable=self\._tmpl_var, values=names, state="readonly", width=300, command=self\._on_template_change\)\n\s+cb\.pack\(side="left"\)\n\s+self\._tmpl_cb = cb'

    def repl(m):
        return m.group(1) + '\n            cb = ctk.CTkComboBox(cb_frame, variable=self._tmpl_var, values=names, state="readonly", width=300, command=self._on_template_change)\n            cb.pack(side="left")\n            self._tmpl_cb = cb'

    text, count = re.subn(pattern, repl, text)
    if count > 0:
        print(f"Fixed combobox indentation in {filepath}")

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(text)

fix_ui('Gmail-Tools/main.py')
