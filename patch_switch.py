import os

def fix_switch(path):
    if not os.path.exists(path): return
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    old = """    def _switch_template(self, idx):
        self._active_template = ALL_TEMPLATES[idx]
        self._tmpl_var.set(f"{self._active_template.icon}  {self._active_template.name}")
        self._hdr_title.config(
            text=f"{self._active_template.icon}  {self._active_template.name}")
        self._attachment_folder = ""
        self._rebuild_config_tab()
        self._rebuild_recipients_tab()
        self._status_lbl.config(text="")"""
        
    old_backup = """    def _switch_template(self, idx):
        self._active_template = ALL_TEMPLATES[idx]
        if hasattr(self, "_tmpl_var"): self._tmpl_var.set(f"{self._active_template.icon}  {self._active_template.name}")
        self._hdr_title.config(
            text=f"{self._active_template.icon}  {self._active_template.name}")
        self._attachment_folder = ""
        self._rebuild_config_tab()
        self._rebuild_recipients_tab()
        self._status_lbl.config(text="")"""        

    new = """    def _switch_template(self, idx):
        self._active_template = ALL_TEMPLATES[idx]
        self._attachment_folder = ""
        self._rebuild_config_tab()
        self._rebuild_recipients_tab()
        if hasattr(self, "_status_lbl") and hasattr(self._status_lbl, "config"):
            self._status_lbl.config(text="")"""
            
    content = content.replace(old, new)
    content = content.replace(old_backup, new)

    # Some versions might have slight spacing differently, use regex if needed
    import re
    content = re.sub(
        r'def _switch_template\(self, idx\):\n\s+self\._active_template = ALL_TEMPLATES\[idx\].*?self\._rebuild_config_tab\(\)',
        r'def _switch_template(self, idx):\n        self._active_template = ALL_TEMPLATES[idx]\n        self._attachment_folder = ""\n        self._rebuild_config_tab()',
        content, flags=re.DOTALL
    )

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

fix_switch("Gmail-Tools/main.py")
fix_switch("Outlook Email Tools/main.py")
