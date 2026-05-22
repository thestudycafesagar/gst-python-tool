import os
import re

files = [
    r"GST\GST 2B Downloader\main.py",
    r"GST\GST 3B Downloader\main.py",
    r"GST\GST Challan Downloader\main.py",
    r"GST\GST R1 Downloader\mai.py",
    r"GST\R1 PDF Downloader\main.py",
    r"GST\IMS Downloader\main.py"
]

target = '''            if n > 0 and hasattr(self, "period_mode_var"):
                self.period_mode_var.set(selected[0].get("FilingFrequency", "Monthly"))
                if hasattr(self, "toggle_inputs"):
                    self.toggle_inputs()'''

replacement = '''            if n > 0 and hasattr(self, "period_mode_var"):
                self.period_mode_var.set(selected[0].get("FilingFrequency", "Monthly"))
                if hasattr(self, "toggle_inputs"):
                    self.toggle_inputs()
                if hasattr(self, "mode_tabs"):
                    self.mode_tabs.configure(state="disabled")'''

for f in files:
    if not os.path.exists(f): continue
    with open(f, "r", encoding="utf-8") as file:
        content = file.read()
    
    if target in content:
        c1 = content.replace(target, replacement)
        with open(f, "w", encoding="utf-8") as file:
            file.write(c1)
        print(f"Patched {f}")
    else:
        print(f"Target not found in {f}")

