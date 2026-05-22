import os
import re

files = [
    r"GST\GST R1 Downloader\mai.py",
    r"GST\R1 PDF Downloader\main.py"
]

pattern = re.compile(
    r'(# 2\. DEFINE TASKS.*?)(selected_q = self\.settings\[\'quarter\'\].*?)(# 3\. EXECUTE LOOP)',
    re.DOTALL
)

replacement = r'''\1
            if "tasks" in self.settings:
                tasks = self.settings["tasks"]
                self.log(f"   dY". Bulk Mode: processing {len(tasks)} periods")
            else:
                selected_q = self.settings.get('quarter', '')
                period_mode = self.settings.get('period_mode', 'Monthly')
                if selected_q not in q_map:
                    return "Config Error", "Invalid Month/Quarter Selection"
    
                if period_mode == "Quarterly":
                    selected_m = q_map[selected_q][-1]
                    tasks = [{"q": selected_q, "m": selected_m}]
                    self.log(f"   dY". Mode: Quarterly ({selected_q} -> {selected_m})")
                else:
                    selected_m = self.settings.get('month', '')
                    if selected_m not in q_map[selected_q]:
                        return "Config Error", "Invalid Month/Quarter Selection"
                    tasks = [{"q": selected_q, "m": selected_m}]
                    self.log(f"   dY". Mode: Monthly ({selected_m})")

            \3'''

for f in files:
    if not os.path.exists(f): continue
    with open(f, "r", encoding="utf-8") as file:
        content = file.read()
    
    match = pattern.search(content)
    if match:
        c1 = pattern.sub(replacement, content)
        with open(f, "w", encoding="utf-8") as file:
            file.write(c1)
        print(f"Patched {f}")
    else:
        print(f"Target not found in {f}")
