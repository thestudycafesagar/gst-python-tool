import os
import re

file_path = r"GST\GST 2B Downloader\main.py"
with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

pattern = re.compile(
    r'(# .*Search box.*?search_entry\.pack\(fill="x", padx=16, pady=\(0, 6\)\)).*?(foot = ctk\.CTkFrame.*?def _load\(\):.*?self\.manual_credentials = selected)',
    re.DOTALL
)

match = pattern.search(content)
if match:
    print("MATCH FOUND!")
    print("Group 1 ends with:", match.group(1)[-50:])
    print("Group 2 starts with:", match.group(2)[:50])
else:
    print("NO MATCH")
