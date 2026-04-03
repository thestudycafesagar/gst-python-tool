import os
import re

with open('main.py', 'r', encoding='utf-8') as f:
    text = f.read()

text = re.sub(
    r'dst_dir = self\.dst_dir\.get\(\)\.strip\(\)\s*if not src or not os\.path\.isfile\(src\):\s*messagebox\.showwarning\("No Source", "Please select a valid PDF file\."\)\s*return\s*if not dst_dir:\s*messagebox\.showwarning\("No Destination", "Please select an output folder\."\)\s*return\s*base_name = os\.path\.splitext\(os\.path\.basename\(src\)\)\[0\]\s*out_path\s*=\s*os\.path\.join\(dst_dir, f"\{base_name\}_converted\.xlsx"\)',
    '''if not src or not os.path.isfile(src):
            messagebox.showwarning("No Source", "Please select a valid PDF file.")
            return

        import time
        dst_dir = os.path.dirname(src)
        base_name = os.path.splitext(os.path.basename(src))[0]
        out_path  = os.path.join(dst_dir, f"{base_name}_converted_{int(time.time())}.xlsx")''',
    text, flags=re.DOTALL
)

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(text)
