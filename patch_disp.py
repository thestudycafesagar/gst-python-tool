import os

files = [
    r"GST\GST 2B Downloader\main.py",
    r"GST\GST 3B Downloader\main.py",
    r"GST\GST Challan Downloader\main.py",
    r"GST\GST R1 Downloader\mai.py",
    r"GST\R1 PDF Downloader\main.py",
    r"GST\IMS Downloader\main.py"
]

target = 'disp = f"{c} ({u})" if c else u'
replacement = 'disp = f"{c} ({u}) [{f_freq}]" if c else f"{u} [{f_freq}]"'

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
