import os

files = [
    r"GST\GST R1 Downloader\mai.py",
    r"GST\R1 PDF Downloader\main.py"
]

target = 'self.log(f"   dY". Bulk Mode: processing {len(tasks)} periods")'
target2 = 'self.log(f"   dY". Mode: Quarterly ({selected_q} -> {selected_m})")'
target3 = 'self.log(f"   dY". Mode: Monthly ({selected_m})")'

for f in files:
    if not os.path.exists(f): continue
    with open(f, "r", encoding="utf-8") as file:
        content = file.read()
    
    c1 = content.replace(target, 'self.log(f"   dY Bulk Mode: processing {len(tasks)} periods")')
    c1 = c1.replace(target2, 'self.log(f"   dY Mode: Quarterly ({selected_q} -> {selected_m})")')
    c1 = c1.replace(target3, 'self.log(f"   dY Mode: Monthly ({selected_m})")')
    
    with open(f, "w", encoding="utf-8") as file:
        file.write(c1)
    print(f"Patched {f}")

