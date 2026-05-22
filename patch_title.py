import os

f = "GST_Suite.py"
with open(f, "r", encoding="utf-8") as file:
    content = file.read()

content = content.replace('path = fd.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls")])', 'path = fd.askopenfilename(title="Import Excel File", filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls")])')

with open(f, "w", encoding="utf-8") as file:
    file.write(content)

print("Patched title!")
