import os

f = "GST_Suite.py"
with open(f, "r", encoding="utf-8") as file:
    content = file.read()

content = content.replace('"Username (GSTIN)"', '"Username"')
content = content.replace('"Username (PAN)"', '"Username"')

with open(f, "w", encoding="utf-8") as file:
    file.write(content)

print("Patched column names")
