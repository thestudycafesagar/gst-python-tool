import os
import re

f = "GST_Suite.py"
with open(f, "r", encoding="utf-8") as file:
    content = file.read()

pattern = re.compile(r'(try:\s*df\.to_excel\(path, index=False\))(.*?)(mb\.showinfo\("Success")', re.DOTALL)

replacement = r'''try:
                    df.to_excel(path, index=False)
                    if not is_it:
                        try:
                            from openpyxl import load_workbook
                            from openpyxl.worksheet.datavalidation import DataValidation
                            wb = load_workbook(path)
                            ws = wb.active
                            dv = DataValidation(type="list", formula1='"Monthly,Quarterly"', allow_blank=True)
                            ws.add_data_validation(dv)
                            dv.add("D2:D1048576")
                            wb.save(path)
                        except Exception as e:
                            pass
                    \3'''

if pattern.search(content):
    c1 = pattern.sub(replacement, content)
    with open(f, "w", encoding="utf-8") as file:
        file.write(c1)
    print("Patched openpyxl validation!")
else:
    print("Pattern not found!")
