import os

f = r"GST\R1 PDF Downloader\main.py"

with open(f, "r", encoding="utf-8") as file:
    content = file.read()

content = content.replace("view_btn.click()", 'self.driver.execute_script("arguments[0].click();", view_btn)')
content = content.replace("summary_btn.click()", 'self.driver.execute_script("arguments[0].click();", summary_btn)')
content = content.replace("pdf_btn.click()", 'self.driver.execute_script("arguments[0].click();", pdf_btn)')

with open(f, "w", encoding="utf-8") as file:
    file.write(content)

print("Patched R1 PDF clicks.")
