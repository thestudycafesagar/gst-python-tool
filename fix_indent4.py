path = r'Income Tax\26 AS Downlaoder\main.py'
with open(path, 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('            ctk.CTkButton(f_frame, text="BROWSE"', '        ctk.CTkButton(f_frame, text="BROWSE"')
with open(path, 'w', encoding='utf-8') as f:
    f.write(text)

path = r'Income Tax\Challan Downloader\main.py'
with open(path, 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('            ctk.CTkButton(f_frame, text="BROWSE"', '        ctk.CTkButton(f_frame, text="BROWSE"')
with open(path, 'w', encoding='utf-8') as f:
    f.write(text)

