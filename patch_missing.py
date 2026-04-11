import re

print('Patching IMS...')
with open(r'GST\IMS Downloader\main.py', 'r', encoding='utf-8') as f:
    src = f.read()

# Instead of relying strictly on parenthesis, let's just find the entire block and inject.
old_block = ''').pack(fill="x", padx=15, pady=(0, 15))

        # LOG BOX'''
new_block = ''').pack(fill="x", padx=15, pady=(0, 5))
        self.btn_download = ctk.CTkButton(card, text="📥 Download Sample Excel", command=self.download_sample, fg_color="#43a047", hover_color="#2e7d32", height=35)
        self.btn_download.pack(fill="x", padx=15, pady=(5, 15))

        # LOG BOX'''

if old_block in src:
    with open(r'GST\IMS Downloader\main.py', 'w', encoding='utf-8') as f:
        f.write(src.replace(old_block, new_block))
    print('Patched IMS')
else:
    print('IMS already patched or block not found')

print('Patching 26 AS...')
with open(r'Income Tax\26 AS Downlaoder\main.py', 'r', encoding='utf-8') as f:
    src26 = f.read()

old_26as = 'ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("26as"), width=100).pack(side="right")'
new_26as = 'ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n            ' + old_26as
if old_26as in src26:
    src26 = src26.replace(old_26as, new_26as)

old_ais = 'ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("ais"), width=100).pack(side="right")'
new_ais = 'ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n            ' + old_ais
if old_ais in src26:
    src26 = src26.replace(old_ais, new_ais)

old_tis = 'ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("tis"), width=100).pack(side="right")'
new_tis = 'ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n            ' + old_tis
if old_tis in src26:
    src26 = src26.replace(old_tis, new_tis)

with open(r'Income Tax\26 AS Downlaoder\main.py', 'w', encoding='utf-8') as f:
    f.write(src26)
print('Patched 26 AS')

print('Patching IT Challan...')
with open(r'Income Tax\Challan Downloader\main.py', 'r', encoding='utf-8') as f:
    srcIT = f.read()

old_it = 'ctk.CTkButton(f_frame, text="BROWSE", command=self.browse_file, width=100).pack(side="right")'
new_it = 'ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047\", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n            ' + old_it
if old_it in srcIT:
    srcIT = srcIT.replace(old_it, new_it)
    with open(r'Income Tax\Challan Downloader\main.py', 'w', encoding='utf-8') as f:
        f.write(srcIT)
    print('Patched IT Challan')
else:
    print('IT Challan already patched')

