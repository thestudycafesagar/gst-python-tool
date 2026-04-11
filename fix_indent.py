def fix(path):
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    with open(path, 'w', encoding='utf-8') as f:
        for line in lines:
            if 'DOWNLOAD SAMPLE' in line:
                continue
            if 'ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("26as")' in line:
                f.write('        ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n')
            if 'ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("ais")' in line:
                f.write('        ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n')
            if 'ctk.CTkButton(f_frame, text="BROWSE", command=lambda: self.browse_file("tis")' in line:
                f.write('        ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n')
            if 'ctk.CTkButton(f_frame, text="BROWSE", command=self.browse_file' in line:
                f.write('        ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n')
                
            f.write(line)

fix(r'Income Tax\26 AS Downlaoder\main.py')
fix(r'Income Tax\Challan Downloader\main.py')

