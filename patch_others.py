import re
import os

files = {
    r'Income Tax\\26 AS Downlaoder\\refund_checker_app.py': 'Income Tax Sample File.xlsx',
    r'Income Tax\\Challan Downloader\\demand_checker_app.py': 'Income Tax Sample File.xlsx',
    r'Income Tax\\ITR - Bot\\GUI_based_app.py': 'Income Tax Sample File.xlsx'
}

for py_file, sample_name in files.items():
    if not os.path.exists(py_file):
        continue
    
    with open(py_file, 'r', encoding='utf-8') as f:
        src = f.read()
        
    if "def download_sample" in src:
        print(f"Skipping {py_file}, already modified.")
        continue

    method_injection = f'''
    def download_sample(self):
        import shutil
        import os
        from tkinter import messagebox, filedialog
        sample_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Challan Downloader", "{sample_name}") if "26 AS" in __file__ or "ITR" in __file__ else os.path.join(os.path.dirname(__file__), "{sample_name}")
        
        # Let's just point to Challan Downloader path 
        sample_path = r"Income Tax\\\\Challan Downloader\\\\{sample_name}"
        if not os.path.exists(sample_path):
            # Fallback
            sample_path = os.path.join(os.path.dirname(__file__), "{sample_name}")

        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="{sample_name}", filetypes=[("Excel", "*.xlsx")])
        if save_path:
            try:
                shutil.copy2(sample_path, save_path)
                messagebox.showinfo("Success", f"Sample downloaded to {{save_path}}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download: {{e}}")

    def browse_file(self):'''

    src = re.sub(r'\\+?\\s*def browse_file\\(self[^)]*\\):', method_injection, src)

    old_browse = 'ctk.CTkButton(f_frame, text="BROWSE", command=self.browse_file, width=100).pack(side="right")'
    new_browse = 'ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\\n        ' + old_browse
    
    if old_browse in src:
        src = src.replace(old_browse, new_browse)
        print(f"Patched generic BROWSE in {py_file}")
    elif 'self.btn_browse.pack(' in src:
        # Fallback to pack regex
        btn_injection = f'''\\\\g<0>
        self.btn_download = ctk.CTkButton(self.card_cred if hasattr(self, 'card_cred') else self.right_frame if hasattr(self, 'right_frame') else self.container if hasattr(self, 'container') else self.sidebar if hasattr(self, 'sidebar') else self, text="📥 Download Sample Excel", command=self.download_sample, fg_color="#43a047", hover_color="#2e7d32", height=35)
        self.btn_download.pack(fill="x", padx=15, pady=(5, 15))'''
        src = re.sub(r'self\\.btn_browse\\.pack\\([^)]*\\)', btn_injection, src)
        print(f"Patched with pack fallback in {py_file}")
    
    with open(py_file, 'w', encoding='utf-8') as f:
        f.write(src)
