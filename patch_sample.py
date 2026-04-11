import os
import glob
import re

mapping = {
    r"GST\GST 2B Downloader\main.py": "GSTR2B Sample File.xlsx",
    r"GST\GST 3B Downloader\main.py": "GSTR3B Sample File.xlsx",
    r"GST\GST Bot\gst_pro_app.py": "GST Verification Tools Sample File.xlsx",
    r"GST\GST Challan Downloader\main.py": "GST Challan Downloader Sample File.xlsx",
    r"GST\GST R1 Downloader\mai.py": "GSTR1 Sample File.xlsx",
    r"GST\IMS Downloader\main.py": "IMS Sample File.xlsx",
    r"GST\R1 PDF Downloader\main.py": "GSTR1 pdf Sample File.xlsx",
    r"Income Tax\26 AS Downlaoder\main.py": "Income Tax Sample File.xlsx",
    r"Income Tax\Challan Downloader\main.py": "Income Tax Sample File.xlsx",
    r"Income Tax\ITR - Bot\main.py": "Income Tax Sample File.xlsx",
}

for py_file, sample_name in mapping.items():
    if not os.path.exists(py_file):
        print(f"Not found: {py_file}")
        continue
    
    with open(py_file, 'r', encoding='utf-8') as f:
        src = f.read()
        
    if "def download_sample" in src:
        print(f"Skipping {py_file}, already modified.")
        continue

    # Add button after Browse File using regex
    # We look for something like .pack(..., pady=(XX, 15)) or similar for btn_browse
    
    # 1. First add the download_sample method just before def browse_file(self):
    method_injection = f'''
    def download_sample(self):
        import shutil
        import os
        from tkinter import messagebox
        sample_path = os.path.join(os.path.dirname(__file__), "{sample_name}")
        if not os.path.exists(sample_path):
            messagebox.showerror("Download Error", f"Sample file not found: {{sample_path}}")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="{sample_name}", filetypes=[("Excel", "*.xlsx")])
        if save_path:
            try:
                shutil.copy2(sample_path, save_path)
                messagebox.showinfo("Success", f"Sample downloaded to {{save_path}}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download: {{e}}")

    def browse_file(self):'''

    src = re.sub(r'\+?\s*def browse_file\(self[^)]*\):', method_injection, src)

    # 2. Add the button itself right after the browse button layout. 
    # Usually it's self.btn_browse = ctk.CTkButton(...) and then self.btn_browse.pack(...)
    
    # Let's use a simpler approach. If there's `self.btn_browse.pack(` we append
    btn_injection = f'''
        \\g<0>
        self.btn_download = ctk.CTkButton(self.card_cred if hasattr(self, 'card_cred') else self.right_frame if hasattr(self, 'right_frame') else self.container if hasattr(self, 'container') else self.sidebar if hasattr(self, 'sidebar') else self, text="📥 Download Sample Excel", command=self.download_sample, fg_color="#43a047", hover_color="#2e7d32", height=35)
        self.btn_download.pack(fill="x", padx=15, pady=(5, 15))'''

    src = re.sub(r'self\.btn_browse\.pack\([^)]*\)', btn_injection, src)
    
    with open(py_file, 'w', encoding='utf-8') as f:
        f.write(src)
        
    print(f"Patched: {py_file}")

