import re
path = r"Income Tax\Challan Downloader\demand_checker_app.py"

with open(path, "r", encoding="utf-8") as f:
    text = f.read()

old_btn = 'ctk.CTkButton(f_frame, text="BROWSE", command=self.browse_file_demand, width=100).pack(side="right")'
new_btn = 'ctk.CTkButton(f_frame, text="📥 DOWNLOAD SAMPLE", command=self.download_sample, width=150, fg_color="#43a047", hover_color="#2e7d32").pack(side="right", padx=(0, 10))\n        ' + old_btn

text = text.replace(old_btn, new_btn)

method_injection = '''
    def download_sample(self):
        import shutil
        import os
        from tkinter import messagebox, filedialog
        sample_path = os.path.join(os.path.dirname(__file__), "Income Tax Sample File.xlsx")
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Income Tax Sample File.xlsx", filetypes=[("Excel", "*.xlsx")])
        if save_path:
            try:
                shutil.copy2(sample_path, save_path)
                messagebox.showinfo("Success", f"Sample downloaded to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download: {e}")

    def browse_file_demand(self):'''

if "def download_sample" not in text:
    text = re.sub(r'\+?\s*def browse_file_demand\(self[^)]*\):', method_injection, text)

with open(path, "w", encoding="utf-8") as f:
    f.write(text)
