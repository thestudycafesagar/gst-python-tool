import os

files = [
    r"GST\GST 2B Downloader\main.py",
    r"GST\GST 3B Downloader\main.py",
    r"GST\GST Challan Downloader\main.py",
    r"GST\GST R1 Downloader\mai.py",
    r"GST\R1 PDF Downloader\main.py"
]

target = '''            if mode == "Monthly":
                items = ["Apr", "May", "Jun", "Jul", "Aug", "Sep",
                         "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
                cols = 6
            else:
                items = ["Q1 (Apr-Jun)", "Q2 (Jul-Sep)",
                         "Q3 (Oct-Dec)", "Q4 (Jan-Mar)"]
                cols = 4
            
            for i, item in enumerate(items):
                var = ctk.BooleanVar(value=False)
                self.period_checkbox_vars[item] = var
                chk = ctk.CTkCheckBox(self.frm_checkboxes, text=item, variable=var, font=("Segoe UI", 12))
                chk.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="w")'''

replacement = '''            def toggle_select_all():
                state = select_all_var.get()
                for var in self.period_checkbox_vars.values():
                    var.set(state)

            top_bar = ctk.CTkFrame(self.frm_checkboxes, fg_color="transparent")
            top_bar.pack(fill="x", pady=(0, 5))
            
            select_all_var = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(top_bar, text="Select All", variable=select_all_var, command=toggle_select_all, font=("Segoe UI", 12, "bold"), text_color="#10B981").pack(side="left")

            chk_grid = ctk.CTkFrame(self.frm_checkboxes, fg_color="transparent")
            chk_grid.pack(fill="both", expand=True)

            if mode == "Monthly":
                items = ["Apr", "May", "Jun", "Jul", "Aug", "Sep",
                         "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
                cols = 6
            else:
                items = ["Q1 (Apr-Jun)", "Q2 (Jul-Sep)",
                         "Q3 (Oct-Dec)", "Q4 (Jan-Mar)"]
                cols = 4
            
            for i, item in enumerate(items):
                var = ctk.BooleanVar(value=False)
                self.period_checkbox_vars[item] = var
                chk = ctk.CTkCheckBox(chk_grid, text=item, variable=var, font=("Segoe UI", 12))
                chk.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="w")'''

for f in files:
    if not os.path.exists(f): continue
    with open(f, "r", encoding="utf-8") as file:
        content = file.read()
    
    if target in content:
        c1 = content.replace(target, replacement)
        with open(f, "w", encoding="utf-8") as file:
            file.write(c1)
        print(f"Patched {f}")
    else:
        print(f"Target not found in {f}")

