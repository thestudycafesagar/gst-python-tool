import os
import re

files = [
    r"GST\GST Challan Downloader\main.py",
    r"GST\GST R1 Downloader\mai.py",
    r"GST\R1 PDF Downloader\main.py",
    r"GST\IMS Downloader\main.py"
]

ui_target = '''        # Quarter & Month
        self.frm_qtr = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_qtr.pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(self.frm_qtr, text="Quarter:", width=140, anchor="w").pack(side="left")
        self.cb_qtr = ctk.CTkComboBox(self.frm_qtr, 
                                      values=["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)", 
                                              "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"],
                                      command=self.update_months_based_on_qtr, width=150)
        self.cb_qtr.set("Quarter 1 (Apr - Jun)")
        self.cb_qtr.pack(side="right", expand=True, fill="x")

        # Month
        self.frm_mon = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_mon.pack(fill="x", padx=15, pady=(2, 15))
        ctk.CTkLabel(self.frm_mon, text="Month:", width=140, anchor="w").pack(side="left")
        
        all_months = ["April", "May", "June", "July", "August", "September", 
                      "October", "November", "December", "January", "February", "March"]
        self.cb_month = ctk.CTkComboBox(self.frm_mon, values=all_months, 
                                        command=self.update_qtr_based_on_month, width=150)
        self.cb_month.set("April")
        self.cb_month.pack(side="right", expand=True, fill="x")'''

ui_replace = '''        # Checkboxes Frame (replaces dropdowns)
        self.frm_checkboxes = ctk.CTkFrame(self.card_period, fg_color="transparent")
        self.frm_checkboxes.pack(fill="both", expand=True, padx=15, pady=5)
        self.period_checkbox_vars = {}
        '''

toggle_pattern = re.compile(r'    def toggle_inputs.*?def _get_saved_user_id', re.DOTALL)

toggle_replace = '''    def toggle_inputs(self, mode_choice=None):
        if mode_choice and hasattr(self, "period_mode_var"):
            self.period_mode_var.set(mode_choice)
        
        mode = self.period_mode_var.get() if hasattr(self, "period_mode_var") else "Monthly"
        
        if hasattr(self, "frm_checkboxes"):
            for w in self.frm_checkboxes.winfo_children():
                w.destroy()
            self.period_checkbox_vars.clear()

            if mode == "Monthly":
                items = ["April", "May", "June", "July", "August", "September",
                         "October", "November", "December", "January", "February", "March"]
                cols = 3
            else:
                items = ["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)",
                         "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"]
                cols = 2
            
            for i, item in enumerate(items):
                var = ctk.BooleanVar(value=False)
                self.period_checkbox_vars[item] = var
                chk = ctk.CTkCheckBox(self.frm_checkboxes, text=item, variable=var, font=("Segoe UI", 12))
                chk.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="w")

    def update_qtr_based_on_month(self, choice):
        pass

    def update_months_based_on_qtr(self, choice):
        pass

    def _get_saved_user_id'''

start_process_target = '''        settings = {
            "year": self.cb_year.get(),
            "month": self.cb_month.get(),
            "quarter": self.cb_qtr.get(),
            "period_mode": self.period_mode_var.get(),
            "all_quarters": False
        }'''

start_process_replace = '''        mode = self.period_mode_var.get() if hasattr(self, "period_mode_var") else "Monthly"
        selected_periods = [lbl for lbl, var in getattr(self, "period_checkbox_vars", {}).items() if var.get()]
        if not selected_periods:
            messagebox.showerror("Error", "Please select at least one period.")
            return

        tasks = []
        q_map_rev = {
            "April": "Quarter 1 (Apr - Jun)", "May": "Quarter 1 (Apr - Jun)", "June": "Quarter 1 (Apr - Jun)",
            "July": "Quarter 2 (Jul - Sep)", "August": "Quarter 2 (Jul - Sep)", "September": "Quarter 2 (Jul - Sep)",
            "October": "Quarter 3 (Oct - Dec)", "November": "Quarter 3 (Oct - Dec)", "December": "Quarter 3 (Oct - Dec)",
            "January": "Quarter 4 (Jan - Mar)", "February": "Quarter 4 (Jan - Mar)", "March": "Quarter 4 (Jan - Mar)"
        }
        q_last_month = {
            "Quarter 1 (Apr - Jun)": "June",
            "Quarter 2 (Jul - Sep)": "September",
            "Quarter 3 (Oct - Dec)": "December",
            "Quarter 4 (Jan - Mar)": "March"
        }

        if mode == "Monthly":
            for m in selected_periods:
                tasks.append({"q": q_map_rev.get(m, ""), "m": m})
        else:
            for q in selected_periods:
                tasks.append({"q": q, "m": q_last_month.get(q, "")})

        settings = {
            "year": self.cb_year.get(),
            "period_mode": mode,
            "tasks": tasks,
            "all_quarters": False
        }'''

worker_pattern = re.compile(r'            selected_q = self\.settings\[\'quarter\'\].*?self\.log\(f"   .. Mode: Monthly \(\{selected_m\}\)"\)', re.DOTALL)

worker_replace = '''            if "tasks" in self.settings:
                tasks = self.settings["tasks"]
                self.log(f"   ?? Bulk Mode: processing {len(tasks)} periods")
            else:
                selected_q = self.settings['quarter']
                period_mode = self.settings.get('period_mode', 'Monthly')
                if selected_q not in q_map:
                    return "Config Error", "Invalid Month/Quarter Selection"

                if period_mode == "Quarterly":
                    selected_m = q_map[selected_q][-1]
                    tasks = [{"q": selected_q, "m": selected_m}]
                    self.log(f"   ?? Mode: Quarterly ({selected_q} -> {selected_m})")
                else:
                    selected_m = self.settings['month']
                    if selected_m not in q_map[selected_q]:
                        return "Config Error", "Invalid Month/Quarter Selection"
                    tasks = [{"q": selected_q, "m": selected_m}]
                    self.log(f"   ?? Mode: Monthly ({selected_m})")'''


for f in files:
    if not os.path.exists(f): continue
    with open(f, "r", encoding="utf-8") as file:
        content = file.read()
    
    c1 = content.replace(ui_target, ui_replace)
    c2 = toggle_pattern.sub(toggle_replace, c1)
    c3 = c2.replace(start_process_target, start_process_replace)
    
    # Optional fallback logic replacement in worker
    c4 = worker_pattern.sub(worker_replace, c3)
    
    with open(f, "w", encoding="utf-8") as file:
        file.write(c4)
    print(f"Patched {f}")

