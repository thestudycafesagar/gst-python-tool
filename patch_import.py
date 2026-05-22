import os

f = "GST_Suite.py"
with open(f, "r", encoding="utf-8") as file:
    content = file.read()

target = 'ctk.CTkButton(rf, text="o.  Save Profile", command=_ov_save,\n                            fg_color="#059669", hover_color="#047857",\n                            height=50, font=("Segoe UI", 16, "bold")).pack(fill="x", pady=(15, 0))'

replacement = """ctk.CTkButton(rf, text="o.  Save Profile", command=_ov_save,
                          fg_color="#059669", hover_color="#047857",
                          height=50, font=("Segoe UI", 16, "bold")).pack(fill="x", pady=(15, 0))

            import_frame = ctk.CTkFrame(rf, fg_color="transparent")
            import_frame.pack(fill="x", pady=(15, 0))
            import_frame.grid_columnconfigure((0,1), weight=1)

            def _dl_sample():
                from tkinter import filedialog as fd
                import pandas as pd
                from tkinter import messagebox as mb
                path = fd.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Sample_{cat_label}_Profiles.xlsx", filetypes=[("Excel", "*.xlsx")])
                if not path: return
                if is_it:
                    df = pd.DataFrame([{"Client Name": "John Doe", "Username (PAN)": "ABCDE1234F", "Password": "Pass123", "Date of Birth (DD/MM/YYYY)": "01/01/1990"}])
                else:
                    df = pd.DataFrame([{"Client Name": "Studycafe", "Username (GSTIN)": "07AAAAA0000A1Z5", "Password": "Pass123", "Filing Frequency": "Monthly"}])
                try:
                    df.to_excel(path, index=False)
                    mb.showinfo("Success", f"Sample downloaded successfully to:\\n{path}")
                except Exception as e:
                    mb.showerror("Error", str(e))

            def _import_excel():
                from tkinter import filedialog as fd
                import pandas as pd
                from tkinter import messagebox as mb
                path = fd.askopenfilename(filetypes=[("Excel", "*.xlsx", "*.xls")])
                if not path: return
                try:
                    df = pd.read_excel(path)
                    conn = _get_ov_db()
                    count = 0
                    for _, row in df.iterrows():
                        row = row.fillna("")
                        c = str(row.get("Client Name", "")).strip()
                        u = str(row.iloc[1] if "Username" not in str(df.columns[1]) else row.get(df.columns[1], "")).strip() 
                        if not u: u = str(row.get("Username (GSTIN)", str(row.get("Username (PAN)", "")))).strip()
                        p = str(row.get("Password", "")).strip()
                        
                        if not u or not p: continue
                        
                        if is_it:
                            d = str(row.get("Date of Birth (DD/MM/YYYY)", "")).strip()
                            existing = conn.execute(f"SELECT id FROM {table_name} WHERE username=?", (u,)).fetchone()
                            if existing:
                                conn.execute(f"UPDATE {table_name} SET password=?, client_name=?, dob=? WHERE username=?", (p, c, d, u))
                            else:
                                conn.execute(f"INSERT INTO {table_name} (username, password, client_name, dob) VALUES (?,?,?,?)", (u, p, c, d))
                        else:
                            f = str(row.get("Filing Frequency", "Monthly")).strip()
                            if not f: f = "Monthly"
                            existing = conn.execute(f"SELECT id FROM {table_name} WHERE username=?", (u,)).fetchone()
                            if existing:
                                conn.execute(f"UPDATE {table_name} SET password=?, client_name=?, filing_frequency=? WHERE username=?", (p, c, f, u))
                            else:
                                conn.execute(f"INSERT INTO {table_name} (username, password, client_name, filing_frequency) VALUES (?,?,?,?)", (u, p, c, f))
                        count += 1
                    conn.commit()
                    conn.close()
                    _ov_refresh()
                    mb.showinfo("Success", f"Imported {count} profiles successfully!")
                except Exception as e:
                    mb.showerror("Error", f"Failed to import:\\n{e}")

            btn_dl = ctk.CTkButton(import_frame, text="Download Sample", command=_dl_sample, 
                                   fg_color="#334155", hover_color="#475569", height=38, font=("Segoe UI", 12, "bold"))
            btn_dl.grid(row=0, column=0, sticky="ew", padx=(0,5))
            btn_imp = ctk.CTkButton(import_frame, text="Import Excel", command=_import_excel, 
                                    fg_color="#2563EB", hover_color="#1D4ED8", height=38, font=("Segoe UI", 12, "bold"))
            btn_imp.grid(row=0, column=1, sticky="ew", padx=(5,0))"""

if target in content:
    c1 = content.replace(target, replacement)
    with open(f, "w", encoding="utf-8") as file:
        file.write(c1)
    print("Patched GST_Suite.py")
else:
    print("Target not found. Doing fallback.")
    target_fallback = 'ctk.CTkButton(rf, text="?  Save Profile", command=_ov_save,\n                          fg_color="#059669", hover_color="#047857",\n                          height=50, font=("Segoe UI", 16, "bold")).pack(fill="x", pady=(15, 0))'
    if target_fallback in content:
        c1 = content.replace(target_fallback, replacement)
        with open(f, "w", encoding="utf-8") as file:
            file.write(c1)
        print("Patched fallback")
    else:
        print("Still not found.")
