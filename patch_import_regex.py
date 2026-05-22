import os
import re

f = "GST_Suite.py"
with open(f, "r", encoding="utf-8") as file:
    content = file.read()

pattern = re.compile(r'(_ov_refresh\(\)\n\s+ctk\.CTkButton\(rf.*?pack\(fill="x", pady=\(15, 0\)\))', re.DOTALL)

replacement = r'''\1

            # -- Import / Export Feature --
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
                path = fd.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls")])
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
            btn_imp.grid(row=0, column=1, sticky="ew", padx=(5,0))'''

if pattern.search(content):
    c1 = pattern.sub(replacement, content)
    with open(f, "w", encoding="utf-8") as file:
        file.write(c1)
    print("Patched with regex!")
else:
    print("Regex failed to find pattern.")
