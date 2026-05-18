import os
import re

filepath = r"c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\GST\GST Bot\gst_pro_app.py"

with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Update GSTWorker.run to handle dicts
content = content.replace(
'''            if self.manual_gstins:
                gstin_list = [str(x).strip() for x in self.manual_gstins if str(x).strip()]
            else:''',
'''            if self.manual_gstins:
                gstin_list = self.manual_gstins
            else:'''
)

content = content.replace(
'''                if 'GSTIN' not in df.columns:
                    self.log("❌ Error: Excel must have a column named 'GSTIN' (case-insensitive)", "error")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    self.driver = None
                    return

                gstin_list = df['GSTIN'].astype(str).unique().tolist()

            gstin_list = list(dict.fromkeys(gstin_list))
            results = []
            total = len(gstin_list)

            self.log(f"📂 Found {total} unique GSTINs. Starting Batch Process...", "info")

            for index, gstin in enumerate(gstin_list):''',
'''                orig_cols = {str(c).strip().upper(): c for c in df.columns}
                if 'GSTIN' not in orig_cols:
                    self.log("❌ Error: Excel must have a column named 'GSTIN' (case-insensitive)", "error")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    self.driver = None
                    return
                gstin_col = orig_cols['GSTIN']
                cname_col = orig_cols.get('CLIENT NAME') or orig_cols.get('CLIENTNAME')
                
                gstin_list = []
                seen = set()
                for _, r in df.iterrows():
                    g = str(r[gstin_col]).strip()
                    if g not in seen and g.lower() != 'nan':
                        seen.add(g)
                        c = str(r[cname_col]).strip() if cname_col else ""
                        if c.lower() == 'nan': c = ""
                        gstin_list.append({"GSTIN": g, "ClientName": c})

            results = []
            total = len(gstin_list)

            self.log(f"📂 Found {total} unique GSTINs. Starting Batch Process...", "info")

            for index, item in enumerate(gstin_list):
                if isinstance(item, dict):
                    gstin = item.get("GSTIN", "")
                    cname = item.get("ClientName", "")
                else:
                    gstin = str(item)
                    cname = ""'''
)

# 2. Update folder saving at the end and in stop block
content = content.replace(
'''                    base_dir = os.path.join(os.getcwd(), "GST Downloaded", "GST Verifier", "reports")
                    os.makedirs(base_dir, exist_ok=True)
                    output_file = os.path.join(base_dir, f"GST_Report_{timestamp}.xlsx")''',
'''                    if len(gstin_list) == 1 and isinstance(gstin_list[0], dict) and gstin_list[0].get("ClientName"):
                        folder_name = gstin_list[0].get("ClientName")
                    elif len(gstin_list) == 1 and isinstance(gstin_list[0], dict):
                        folder_name = gstin_list[0].get("GSTIN")
                    elif len(gstin_list) == 1:
                        folder_name = str(gstin_list[0])
                    else:
                        folder_name = "Bulk_Reports"
                    
                    import re as _re_tmp
                    folder_name = _re_tmp.sub(r'[\\\\/*?:"<>|]', "", folder_name).strip()
                    
                    base_dir = os.path.join(os.getcwd(), "GST Downloaded", "GST Verifier", folder_name)
                    os.makedirs(base_dir, exist_ok=True)
                    output_file = os.path.join(base_dir, f"GST_Report_{timestamp}.xlsx")'''
)


# 3. Update add_gstin UI and saving logic
content = content.replace(
'''        ctk.CTkLabel(card, text="GSTIN Number", font=("Segoe UI", 12)).pack(anchor="w")
        ent = ctk.CTkEntry(card, placeholder_text="e.g. 27ABCDE1234F1Z5", height=36)
        ent.pack(fill="x", pady=(4, 16))
        ent.focus_set()
        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x")
        def _save():
            g = ent.get().strip().upper()
            if not g:
                messagebox.showwarning("Missing", "Enter a GSTIN number.", parent=dialog)
                return
            db_path = _os.path.join(_os.environ.get("APPDATA", _os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
            _os.makedirs(_os.path.dirname(db_path), exist_ok=True)
            try:
                conn = _sq.connect(db_path)
                conn.execute("CREATE TABLE IF NOT EXISTS gst_gstin_list (id INTEGER PRIMARY KEY AUTOINCREMENT, gstin TEXT UNIQUE)")
                conn.execute("INSERT OR REPLACE INTO gst_gstin_list (gstin) VALUES (?)", (g,))
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dialog)
                return
            if g not in self.gstin_list:
                self.gstin_list.append(g)''',
'''        ctk.CTkLabel(card, text="Client Name (Optional)", font=("Segoe UI", 12)).pack(anchor="w")
        ent_c = ctk.CTkEntry(card, placeholder_text="Enter Client Name", height=36)
        ent_c.pack(fill="x", pady=(4, 8))
        
        ctk.CTkLabel(card, text="GSTIN Number", font=("Segoe UI", 12)).pack(anchor="w")
        ent = ctk.CTkEntry(card, placeholder_text="e.g. 27ABCDE1234F1Z5", height=36)
        ent.pack(fill="x", pady=(4, 16))
        ent.focus_set()
        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x")
        def _save():
            c = ent_c.get().strip()
            g = ent.get().strip().upper()
            if not g:
                messagebox.showwarning("Missing", "Enter a GSTIN number.", parent=dialog)
                return
            db_path = _os.path.join(_os.environ.get("APPDATA", _os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
            _os.makedirs(_os.path.dirname(db_path), exist_ok=True)
            try:
                conn = _sq.connect(db_path)
                conn.execute("CREATE TABLE IF NOT EXISTS gst_gstin_list (id INTEGER PRIMARY KEY AUTOINCREMENT, gstin TEXT UNIQUE)")
                try: conn.execute("ALTER TABLE gst_gstin_list ADD COLUMN client_name TEXT")
                except: pass
                
                existing = conn.execute("SELECT id FROM gst_gstin_list WHERE gstin=?", (g,)).fetchone()
                if existing:
                    conn.execute("UPDATE gst_gstin_list SET client_name=? WHERE gstin=?", (c, g))
                else:
                    conn.execute("INSERT INTO gst_gstin_list (gstin, client_name) VALUES (?, ?)", (g, c))
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dialog)
                return
            found = False
            for item in self.gstin_list:
                if isinstance(item, dict) and item.get("GSTIN") == g:
                    found = True; break
                elif item == g:
                    found = True; break
            if not found:
                self.gstin_list.append({"GSTIN": g, "ClientName": c})'''
)

# 4. Update load_gstins_from_db
content = content.replace(
'''        try:
            conn = _sq.connect(db_path)
            rows = conn.execute("SELECT gstin FROM gst_gstin_list ORDER BY gstin").fetchall()
            conn.close()
        except Exception:
            rows = []
        if not rows:
            messagebox.showinfo("No GSTINs", "No GSTINs saved yet. Add via the Add GSTIN button.", parent=self)
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Load GSTINs")
        dialog.geometry("400x460")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        ctk.CTkLabel(dialog, text="Select GSTINs to Load",
                     font=("Segoe UI", 14, "bold")).pack(pady=(16, 8))
        sel_all_var = ctk.BooleanVar()
        vars_ = {}
        def _toggle_all():
            state = sel_all_var.get()
            for v in vars_.values():
                v.set(state)
        ctk.CTkCheckBox(dialog, text="Select All", variable=sel_all_var,
                        command=_toggle_all,
                        font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
        scroll = ctk.CTkScrollableFrame(dialog, height=300)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        for (g,) in rows:
            v = ctk.BooleanVar()
            ctk.CTkCheckBox(scroll, text=g, variable=v).pack(anchor="w", padx=10, pady=3)
            vars_[g] = v
        def _load():
            selected = [g for g, v in vars_.items() if v.get()]
            if not selected:
                messagebox.showwarning("No Selection", "Select at least one GSTIN.", parent=dialog)
                return
            self.gstin_list = selected''',
'''        try:
            conn = _sq.connect(db_path)
            cur = conn.cursor()
            cur.execute("SELECT * FROM gst_gstin_list ORDER BY gstin")
            cols = [d[0] for d in cur.description]
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]
            conn.close()
        except Exception:
            rows = []
        if not rows:
            messagebox.showinfo("No GSTINs", "No GSTINs saved yet. Add via the Add GSTIN button.", parent=self)
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Load GSTINs")
        dialog.geometry("400x460")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        ctk.CTkLabel(dialog, text="Select GSTINs to Load",
                     font=("Segoe UI", 14, "bold")).pack(pady=(16, 8))
        sel_all_var = ctk.BooleanVar()
        vars_ = {}
        def _toggle_all():
            state = sel_all_var.get()
            for v in vars_.values():
                v.set(state)
        ctk.CTkCheckBox(dialog, text="Select All", variable=sel_all_var,
                        command=_toggle_all,
                        font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
        scroll = ctk.CTkScrollableFrame(dialog, height=300)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        for rdata in rows:
            g = rdata.get("gstin", "")
            c = rdata.get("client_name") or ""
            v = ctk.BooleanVar()
            disp = f"{c} ({g})" if c else g
            ctk.CTkCheckBox(scroll, text=disp, variable=v).pack(anchor="w", padx=10, pady=3)
            vars_[(g, c)] = v
        def _load():
            selected = [{"GSTIN": g, "ClientName": c} for (g, c), v in vars_.items() if v.get()]
            if not selected:
                messagebox.showwarning("No Selection", "Select at least one GSTIN.", parent=dialog)
                return
            self.gstin_list = selected'''
)

# 5. Update view_gstin_data
content = content.replace(
'''            try:
                c = _sq.connect(db_path)
                rs = c.execute("SELECT id, gstin FROM gst_gstin_list ORDER BY gstin").fetchall()
                c.close()
            except Exception:
                rs = []
            if not rs:
                ctk.CTkLabel(scroll, text="No GSTINs saved yet.",
                             font=("Segoe UI", 12), text_color="gray").pack(pady=30)
                return
            for rid, gnum in rs:
                row_f = ctk.CTkFrame(scroll, fg_color=("#f8fafc", "#273549"),
                                     corner_radius=8, border_width=1,
                                     border_color=("#e2e8f0", "#334155"))
                row_f.pack(fill="x", padx=4, pady=4)
                row_f.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(row_f, text=f"  {gnum}",
                             font=("Segoe UI", 13, "bold"),
                             anchor="w").grid(row=0, column=0, sticky="w", padx=12, pady=10)
                def _del(r=rid, n=gnum):
                    if not messagebox.askyesno("Delete", f"Delete GSTIN '{n}'?", parent=dialog):
                        return
                    try:
                        c2 = _sq.connect(db_path)
                        c2.execute("DELETE FROM gst_gstin_list WHERE id=?", (r,))
                        c2.commit()
                        c2.close()
                    except Exception:
                        pass
                    if n in self.gstin_list:
                        self.gstin_list.remove(n)''',
'''            try:
                c = _sq.connect(db_path)
                cur = c.cursor()
                cur.execute("SELECT * FROM gst_gstin_list ORDER BY gstin")
                cols = [d[0] for d in cur.description]
                rs = [dict(zip(cols, r)) for r in cur.fetchall()]
                c.close()
            except Exception:
                rs = []
            if not rs:
                ctk.CTkLabel(scroll, text="No GSTINs saved yet.",
                             font=("Segoe UI", 12), text_color="gray").pack(pady=30)
                return
            for rdata in rs:
                rid = rdata.get("id")
                gnum = rdata.get("gstin", "")
                cname = rdata.get("client_name") or ""
                row_f = ctk.CTkFrame(scroll, fg_color=("#f8fafc", "#273549"),
                                     corner_radius=8, border_width=1,
                                     border_color=("#e2e8f0", "#334155"))
                row_f.pack(fill="x", padx=4, pady=4)
                row_f.grid_columnconfigure(0, weight=1)
                disp = f"  {cname} ({gnum})" if cname else f"  {gnum}"
                ctk.CTkLabel(row_f, text=disp,
                             font=("Segoe UI", 13, "bold"),
                             anchor="w").grid(row=0, column=0, sticky="w", padx=12, pady=10)
                def _del(r=rid, n=gnum):
                    if not messagebox.askyesno("Delete", f"Delete GSTIN '{n}'?", parent=dialog):
                        return
                    try:
                        c2 = _sq.connect(db_path)
                        c2.execute("DELETE FROM gst_gstin_list WHERE id=?", (r,))
                        c2.commit()
                        c2.close()
                    except Exception:
                        pass
                    new_list = []
                    for item in self.gstin_list:
                        if isinstance(item, dict) and item.get("GSTIN") == n: continue
                        if item == n: continue
                        new_list.append(item)
                    self.gstin_list = new_list'''
)

with open(filepath, "w", encoding="utf-8") as f:
    f.write(content)

print("Patch applied to gst_pro_app.py")
