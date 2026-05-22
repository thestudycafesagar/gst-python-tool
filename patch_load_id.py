"""
Patch script v2: Adds search box + selected counter to ALL load_id_pass dialogs
Handles both GST-style and IT-style dialog implementations.
Run from the project root: python patch_load_id.py
"""
import os

ROOT = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────────────────────────────────────
# PATTERN A  ─  GST tools (already patched 5 files; kept here for reference)
# PATTERN B  ─  IT tools (different structure, no foot frame)
# ─────────────────────────────────────────────────────────────────────────────

# ── IT tool OLD block ─────────────────────────────────────────────────────────
IT_OLD = '''\
        ctk.CTkLabel(dialog, text="Select IT Profiles to Load", font=("Segoe UI", 14, "bold")).pack(pady=(16, 8))
        
        sel_all_var = ctk.BooleanVar()
        vars_ = {}
        
        def _toggle_all():
            state = sel_all_var.get()
            for v in vars_.values():
                v.set(state)
        
        ctk.CTkCheckBox(dialog, text="Select All", variable=sel_all_var, command=_toggle_all,
                        font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
                        
        scroll = ctk.CTkScrollableFrame(dialog, height=300)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        
        for u, p, d, cname in rows:
            v = ctk.BooleanVar()
            vars_[(u, p, d)] = v
            disp_text = f"{u} ({cname})" if cname else u
            ctk.CTkCheckBox(scroll, text=disp_text, variable=v).pack(anchor="w", padx=10, pady=3)
            
        def _load():
            selected = [{"PAN": u, "Password": p, "DOB": d} for (u, p, d), v in vars_.items() if v.get()]
            if not selected:
                return
            
            self.manual_credentials = selected
            self.excel_file_path = ""
            self._refresh_manual_controls()
            
            n = len(selected)
            label = selected[0]["PAN"] if n == 1 else f"Loaded {n} profiles"
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, label)
            
            self.log_to_gui(f"Loaded {n} Profiles from database")
            dialog.destroy()
            
        ctk.CTkButton(dialog, text="✅ Load Selected", command=_load, height=35).pack(pady=10)'''

IT_NEW = '''\
        ctk.CTkLabel(dialog, text="Select IT Profiles to Load", font=("Segoe UI", 14, "bold")).pack(pady=(16, 6))

        # ── Search box ────────────────────────────────────────────────────────
        search_var = ctk.StringVar()
        search_entry = ctk.CTkEntry(dialog, placeholder_text="🔍  Search by name or username...",
                                    textvariable=search_var, height=34)
        search_entry.pack(fill="x", padx=16, pady=(0, 6))

        # ── Select-all + counter row ──────────────────────────────────────────
        top_row = ctk.CTkFrame(dialog, fg_color="transparent")
        top_row.pack(fill="x", padx=16, pady=(0, 4))
        sel_all_var = ctk.BooleanVar()
        counter_var = ctk.StringVar(value="0 selected")
        counter_lbl = ctk.CTkLabel(top_row, textvariable=counter_var,
                                   font=("Segoe UI", 11, "bold"), text_color="#059669")
        counter_lbl.pack(side="right")

        # ── Scrollable profile list ───────────────────────────────────────────
        scroll = ctk.CTkScrollableFrame(dialog, height=240)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 6))
        vars_ = {}
        widgets_ = {}

        for u, p, d, cname in rows:
            v = ctk.BooleanVar()
            vars_[(u, p, d)] = v
            disp_text = f"{cname} ({u})" if cname else u
            chk = ctk.CTkCheckBox(scroll, text=disp_text, variable=v,
                                  command=lambda: _refresh_counter())
            chk.pack(anchor="w", padx=10, pady=3)
            widgets_[(u, p, d)] = (chk, disp_text)

        def _refresh_counter():
            n = sum(1 for v in vars_.values() if v.get())
            counter_var.set(f"{n} selected")
            visible = [w for w, _ in widgets_.values() if w.winfo_ismapped()]
            sel_all_var.set(n > 0 and n == len(visible))

        def _toggle_all():
            state = sel_all_var.get()
            for key, (chk, _) in widgets_.items():
                if chk.winfo_ismapped():
                    vars_[key].set(state)
            _refresh_counter()

        def _on_search(*_):
            q = search_var.get().strip().lower()
            for key, (chk, disp) in widgets_.items():
                if q == "" or q in disp.lower():
                    chk.pack(anchor="w", padx=10, pady=3)
                else:
                    chk.pack_forget()
            _refresh_counter()

        ctk.CTkCheckBox(top_row, text="Select All", variable=sel_all_var, command=_toggle_all,
                        font=("Segoe UI", 12, "bold")).pack(side="left")
        search_var.trace_add("write", _on_search)
        search_entry.focus_set()

        def _load():
            selected = [{"PAN": u, "Password": p, "DOB": d} for (u, p, d), v in vars_.items() if v.get()]
            if not selected:
                messagebox.showwarning("No Selection", "Please select at least one profile.", parent=dialog)
                return
            self.manual_credentials = selected
            self.excel_file_path = ""
            self._refresh_manual_controls()
            n = len(selected)
            label = selected[0]["PAN"] if n == 1 else f"Loaded {n} profiles"
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, label)
            self.log_to_gui(f"Loaded {n} Profiles from database")
            dialog.destroy()

        foot = ctk.CTkFrame(dialog, fg_color="transparent")
        foot.pack(fill="x", padx=16, pady=(0, 14))
        ctk.CTkButton(foot, text="Cancel", width=110, command=dialog.destroy).pack(side="right")
        ctk.CTkButton(foot, text="✅ Load Selected", width=140, fg_color="#059669",
                      hover_color="#047857", command=_load).pack(side="right", padx=(0, 8))'''

# ── 26AS / Refund Checker variant (uses entry_file_26as instead of entry_file) ──
AS_OLD_SUFFIX = '''\
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, label)
            
            self.log_to_gui(f"Loaded {n} Profiles from database")
            dialog.destroy()
            
        ctk.CTkButton(dialog, text="✅ Load Selected", command=_load, height=35).pack(pady=10)'''

# ITR Bot variant (uses entry_file too but different class)
ITRBOT_OLD = '''\
        ctk.CTkLabel(dialog, text="Select IT Profiles to Load", font=("Segoe UI", 14, "bold")).pack(pady=(16, 8))
        
        sel_all_var = ctk.BooleanVar()
        vars_ = {}
        
        def _toggle_all():
            state = sel_all_var.get()
            for v in vars_.values():
                v.set(state)
        
        ctk.CTkCheckBox(dialog, text="Select All", variable=sel_all_var, command=_toggle_all,
                        font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
                        
        scroll = ctk.CTkScrollableFrame(dialog, height=300)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        
        for u, p, d, cname in rows:
            v = ctk.BooleanVar()
            vars_[(u, p, d)] = v
            disp_text = f"{u} ({cname})" if cname else u
            ctk.CTkCheckBox(scroll, text=disp_text, variable=v).pack(anchor="w", padx=10, pady=3)
            
        def _load():
            selected = [{"PAN": u, "Password": p, "DOB": d} for (u, p, d), v in vars_.items() if v.get()]
            if not selected:
                return'''


def patch_file(rel_path, old_text, new_text, old_geom, new_geom):
    abs_path = os.path.join(ROOT, rel_path)
    if not os.path.exists(abs_path):
        print(f"  SKIP (not found): {rel_path}")
        return False
    with open(abs_path, "r", encoding="utf-8") as f:
        src = f.read()
    if old_text not in src:
        print(f"  SKIP (pattern not found): {rel_path}")
        return False
    src = src.replace(old_text, new_text, 1)
    if old_geom and new_geom:
        src = src.replace(old_geom, new_geom, 1)
    with open(abs_path, "w", encoding="utf-8") as f:
        f.write(src)
    print(f"  PATCHED: {rel_path}")
    return True


IT_FILES = [
    r"Income Tax\Challan Downloader\main.py",
    r"Income Tax\Challan Downloader\demand_checker_app.py",
    r"Income Tax\ITR - Bot\GUI_based_app.py",
    r"Income Tax\26 AS Downlaoder\refund_checker_app.py",
]

if __name__ == "__main__":
    print("Patching IT load_id_pass dialogs...\n")
    for fp in IT_FILES:
        patch_file(
            fp,
            IT_OLD,
            IT_NEW,
            'dialog.geometry("400x460")',
            'dialog.geometry("440x580")',
        )
    print("\nDone!")
