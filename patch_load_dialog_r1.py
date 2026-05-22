import os
import re

f = r"GST\GST R1 Downloader\mai.py"

pattern = re.compile(
    r'(# .*Search box.*?search_entry\.pack\(fill="x", padx=16, pady=\(0, 6\)\)).*?(def _load\(\):.*?self\.manual_credentials = selected)',
    re.DOTALL
)

replacement = r'''\1

        # Scrollable profile list
        scroll = ctk.CTkScrollableFrame(dialog, height=260)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 6))
        
        selected_var = ctk.StringVar(value="")
        data_map = {}
        widgets_ = {}

        for i, rdata in enumerate(rows):
            u = rdata.get("username", "")
            p = rdata.get("password", "")
            c = rdata.get("client_name") or ""
            f_freq = rdata.get("filing_frequency") or "Monthly"
            
            disp = f"{c} ({u})" if c else u
            uid = f"prof_{i}"
            data_map[uid] = (u, p, c, f_freq)
            
            chk = ctk.CTkRadioButton(scroll, text=disp, variable=selected_var, value=uid)
            chk.pack(anchor="w", padx=10, pady=5)
            widgets_[uid] = (chk, disp)

        def _on_search(*_):
            q = search_var.get().strip().lower()
            for key, (chk, disp) in widgets_.items():
                if q == "" or q in disp.lower():
                    chk.pack(anchor="w", padx=10, pady=5)
                else:
                    chk.pack_forget()

        search_var.trace_add("write", _on_search)
        search_entry.focus_set()

        def _load():
            uid = selected_var.get()
            if not uid or uid not in data_map:
                messagebox.showwarning("No Selection", "Please select a profile.", parent=dialog)
                return
            
            u, p, c, f_freq = data_map[uid]
            selected = [{"Username": u, "Password": p, "ClientName": c, "FilingFrequency": f_freq}]
            
            self.manual_credentials = selected'''

with open(f, "r", encoding="utf-8") as file:
    content = file.read()

match = pattern.search(content)
if match:
    c1 = pattern.sub(replacement, content)
    with open(f, "w", encoding="utf-8") as file:
        file.write(c1)
    print(f"Patched {f}")
else:
    print(f"Target not found in {f}")

