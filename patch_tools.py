import os
import re

gst_dir = r"c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\GST"

for root, dirs, files in os.walk(gst_dir):
    for file in files:
        if file in ("main.py", "mai.py"):
            filepath = os.path.join(root, file)
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()

            modified = False

            # 1. Update load_id_pass query
            if "SELECT username, password FROM gst_profiles ORDER BY username" in content:
                content = content.replace(
                    'rows = conn.execute("SELECT username, password FROM gst_profiles ORDER BY username").fetchall()',
                    'cur = conn.cursor()\n            cur.execute("SELECT * FROM gst_profiles ORDER BY username")\n            cols = [d[0] for d in cur.description]\n            rows = [dict(zip(cols, r)) for r in cur.fetchall()]'
                )
                modified = True

            # 2. Update loop over rows
            old_loop = '''        for u, p in rows:
            v = ctk.BooleanVar()
            ctk.CTkCheckBox(scroll, text=u, variable=v).pack(anchor="w", padx=10, pady=3)
            vars_[(u, p)] = v'''
            new_loop = '''        for rdata in rows:
            u = rdata.get("username", "")
            p = rdata.get("password", "")
            c = rdata.get("client_name") or ""
            v = ctk.BooleanVar()
            disp = f"{c} ({u})" if c else u
            ctk.CTkCheckBox(scroll, text=disp, variable=v).pack(anchor="w", padx=10, pady=3)
            vars_[(u, p, c)] = v'''
            if old_loop in content:
                content = content.replace(old_loop, new_loop)
                modified = True

            # 3. Update selected comprehension
            old_sel = 'selected = [{"Username": u, "Password": p} for (u, p), v in vars_.items() if v.get()]'
            new_sel = 'selected = [{"Username": u, "Password": p, "ClientName": c} for (u, p, c), v in vars_.items() if v.get()]'
            if old_sel in content:
                content = content.replace(old_sel, new_sel)
                modified = True
                
            # 4. Update manual_credentials fallback in some places
            # If there's another place creating {"Username": username, "Password": password}
            old_add = 'self.manual_credentials = [{"Username": username, "Password": password}]'
            new_add = 'self.manual_credentials = [{"Username": username, "Password": password, "ClientName": ""}]'
            if old_add in content:
                content = content.replace(old_add, new_add)
                modified = True

            # 5. Update user_root_base
            old_root = 'user_root_base = os.path.join(base_dir, username)'
            new_root = '''cname = str(row.get("ClientName", row.get("Client Name", ""))).strip()
                if cname and str(cname).lower() != "nan":
                    folder_name = cname
                else:
                    folder_name = username
                
                # Replace invalid path characters
                import re as _re_tmp
                folder_name = _re_tmp.sub(r'[\\\\/*?:"<>|]', "", folder_name).strip()
                
                user_root_base = os.path.join(base_dir, folder_name)'''
            if old_root in content:
                content = content.replace(old_root, new_root)
                modified = True

            if modified:
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(content)
                print(f"Patched {filepath}")

print("Done")
