import os
import re

files = [
    r"GST\GST 2B Downloader\main.py",
    r"GST\GST 3B Downloader\main.py",
    r"GST\GST Challan Downloader\main.py",
    r"GST\GST R1 Downloader\mai.py",
    r"GST\R1 PDF Downloader\main.py"
]

toggle_target = '''            if mode == "Monthly":
                items = ["April", "May", "June", "July", "August", "September",
                         "October", "November", "December", "January", "February", "March"]
                cols = 3
            else:
                items = ["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)",
                         "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"]
                cols = 2'''

toggle_replace = '''            if mode == "Monthly":
                items = ["Apr", "May", "Jun", "Jul", "Aug", "Sep",
                         "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
                cols = 6
            else:
                items = ["Q1 (Apr-Jun)", "Q2 (Jul-Sep)",
                         "Q3 (Oct-Dec)", "Q4 (Jan-Mar)"]
                cols = 4'''

start_process_target = '''        tasks = []
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
                tasks.append({"q": q, "m": q_last_month.get(q, "")})'''

start_process_replace = '''        tasks = []
        ui_m = {"Apr": "April", "May": "May", "Jun": "June", "Jul": "July", "Aug": "August", "Sep": "September",
                "Oct": "October", "Nov": "November", "Dec": "December", "Jan": "January", "Feb": "February", "Mar": "March"}
        ui_q = {"Q1 (Apr-Jun)": "Quarter 1 (Apr - Jun)", "Q2 (Jul-Sep)": "Quarter 2 (Jul - Sep)", 
                "Q3 (Oct-Dec)": "Quarter 3 (Oct - Dec)", "Q4 (Jan-Mar)": "Quarter 4 (Jan - Mar)"}
        
        q_map_rev = {
            "April": "Quarter 1 (Apr - Jun)", "May": "Quarter 1 (Apr - Jun)", "June": "Quarter 1 (Apr - Jun)",
            "July": "Quarter 2 (Jul - Sep)", "August": "Quarter 2 (Jul - Sep)", "September": "Quarter 2 (Jul - Sep)",
            "October": "Quarter 3 (Oct - Dec)", "November": "Quarter 3 (Oct - Dec)", "December": "Quarter 3 (Oct - Dec)",
            "January": "Quarter 4 (Jan - Mar)", "February": "Quarter 4 (Jan - Mar)", "March": "Quarter 4 (Jan - Mar)"
        }
        q_last_month = {
            "Quarter 1 (Apr - Jun)": "June", "Quarter 2 (Jul - Sep)": "September",
            "Quarter 3 (Oct - Dec)": "December", "Quarter 4 (Jan - Mar)": "March"
        }

        if mode == "Monthly":
            for m in selected_periods:
                full_m = ui_m.get(m, m)
                tasks.append({"q": q_map_rev.get(full_m, ""), "m": full_m})
        else:
            for q in selected_periods:
                full_q = ui_q.get(q, q)
                tasks.append({"q": full_q, "m": q_last_month.get(full_q, "")})'''


for f in files:
    if not os.path.exists(f): continue
    with open(f, "r", encoding="utf-8") as file:
        content = file.read()
    
    if toggle_target in content:
        c1 = content.replace(toggle_target, toggle_replace)
    else:
        print(f"toggle_target not found in {f}")
        c1 = content
        
    if start_process_target in c1:
        c2 = c1.replace(start_process_target, start_process_replace)
    else:
        print(f"start_process_target not found in {f}")
        c2 = c1
    
    with open(f, "w", encoding="utf-8") as file:
        file.write(c2)
    print(f"Patched {f}")

