import re

with open(r'c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\Income Tax\Challan Downloader\main.py', 'r', encoding='utf-8') as f:
    text = f.read()

start_kw = '            # ==========================================================\n            # STEP A: Click "e-File"'
end_kw = 'return ("Success" if downloaded_years else "Failed"), summary'

s_idx = text.find(start_kw)
e_idx = text.find(end_kw) + len(end_kw)

if s_idx == -1 or e_idx == -1:
    print("Not found")
else:
    block = text[s_idx:e_idx]
    
    new_block = """            all_downloaded_years = []
            all_failed_years = []
            
            for current_act in ["Income-tax Act, 2025", "Income-tax Act, 1961"]:
                self.log(f"\\n   ?? [ PROCESSING: {current_act} ]")
                downloaded_years = []
                failed_years = []

"""
    lines = block.split('\n')
    for line in lines:
        if line.strip() == "":
            new_block += "\n"
        elif "downloaded_years = []" in line or "failed_years = []" in line:
            continue
        elif "summary = f\"Downloaded:" in line or "summary += f\" | Failed:" in line or "self.log(f\"   ?? {summary}\")" in line or "return (\"Success\" if downloaded_years else \"Failed\"), summary" in line:
            continue
        elif "if not all_years:" in line:
            new_block += "                if not all_years:\n"
            new_block += "                    self.log(\"   ?? No challan records found for this act.\")\n"
            new_block += "                    continue\n"
        elif "return \"Success\", \"No Challan Records Found\"" in line:
            continue
        elif "act_1961_radio = driver.find_element(By.XPATH, \"//div[contains(text(), 'Income-tax Act, 1961')]/ancestor::label\")" in line:
            new_block += line.replace("'Income-tax Act, 1961'", f"{{current_act}}").replace("act_1961_radio", "act_radio") + "\n"
            new_block = new_block.replace("driver.execute_script(\"arguments[0].click();\", act_1961_radio)", "driver.execute_script(\"arguments[0].click();\", act_radio)")
        elif "return \"Failed\", " in line:
            # We don't want to completely fail and stop if just one act fails
            # But let's just keep returns as they are for severe navigation errors, or try continuing
            new_block += "    " + line + "\n"
        else:
            if line.startswith("            "):
                new_block += "    " + line + "\n"
            else:
                new_block += line + "\n"
    
    new_block += """
                all_downloaded_years.extend(downloaded_years)
                all_failed_years.extend(failed_years)
                
            summary = f"Downloaded: {', '.join(all_downloaded_years) or 'None'}"
            if all_failed_years:
                summary += f" | Failed: {', '.join(all_failed_years)}"
            self.log(f"   ?? {summary}")
            return ("Success" if all_downloaded_years else ("Success" if not all_failed_years else "Failed")), summary"""
            
    text = text[:s_idx] + new_block + text[e_idx:]
    with open(r'c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\Income Tax\Challan Downloader\main.py', 'w', encoding='utf-8') as f:
        f.write(text)
    print("Rewritten")
