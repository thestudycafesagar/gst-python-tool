with open("Gmail-Tools/main.py", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace('log_cb(f"\\n[DONE] Finished', 'log_cb(f"\\\\n[DONE] Finished')
text = text.replace('log_cb("\\n[DONE] All emails', 'log_cb("\\\\n[DONE] All emails')
text = text.replace('log_cb(f"\\n', 'log_cb(f"\\\\n')

with open("Gmail-Tools/main.py", "w", encoding="utf-8") as f:
    f.write(text.replace('\n[DONE]', '\\n[DONE]'))
