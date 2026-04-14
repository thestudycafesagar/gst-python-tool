with open("Gmail-Tools/main.py", "r", encoding="utf-8") as f:
    t = f.read()

t = t.replace('log_cb(f"\\n[DONE', 'log_cb(f"\\\\n[DONE')
t = t.replace('log_cb("\\n[DONE', 'log_cb("\\\\n[DONE')

# Let's just fix any instances where "\n" became actual newline immediately before "[DONE"
t = t.replace('log_cb(f"\\n\\n', 'log_cb(f"\\\\n\\\\n') 
t = t.replace('f"\\n', 'f"\\\\n')

with open("Gmail-Tools/main.py", "w", encoding="utf-8") as f:
    f.write(t)
