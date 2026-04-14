with open("Gmail-Tools/main.py", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace('f"\\n[DONE', 'f"\\\\n[DONE')
text = text.replace('"\\n[DONE', '"\\\\n[DONE')

# But the error showed literal newline
text = text.replace('log_cb(f"\\n', 'log_cb(f"\\\\n')

with open("Gmail-Tools/main.py", "w", encoding="utf-8") as f:
    f.write(text)
