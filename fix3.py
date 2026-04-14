with open("GST_Suite.py", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace('"tk": True', '"tk": False')

with open("GST_Suite.py", "w", encoding="utf-8") as f:
    f.write(text)
