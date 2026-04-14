import re

with open("GST_Suite.py", "r", encoding="utf-8") as f:
    text = f.read()

# Fix the broken array
text = text.replace(
    'Payment Reminder."}, "is_card_only": True, "action_cat": "email",',
    'Payment Reminder.", "is_card_only": True, "action_cat": "email"},'
)

text = text.replace(
    'Payment Reminder."}, "is_card_only": True, "action_cat": "gmail",',
    'Payment Reminder.", "is_card_only": True, "action_cat": "gmail"},'
)


with open("GST_Suite.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Patch syntax fix complete")
