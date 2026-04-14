import re

with open("GST_Suite.py", "r", encoding="utf-8") as f:
    text = f.read()

# 1. Update MAIL_GROUP_TOOLS
text = re.sub(
    r'(\{"key": "Email_Suite".*?Payment Reminder\."\})(,?)',
    r'\1, "is_card_only": True, "action_cat": "email"\2',
    text
)
text = re.sub(
    r'(\{"key": "Gmail_Suite".*?Payment Reminder\."\})(,?)',
    r'\1, "is_card_only": True, "action_cat": "gmail"\2',
    text
)

# 2. Update _get_or_build_category loop
search_loop = """        for t in tools:
            tv.add(t["tab"])
            if not self._is_tool_allowed(t.get("key")):
                self._loaded[t["tab"]] = True   # skip lazy-loader"""
replace_loop = """        for t in tools:
            if t.get("is_card_only"):
                continue
            tv.add(t["tab"])
            if not self._is_tool_allowed(t.get("key")):
                self._loaded[t["tab"]] = True   # skip lazy-loader"""
text = text.replace(search_loop, replace_loop)

# 3. Update _build_category_overview click attachment
search_click = """            if tv is not None and not is_locked:
                def _make_attach(tab_name=tool["tab"], _card=card):
                    def _click(_=None): tv.set(tab_name)
                    def _enter(_=None):"""
replace_click = """            if not is_locked:
                action_cat = tool.get("action_cat")
                def _make_attach(tab_name=tool["tab"], _card=card, _action_cat=action_cat):
                    def _click(_=None):
                        if _action_cat:
                            self._show_category(_action_cat)
                        elif tv is not None:
                            tv.set(tab_name)
                    def _enter(_=None):"""
text = text.replace(search_click, replace_click)

with open("GST_Suite.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Patch complete")
