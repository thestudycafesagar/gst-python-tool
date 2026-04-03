import os
import re
import sys

import_re = re.compile(r'^\s*(?:from\s+([a-zA-Z0-9_\.]+)\s+import|import\s+([a-zA-Z0-9_\.,\s]+))', re.MULTILINE)

all_imports = set()
for root, dirs, files in os.walk('.'):
    if any(x in root for x in ('.venv', '__pycache__', 'build', 'output', 'GST_Downloads', 'IMS_Downloads')):
        continue
    for file in files:
        if file.endswith('.py'):
            try:
                with open(os.path.join(root, file), 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    matches = import_re.findall(content)
                    for m in matches:
                        # m is a tuple: (from_module, import_modules)
                        if m[0]:
                            all_imports.add(m[0].split('.')[0])
                        if m[1]:
                            for var in m[1].split(','):
                                all_imports.add(var.strip().split('.')[0])
            except Exception as e:
                pass

stdlib = set(sys.stdlib_module_names) if hasattr(sys, 'stdlib_module_names') else set()

# Also get current directory folders to exclude local modules
local_modules = {d for d in os.listdir('.') if os.path.isdir(d)}
local_modules.update({f.replace('.py', '') for f in os.listdir('.') if f.endswith('.py')})
local_modules.add('main')
local_modules.add('app')

third_party = set()
for m in all_imports:
    if m and m not in stdlib and m not in local_modules and not m.startswith('_') and not m.startswith('.'):
        third_party.add(m)

print("THIRD_PARTY_DEPS:")
for d in sorted(list(third_party)):
    print(d)
