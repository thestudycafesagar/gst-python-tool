import os
import ast
import sys

def get_imports(filepath):
    imports = set()
    try:
        with open(filepath, 'rb') as f:
            code = f.read().decode('utf-8', 'ignore')
        tree = ast.parse(code)
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    imports.add(alias.name.split('.')[0])
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    imports.add(node.module.split('.')[0])
    except Exception:
        pass
    return imports

all_imports = set()
for root, dirs, files in os.walk('.'):
    if any(x in root for x in ('.venv', '__pycache__', 'build', 'output', 'GST_Downloads', 'IMS_Downloads')):
        continue
    for file in files:
        if file.endswith('.py'):
            all_imports.update(get_imports(os.path.join(root, file)))

stdlib = set(sys.stdlib_module_names) if hasattr(sys, 'stdlib_module_names') else set()
third_party_candidates = {m for m in all_imports if m not in stdlib and not m.startswith('_')}

print(sorted(list(third_party_candidates)))
