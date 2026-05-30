import base64, re
with open(r'C:\Users\HP\Downloads\Sample File (1).xlsx', 'rb') as f:
    new_b64 = base64.b64encode(f.read()).decode('utf-8')

lines = []
for i in range(0, len(new_b64), 80):
    lines.append('        "' + new_b64[i:i+80] + '"')
wrapped_b64 = '\n'.join(lines)

path = r'C:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\GST_RECO\mainpy-reco-speqtra.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

new_content = re.sub(r'(_SAMPLE_B64\s*=\s*\(\n).*?(\n\s*\))', r'\g<1>' + wrapped_b64.replace('\\', '\\\\') + r'\g<2>', content, flags=re.DOTALL)

with open(path, 'w', encoding='utf-8') as f:
    f.write(new_content)

print('Success!')
