
import os
from pathlib import Path
from bank_to_excel import process_pdf, DataCleaner, export_to_excel_bytes

folder = Path('Statements')
results = {}
for file in folder.glob('*.pdf'):
    print(f'Processing {file.name}...')
    df, info = process_pdf(str(file), file.name, False, 'pytesseract', None)
    if df is not None and not df.empty:
        dc = DataCleaner()
        cleaned = dc.clean(df)
        results[file.name] = cleaned
        print(f'  -> Found cols: {list(cleaned.columns)}')

if results:
    bio = export_to_excel_bytes(results)
    with open('test_output.xlsx', 'wb') as f:
        f.write(bio)
    print('Saved test_output.xlsx')
else:
    print('No data extracted')

