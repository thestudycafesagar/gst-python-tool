import traceback, bank_to_excel
try:
    df, info = bank_to_excel.process_pdf('Statements/hdfc fy 25-26.pdf', 'h_final.xlsx', False, 'tesseract')
    print('success, rows:', info.get('rows_extracted'), info.get('error'))
except Exception:
    traceback.print_exc()
