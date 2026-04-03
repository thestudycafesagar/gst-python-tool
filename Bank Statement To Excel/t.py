import traceback, bank_to_excel, pandas as pd
try:
    df, info = bank_to_excel.process_pdf('Statements/hdfc fy 25-26.pdf', 'h_final5.xlsx', False, 'tesseract')
    df2 = pd.read_excel('h_final5.xlsx', skiprows=8)
    print('Cols:', df2.columns.tolist()[:10])
    print('Metadata:', info.get('account_metadata'))
except Exception:
    traceback.print_exc()
