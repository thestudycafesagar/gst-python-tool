import pdfplumber, pandas as pd; import bank_to_excel
p = pdfplumber.open(r'Statements/hdfc fy 25-26.pdf')
t = p.pages[0].extract_tables()
if t:
    df1 = pd.DataFrame(t[0][1:], columns=t[0][0])
    ext = bank_to_excel.TransactionExtractor(bank_key='HDFC')
    res = ext.extract([df1], [''])
    print(res.columns.tolist())
    print(res.head(2).to_dict('records'))

    print(ext._map_columns(df1.columns.tolist()))
    print('D:', ext.col_display_names)
    print([len(str(x).split('\n')) for x in t[0][1]])