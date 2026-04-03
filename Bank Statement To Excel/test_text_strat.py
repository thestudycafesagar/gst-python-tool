import pdfplumber, pandas as pd; import bank_to_excel
p = pdfplumber.open(r'Statements/hdfc fy 25-26.pdf')
t = p.pages[0].extract_table(table_settings={'vertical_strategy':'text', 'horizontal_strategy':'text'})
df1 = pd.DataFrame(t[1:], columns=t[0])
ext = bank_to_excel.TransactionExtractor(bank_key='HDFC')
res = ext.extract([df1], [''])
print(res.head(5).to_dict('records'))
