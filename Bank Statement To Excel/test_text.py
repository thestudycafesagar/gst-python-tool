import bank_to_excel
processor = bank_to_excel.PDFProcessor(False, 'tesseract')
texts = processor.extract_pages_text(r'Statements/hdfc fy 25-26.pdf')
ext = bank_to_excel.TransactionExtractor(bank_key='HDFC')
df = ext._extract_from_text(texts)
print(df.head(2).to_dict('records'))
