import traceback; import bank_to_excel; res = bank_to_excel.process_pdf(r'Statements/bank of india.pdf', 'bank of india.pdf', False, 'tesseract', None); print(res[1]) 
