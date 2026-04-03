import traceback
import bank_to_excel

try:
    res = bank_to_excel.process_pdf(r'Statements/bank of india.pdf', 'bank of india.pdf', False, 'tesseract', None)
    print(res[1].get('error'))
except Exception as e:
    traceback.print_exc()
