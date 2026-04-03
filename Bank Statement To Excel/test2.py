import traceback
import bank_to_excel

try:
    print("Starting processing...")
    res = bank_to_excel.process_pdf(r'Statements/bank of india.pdf', 'bank of india.pdf', False, 'tesseract', None)
    print("Done processing, error string is:")
    print(res[1].get('error'))
except Exception as e:
    print("Exception caught:")
    traceback.print_exc()
