
from bank_to_excel import process_pdf
from openpyxl import load_workbook

# Just extract first 1 page to avoid hanging
import pdfplumber
def custom_extract(pdf_path):
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[:1]
        t = pages[0].extract_tables()
        return t
print('Done')
