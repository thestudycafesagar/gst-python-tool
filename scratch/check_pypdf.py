import traceback
try:
    from pypdf import PdfReader, PdfWriter
    print("SUCCESS: pypdf imported successfully")
except Exception:
    print("FAILURE: pypdf import failed")
    print(traceback.format_exc())
