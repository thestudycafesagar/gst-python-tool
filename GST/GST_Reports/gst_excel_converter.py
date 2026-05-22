import json
from pathlib import Path
from gst_excel_utils import XLSX_OK

# Import specific converters
from gstr1_excel import gstr1_to_excel
from gstr2a_excel import gstr2a_to_excel
from gstr2b_excel import gstr2b_to_excel
from gstr3b_excel import gstr3b_to_excel

# =============================================================================
# Public dispatcher
# =============================================================================

def convert_to_excel(form: str, json_path: str, excel_path: str):
    """
    Load JSON from json_path and write formatted Excel to excel_path.

    Args:
        form:       One of "GSTR-1", "GSTR-2A", "GSTR-2B", "GSTR-3B"
        json_path:  Path to the source .json file
        excel_path: Destination .xlsx path (parent dir created automatically)

    Raises:
        ImportError  if openpyxl is not installed
        ValueError   if form is unknown
    """
    if not XLSX_OK:
        raise ImportError("openpyxl is required.  Run:  pip install openpyxl")

    with open(json_path, encoding="utf-8") as fh:
        data = json.load(fh)

    dispatchers = {
        "GSTR-1":  gstr1_to_excel,
        "GSTR-2A": gstr2a_to_excel,
        "GSTR-2B": gstr2b_to_excel,
        "GSTR-3B": gstr3b_to_excel,
    }
    fn = dispatchers.get(form)
    if fn is None:
        raise ValueError(f"Unknown form '{form}'. Expected one of: {list(dispatchers)}")

    fn(data, excel_path)
