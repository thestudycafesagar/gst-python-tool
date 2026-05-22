import json
from pathlib import Path
from gst_excel_utils import (
    openpyxl, _write_sheet, _empty_sheet, _b2b_rows, _cdn_rows, 
    _B2B_NUM_COLS, _CDN_NUM_COLS, _s, _n, _state_name, _doc_name,
    Font, PatternFill, Alignment, Border, Side, get_column_letter
)

_MONTHS = {
    "01": "January", "02": "February", "03": "March", "04": "April",
    "05": "May", "06": "June", "07": "July", "08": "August",
    "09": "September", "10": "October", "11": "November", "12": "December"
}
_MON_ABB = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
    "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
    "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
}

def _get_month(fp):
    if not fp: return ""
    m = str(fp)[:2]
    return _MONTHS.get(m, fp)

def _nf(rec, *keys):
    """Try multiple keys for a numeric value."""
    if not isinstance(rec, dict): return 0.0
    for k in keys:
        v = rec.get(k)
        if v is not None:
            return _n(v)
    return 0.0

def _orig_month(idt: str) -> str:
    """'DD-MM-YYYY' or 'DD-Mon-YYYY' → 'Apr 2025'"""
    if not idt:
        return ""
    parts = idt.replace("/", "-").split("-")
    if len(parts) != 3:
        return ""
    _, mm, yyyy = parts
    if not yyyy.isdigit():
        return ""
    if mm.isdigit():
        return f"{_MON_ABB.get(mm.zfill(2), mm)} {yyyy}"
    # Abbreviated month name like 'Apr'
    return f"{mm[:3].capitalize()} {yyyy}"

def _showing_month(fp: str) -> str:
    """fp='042025' → 'Apr 2025'"""
    if not fp or len(fp) < 6:
        return ""
    mm, yyyy = fp[:2], fp[2:]
    return f"{_MON_ABB.get(mm, mm)} {yyyy}"

def _write_portal_sheet(ws, headers: list, rows: list, num_col_indices: set = None, 
                       profile: dict = None, report_name: str = "", fy: str = ""):
    """Write sheet in Portal/Speqta layout with metadata at the top."""
    # Metadata rows
    trdnm = (profile.get("lgl_nm") or profile.get("trdnm") or profile.get("bname") or "").upper()
    gstin = (profile.get("gstin") or "").upper()
    
    # Row 1-4: Metadata
    meta = [
        ("Trade Name::    ", trdnm),
        ("GSTIN::    ", gstin),
        ("Year::    ", fy),
        ("Report Name", report_name)
    ]
    
    bold_calibri = Font(bold=True, size=11, name="Calibri")
    for i, (label, val) in enumerate(meta, 1):
        ws.cell(row=i, column=3, value=label).font = bold_calibri
        ws.cell(row=i, column=4, value=val).font = bold_calibri

    # Row 5: Headers
    hdr_fill = PatternFill("solid", fgColor="002060") # Dark Blue
    hdr_font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    bdr = Border(left=Side(style="thin"), right=Side(style="thin"), 
                 top=Side(style="thin"), bottom=Side(style="thin"))
    
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = bdr
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data rows
    body_font = Font(size=10, name="Calibri")
    num_cols = num_col_indices or set()
    for r, row in enumerate(rows, 6):
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = body_font
            cell.border = bdr
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.00"
            elif c in num_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.00"
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Column widths
    for c, h in enumerate(headers, 1):
        col_letter = get_column_letter(c)
        ws.column_dimensions[col_letter].width = max(12, len(str(h)) + 2)

    ws.freeze_panes = "A6"

_INV_TYP_MAP = {
    "R": "Regular", "S": "SEZ (with Payment)", "SEWP": "SEZ (with Payment)",
    "SEWOP": "SEZ (without Payment)", "DE": "Deemed Exports", "CB": "SEZ (without Payment)",
    "": "Regular",
}

def _map_inv_typ(raw: str) -> str:
    return _INV_TYP_MAP.get(str(raw).strip(), raw) if raw else "Regular"

def _dedup_by_period(data_list: list) -> list:
    """
    Merge/deduplicate data objects for the same GST period.
    When both a ZIP file and an API-sections file exist for the same period,
    merge them: ZIP provides individual invoice data, API sections fill in
    any sections missing from the ZIP (b2cl, exp, cdnr, etc.) plus the summary.
    """
    seen = {}
    for d in data_list:
        fp = str(d.get("fp") or d.get("rtn_prd") or "")
        if not fp:
            continue
        existing = seen.get(fp)
        if existing is None:
            seen[fp] = d
        else:
            is_zip  = "sections" not in d      and ("b2b" in d or "b2cl" in d or "gstin" in d)
            ex_zip  = "sections" not in existing and ("b2b" in existing or "b2cl" in existing or "gstin" in existing)

            if is_zip and not ex_zip:
                # new=ZIP, existing=API → merge: ZIP base + API sections for gaps
                merged = dict(d)
                merged["sections"] = existing.get("sections", {})
                merged["summary"]  = merged.get("summary") or existing.get("summary", "")
                merged["profile"]  = merged.get("profile") or existing.get("profile", {})
                seen[fp] = merged
            elif not is_zip and ex_zip:
                # new=API, existing=ZIP → patch existing with API sections
                merged = dict(existing)
                merged["sections"] = d.get("sections", {})
                merged["summary"]  = merged.get("summary") or d.get("summary", "")
                merged["profile"]  = merged.get("profile") or d.get("profile", {})
                seen[fp] = merged
            elif is_zip and ex_zip:
                pass  # two ZIPs — keep existing
            # else: two API files — keep existing
    return list(seen.values()) if seen else data_list

def gstr1_consolidated_to_excel(data_list: list, out_path: str, profile: dict = None):
    """Combine multiple months of GSTR-1 data into a single consolidated Excel matching Portal layout."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    if not data_list:
        _empty_sheet(wb)
        wb.save(out_path)
        return

    # Remove duplicate periods (e.g. 2025_03.json and 2026_03.json both for March 2026)
    data_list = _dedup_by_period(data_list)

    # Sort in FY order: April(04)→May→...→March(03)
    def _fy_sort_key(d):
        fp = str(d.get("fp") or d.get("rtn_prd") or "")
        if len(fp) < 6: return 99
        m = int(fp[:2]) if fp[:2].isdigit() else 99
        return m - 4 if m >= 4 else m + 8   # Apr=0, May=1, ..., Dec=8, Jan=9, Feb=10, Mar=11
    data_list = sorted(data_list, key=_fy_sort_key)

    # Diagnostic: Print the periods being consolidated
    print(f"DEBUG: Consolidating {len(data_list)} periods:")
    for d in data_list:
        print(f"  - FP: {d.get('fp') or d.get('rtn_prd')} | Source: {d.get('source')} | B2B Count: {d.get('b2b_count')}")

    # Build a global GSTIN -> Trade Name map from all summaries and cache
    name_map = {}
    # 1. Try local cache if available
    try:
        cfg_path = Path(__file__).parent / "app_config.json"
        if cfg_path.exists():
            with open(cfg_path, encoding="utf-8") as f:
                cfg = json.load(f)
                name_map.update(cfg.get("gstin_name_cache", {}))
    except: pass

    # 2. Harvest names from all available dashboard summaries in the data_list
    for data in data_list:
        summary_str = data.get("summary", "")
        if not summary_str: continue
        try:
            sum_obj = json.loads(summary_str) if isinstance(summary_str, str) else summary_str
            
            def _scan_for_names(obj_list):
                if not obj_list: return
                for sec in obj_list:
                    # Check top-level cpty_sum
                    for party in sec.get("cpty_sum", []):
                        ctin = party.get("ctin")
                        name = party.get("trdnm") or party.get("trade_name") or party.get("trdNm") or party.get("trade_nm")
                        if ctin and name: name_map[ctin] = name
                    # Recursively check sub_sections
                    if "sub_sections" in sec:
                        _scan_for_names(sec["sub_sections"])

            sec_sums = sum_obj.get("data", {}).get("sec_sum", [])
            _scan_for_names(sec_sums)
        except: pass

    # Extract global metadata from first data that has profile/gstin
    prof = profile or {}
    for d in data_list:
        if prof:
            break
        prof = d.get("profile") or {}
        if not prof:
            prof = {
                "lgl_nm": d.get("bname") or d.get("trdnm") or d.get("lgl_nm"),
                "gstin":  d.get("gstin") or d.get("ctin"),
            }
            if not any(prof.values()):
                prof = {}

    fy_str = ""
    if data_list:
        fp_list = sorted(str(d.get("fp") or d.get("rtn_prd") or "") for d in data_list if d.get("fp") or d.get("rtn_prd"))
        if fp_list:
            # fp format: MMYYYY → extract FY start year from April (month=04)
            years_seen = sorted({int(fp[-4:]) for fp in fp_list if len(fp) >= 6 and fp[-4:].isdigit()})
            months_seen = [int(fp[:2]) for fp in fp_list if len(fp) >= 6 and fp[:2].isdigit()]
            if years_seen:
                # FY start = earliest year that has an April or later month
                fy_start = min(y for y in years_seen if any(m >= 4 for m in months_seen) or y)
                fy_str = f"{fy_start} - {str(fy_start + 1)[2:]}"

    _default_headers = {
        "b2b":    ["Sr.No.", "Original Month", "Showing in Month", "Category", "GSTIN", "Invoice Date", "Invoice No.", "Customer Name", "Invoice Value", "Taxable Value", "GST %", "IGST", "CGST", "SGST", "Cess", "Place Of Supply", "RCM Applicable", "E-Commerce GSTIN", "Return Type"],
        "b2cl":   ["Month", "Customer Name", "No. of Invoice", "Invoice Date", "Taxable Value", "GST %", "IGST", "Cess", "Invoice Value", "Place Of Supply", "E-Commerce GSTIN", "Return Type"],
        "b2cs":   ["Month", "Taxable Value", "GST %", "IGST", "CGST", "SGST", "Cess", "Place Of Supply", "E-Commerce GSTIN", "Supply Type", "Return Type"],
        "cdnr":   ["Month", "GSTIN", "Customer Name", "Type of Note    (Cr. / Dr.)", "RCM", "Cr. / Dr. Note No.", "Cr. / Dr. Note Date", "Invoice Type", "GST %", "Taxable Value", "IGST", "CGST", "SGST ", "CESS", "Note/Refund Voucher Value", "Place of supply", "Return Type"],
        "cdnur":  ["Month", "Customer Name", "Supply Type", "Type of note (Debit/ Credit)", "Pre GST Regime Dr./ Cr. Notes", "Debit Note/ credit note/ Refund voucher No.", "Debit Note/ credit note/ Refund voucher Date", "Original Invoice No", "Original Invoice Date", "GST %", "Taxable Value", "IGST", "CGST", "SGST", "Cess", "Note/Refund Voucher Value", "Place of supply", "Return Type"],
        "exp":    ["Month", "Type Of Export", "Customer Name", "Invoice No.", "Invoice Date", "Invoice Value", " Port Code", "Shipping bill/ bill of export No", "Shipping bill/ bill of export Date", "GST %", "Taxable Value", "IGST", "Cess", "Return Type"],
        "nil":    ["Month", "Type", "Value", "Return Type", "Nil rated", "Exempted", "Non-GST Supply"],
        "at":     ["Month", "Gross Advance Received", "GST %", "IGST", "CGST", "SGST", "Cess", "Place of Supply", "Return Type"],
        "txp":    ["Month", "Advance adjusted", "Rate", "IGST", "CGST", "SGST", "Cess", "Place Of Supply", "Return Type"],
        "hsnsum": ["Month", "HSN", "Record Type", "Description", "UQC", "Total Quantity", "Taxable Value", "GST %", "IGST", "CGST", "SGST", "Cess", "Total Value", "Return Type"],
        "dociss": ["Nature of Document", "Sr. No. From", "Sr. No. To", "Total Number", "Cancelled", "Net Issued", "Return Type"],
        "ecom":   ["Month", "GSTIN of e-Commerce Operator", "Trade Name", "ECO Liable to Pay Tax \nu/s 9(5)", "Net Value of Supplies", "IGST", "CGST", "SGST", "Cess"],
    }

    sections_to_process = [
        ("b2b",    "B2B Invoices",      "GSTR-1 - B2B Invoices"),
        ("b2cl",   "B2C Large",         "GSTR-1 - B2C Large"),
        ("b2cs",   "B2C Small",         "GSTR-1 - B2C Small"),
        ("cdnr",   "CDN",               "GSTR-1 - CDN Invoices"),
        ("cdnur",  "CDNUR",             "GSTR-1 - CDNUR Invoices"),
        ("exp",    "Export",            "GSTR-1 - Export Invoices"),
        ("nil",    "Nil-Exempt-NonGST", "GSTR-1 - Nil-Exempt-NonGST"),
        ("at",     "Adv Received",      "GSTR-1 - Advance Received"),
        ("txp",    "Adv Adjtd",         "GSTR-1 - Advance Adjusted"),
        ("hsnsum", "HSN",               "GSTR-1 - HSN Summary"),
        ("dociss", "Document Summary",  "GSTR-1 - Document Summary"),
        ("ecom",   "E-Com Tab-14",      "GSTR-1 - Table 14-Ecom")
    ]

    def _get_sec(data, key):
        if not isinstance(data, dict): return []
        secs = data.get("sections") or data
        if not isinstance(secs, dict): return []
        
        alt_keys = [key, key.lower(), key.upper()]
        if key == "hsnsum": alt_keys.extend(["hsn", "HSN", "hsnsum", "HSNSUM"])
        if key == "dociss": alt_keys.extend(["doc_issue", "DOC_ISSUE", "doc_det", "dociss", "DOCISS"])
        if key == "b2b": alt_keys.extend(["b2ba", "B2BA"])
        if key == "cdnr": alt_keys.extend(["cdnra", "CDNRA"])
        if key == "cdnur": alt_keys.extend(["cdnura", "CDNURA"])

        val = None
        for k in alt_keys:
            val = data.get(k) or secs.get(k)
            if val is not None: break

        if val is None: return []
        
        # If val is a string, it might be double-encoded JSON
        if isinstance(val, str):
            try: 
                loaded = json.loads(val)
                # If loaded is also a dict with the key inside, unpack it
                if isinstance(loaded, dict):
                    for k in alt_keys + ["data", "inv", "itms"]:
                        if k in loaded:
                            val = loaded[k]
                            break
                    else:
                        val = loaded
                else:
                    val = loaded
            except: 
                return []

        # Reject GST portal error responses
        if isinstance(val, dict) and val.get("status") == 0 and "error" in val:
            return []

        if isinstance(val, list): return val
        if not isinstance(val, dict): return []

        # Look for the data list inside common portal wrapper keys
        for k in [key, key.lower(), key.upper(), "doc_det", "doc_issue", "hsn", "data", "inv", "nt", "itms", "processedInvoice", "cpty"]:
            if k in val and isinstance(val[k], list): return val[k]

        if isinstance(val.get("data"), dict):
            d = val["data"]
            for k in [key, key.lower(), key.upper(), "processedInvoice", "doc_det", "itms", "inv", "cpty", "ctin_list"]:
                if k in d and isinstance(d[k], list): return d[k]
            for k, v in d.items():
                if k.lower().startswith(key.lower()) and isinstance(v, list): return v
        elif isinstance(val.get("data"), list):
            return val["data"]

        # Only return val as a single-item list if it has meaningful content
        if val and any(isinstance(v, (list, dict, int, float, str)) and v != "" for v in val.values()):
            return [val]
        return []

    for key, sheet_name, report_title in sections_to_process:
        all_rows = []
        headers = []
        num_cols = set()

        for data in data_list:
            # Try multiple keys to find the period/month
            fp_raw = data.get("fp") or data.get("rtn_prd") or data.get("period") or ""
            month = _get_month(fp_raw)
            sec_data = _get_sec(data, key)

            records = sec_data if isinstance(sec_data, list) else ([sec_data] if sec_data else [])

            if key == "b2b":
                # Compugst-style columns: Sr.No., Original Month, Showing in Month, Category,
                # GSTIN, Invoice Date, Invoice No., Customer Name,
                # Invoice Value, Taxable Value, GST %, IGST, CGST, SGST, Cess,
                # Place Of Supply, RCM Applicable, E-Commerce GSTIN, Return Type
                if not headers: headers = ["Sr.No.", "Original Month", "Showing in Month", "Category", "GSTIN", "Invoice Date", "Invoice No.", "Customer Name", "Invoice Value", "Taxable Value", "GST %", "IGST", "CGST", "SGST", "Cess", "Place Of Supply", "RCM Applicable", "E-Commerce GSTIN", "Return Type"]
                num_cols = {9, 10, 11, 12, 13, 14, 15} # Value, Taxable, Rate, IGST, CGST, SGST, Cess
                showing_month = _showing_month(str(fp_raw))
                if records:
                    for party in records:
                        ctin = _s(party.get("ctin") or party.get("stin") or "")
                        name = _s(party.get("trdnm") or party.get("trade_name") or party.get("trdNm") or "")
                        if not name and ctin in name_map:
                            name = name_map[ctin]
                        
                        # Robust invoice discovery within party
                        invoices = []
                        for k in ["inv", "b2b", "processedInvoice", "itms", "nt"]:
                            if k in party and isinstance(party[k], list):
                                invoices = party[k]
                                break
                        if not invoices:
                            invoices = [party] if (party.get("inum") or party.get("nt_num") or party.get("val")) else []
                        # Deduplicate invoices by inum (portal pagination can produce repeats)
                        seen_inv = set()
                        deduped = []
                        for inv in invoices:
                            inv_key = inv.get("inum") or inv.get("nt_num")
                            if inv_key:
                                if inv_key in seen_inv: continue
                                seen_inv.add(inv_key)
                            deduped.append(inv)
                        invoices = deduped
                            
                        for inv in invoices:
                            rcm = _s(inv.get("rchrg", "N"))
                            rcm_label = "Yes" if rcm == "Y" else "No"
                            idt = _s(inv.get("idt"))
                            orig_month = _orig_month(idt)
                            category = _map_inv_typ(inv.get("inv_typ", ""))
                            for itm in (inv.get("itms") or [inv]):
                                d = itm.get("itm_det") or itm
                                # Portal API returns inv-prefixed keys when no itms breakdown
                                txval = _nf(d, "txval", "invtxval")
                                iamt  = _nf(d, "iamt",  "inviamt")
                                camt  = _nf(d, "camt",  "invcamt")
                                samt  = _nf(d, "samt",  "invsamt")
                                csamt = _nf(d, "csamt", "invcsamt")
                                # Derive rate from taxes when not provided
                                rt = _n(d.get("rt"))
                                if not rt and txval:
                                    rt = round((iamt + camt + samt) / txval * 100) if txval else 0
                                srno = len(all_rows) + 1
                                all_rows.append([
                                    srno, orig_month, showing_month, category,
                                    ctin, idt, _s(inv.get("inum") or inv.get("inv_no") or inv.get("inum_") or ""), name,
                                    _n(inv.get("val")), txval, rt,
                                    iamt, camt, samt, csamt,
                                    _state_name(_s(inv.get("pos"))), rcm_label, "", "GSTR-1"
                                ])
                else:
                    # Fallback: extract party-level totals from the dashboard summary
                    summary_str = data.get("summary", "")
                    if summary_str:
                        try:
                            sum_obj = json.loads(summary_str) if isinstance(summary_str, str) else summary_str
                            seen_ctins = set()
                            for sec in sum_obj.get("data", {}).get("sec_sum", []):
                                sn = sec.get("sec_nm", "")
                                if sn == "B2B" or sn.startswith("B2B_"):
                                    for party in sec.get("cpty_sum", []):
                                        ctin = _s(party.get("ctin", ""))
                                        if not ctin or ctin in seen_ctins:
                                            continue
                                        seen_ctins.add(ctin)
                                        name = _s(party.get("trdnm") or party.get("trade_name") or party.get("trdNm") or "")
                                        n = party.get("ttl_rec", 1)
                                        srno = len(all_rows) + 1
                                        all_rows.append([
                                            srno, "", showing_month, "B2B",
                                            ctin, "", f"({n} invs)", name,
                                            _n(party.get("ttl_val")), _n(party.get("ttl_tax")), "",
                                            _n(party.get("ttl_igst")),
                                            _n(party.get("ttl_cgst")),
                                            _n(party.get("ttl_sgst")),
                                            _n(party.get("ttl_cess")),
                                            "", "No", "", "GSTR-1"
                                        ])
                        except Exception:
                            pass

            elif key == "b2cl":
                if not records: continue
                if not headers: headers = ["Month", "Customer Name", "Invoice No.", "Invoice Date", "Taxable Value", "GST %", "IGST", "Cess", "Invoice Value", "Place Of Supply", "E-Commerce GSTIN", "Return Type"]
                num_cols = {4, 5, 6, 7, 8}
                for party in records:
                    invoices = party.get("inv") or ([party] if party.get("inum") else [])
                    name = _s(party.get("trdnm") or party.get("trade_name") or "")
                    for inv in invoices:
                        for itm in (inv.get("itms") or [inv]):
                            d = itm.get("itm_det") or itm
                            txval = _nf(d, "txval", "invtxval")
                            igst  = _nf(d, "iamt",  "inviamt")
                            csamt = _nf(d, "csamt", "invcsamt")
                            rt    = _nf(d, "rt")
                            if not rt and txval:
                                rt = round((igst + _nf(d, "camt","invcamt") + _nf(d, "samt","invsamt")) / txval * 100)
                            all_rows.append([
                                month, name, _s(inv.get("inum") or inv.get("inv_no") or ""), _s(inv.get("idt")),
                                txval, rt, igst, csamt, _n(inv.get("val")),
                                _state_name(_s(inv.get("pos") or party.get("pos"))), None, "GSTR-1"
                            ])

            elif key == "b2cs":
                if not headers: headers = ["Month", "Taxable Value", "GST %", "IGST", "CGST", "SGST", "Cess", "Place Of Supply", "E-Commerce GSTIN", "Supply Type", "Return Type"]
                num_cols = {2, 4, 5, 6, 7}
                if records:
                    for inv in records:
                        all_rows.append([
                            month, _nf(inv, "txval", "invtxval"), _nf(inv, "rt"), _nf(inv, "iamt", "inviamt"),
                            _nf(inv, "camt", "invcamt"), _nf(inv, "samt", "invsamt"), _nf(inv, "csamt", "invcsamt"),
                            _state_name(_s(inv.get("pos"))), "", _s(inv.get("typ", "OE")), "GSTR-1"
                        ])
                else:
                    # Fallback: extract state-wise B2CS totals from the dashboard summary
                    sum_str = data.get("summary", "")
                    if sum_str:
                        try:
                            s = json.loads(sum_str) if isinstance(sum_str, str) else sum_str
                            for sec in s.get("data", {}).get("sec_sum", []):
                                if sec.get("sec_nm") == "B2CS":
                                    for st in sec.get("cpty_sum", []):
                                        txval = _n(st.get("ttl_tax"))
                                        igst  = _n(st.get("ttl_igst"))
                                        cgst  = _n(st.get("ttl_cgst"))
                                        sgst  = _n(st.get("ttl_sgst"))
                                        cess  = _n(st.get("ttl_cess"))
                                        pos   = _state_name(_s(st.get("state_cd", "")))
                                        # Estimate rate from taxes/txval
                                        tax_total = (igst or 0) + (cgst or 0) + (sgst or 0)
                                        rt = round(tax_total / txval * 100) if txval else 18
                                        all_rows.append([
                                            month, txval, rt, igst, cgst, sgst, cess,
                                            pos, "", "OE", "GSTR-1"
                                        ])
                        except Exception:
                            pass

            elif key == "cdnr":
                if not headers: headers = ["Month", "GSTIN", "Customer Name", "Type of Note    (Cr. / Dr.)", "RCM", "Cr. / Dr. Note No.", "Cr. / Dr. Note Date", "Invoice Type", "GST %", "Taxable Value", "IGST", "CGST", "SGST ", "CESS", "Note/Refund Voucher Value", "Place of supply", "Return Type"]
                num_cols = {9, 10, 11, 12, 13, 14, 15}
                if records:
                    for party in records:
                        ctin = _s(party.get("ctin") or "")
                        name = _s(party.get("trdnm") or party.get("trade_name") or party.get("trdNm") or "")
                        if not name and ctin in name_map:
                            name = name_map.get(ctin, "")
                        notes = party.get("nt") or party.get("inv") or ([party] if party.get("nt_num") else [])
                        for nt in notes:
                            ntty = _s(nt.get("ntty") or "")
                            if ntty == "C": ntty = "Credit Note"
                            elif ntty == "D": ntty = "Debit Note"
                            rcm = _s(nt.get("rchrg", "N"))
                            rcm_label = "Yes" if rcm == "Y" else "No"
                            for itm in (nt.get("itms") or [nt]):
                                d = itm.get("itm_det") or itm
                                all_rows.append([
                                    month, ctin, name, ntty, rcm_label,
                                    _s(nt.get("nt_num") or nt.get("nt_no") or ""), _s(nt.get("nt_dt")),
                                    _s(nt.get("inv_typ", "Regular")),
                                    _n(d.get("rt")), abs(_n(d.get("txval"))),
                                    abs(_n(d.get("iamt"))), abs(_n(d.get("camt"))), abs(_n(d.get("samt"))), abs(_n(d.get("csamt"))),
                                    abs(_n(nt.get("val"))),
                                    _state_name(_s(nt.get("pos") or party.get("pos"))), "GSTR-1"
                                ])
                else:
                    # Fallback: extract party-level CDN totals from the dashboard summary
                    summary_str = data.get("summary", "")
                    if summary_str:
                        try:
                            sum_obj = json.loads(summary_str) if isinstance(summary_str, str) else summary_str
                            seen_ctins = set()
                            for sec in sum_obj.get("data", {}).get("sec_sum", []):
                                sn = sec.get("sec_nm", "")
                                if sn == "CDNR" or sn.startswith("CDNR_"):
                                    # Handle nested sub_sections in summary
                                    sub_secs = sec.get("sub_sections", [sec])
                                    for ss in sub_secs:
                                        for party in ss.get("cpty_sum", []):
                                            ctin = _s(party.get("ctin", ""))
                                            if not ctin or ctin in seen_ctins:
                                                continue
                                            seen_ctins.add(ctin)
                                            name = _s(party.get("trdnm") or party.get("trade_name") or party.get("trdNm") or "")
                                            n = party.get("ttl_rec", 1)
                                            all_rows.append([
                                                month, ctin, name, "Note (Summary)", "No",
                                                f"({n} note{'s' if n != 1 else ''})", "",
                                                "Regular", "", abs(_n(party.get("ttl_tax"))),
                                                abs(_n(party.get("ttl_igst"))),
                                                abs(_n(party.get("ttl_cgst"))),
                                                abs(_n(party.get("ttl_sgst"))),
                                                abs(_n(party.get("ttl_cess"))),
                                                abs(_n(party.get("ttl_val"))),
                                                "", "GSTR-1"
                                            ])
                        except Exception:
                            pass

            elif key == "cdnur":
                if not headers: headers = ["Month", "Customer Name", "Supply Type", "Type of note (Debit/ Credit)", "Pre GST Regime Dr./ Cr. Notes", "Debit Note/ credit note/ Refund voucher No.", "Debit Note/ credit note/ Refund voucher Date", "Original Invoice No", "Original Invoice Date", "GST %", "Taxable Value", "IGST", "CGST", "SGST", "Cess", "Note/Refund Voucher Value", "Place of supply", "Return Type"]
                num_cols = {11, 12, 13, 14, 15, 16}
                _cdnur_typ_map = {
                    "EXPWOP": "Export without payment of GST",
                    "EXPWP":  "Export with payment of GST",
                    "B2CL":   "B2C Large",
                    "INTER":  "Inter-State",
                    "INTRA":  "Intra-State",
                }
                for nt in records:
                    ntty = _s(nt.get("ntty") or "")
                    if ntty == "C": ntty = "Credit Note"
                    elif ntty == "D": ntty = "Debit Note"
                    supply_type = _cdnur_typ_map.get(_s(nt.get("typ") or ""), _s(nt.get("typ") or ""))
                    p_gst = "Yes" if _s(nt.get("p_gst", "N")) == "Y" else "No"
                    for itm in (nt.get("itms") or [nt]):
                        d = itm.get("itm_det") or itm
                        all_rows.append([
                            month, None, supply_type, ntty, p_gst,
                            _s(nt.get("nt_num") or nt.get("nt_no") or ""), _s(nt.get("nt_dt")),
                            _s(nt.get("inum") or nt.get("inv_no") or ""), _s(nt.get("idt") or ""),
                            _n(d.get("rt")), abs(_n(d.get("txval"))),
                            abs(_n(d.get("iamt"))), abs(_n(d.get("camt"))), abs(_n(d.get("samt"))), abs(_n(d.get("csamt"))),
                            abs(_n(nt.get("val"))), _state_name(_s(nt.get("pos"))), "GSTR-1"
                        ])

            elif key == "exp":
                if not records: continue
                if not headers: headers = ["Month", "Type Of Export", "Customer Name", "Invoice No.", "Invoice Date", "Invoice Value", " Port Code", "Shipping bill/ bill of export No", "Shipping bill/ bill of export Date", "GST %", "Taxable Value", "IGST", "Cess", "Return Type"]
                num_cols = {5, 9, 10, 11, 12}
                _exp_typ_map = {
                    "WOPAY": "Export without payment of GST",
                    "WPAY":  "Export with payment of GST",
                }
                for party in records:
                    exp_typ = _exp_typ_map.get(_s(party.get("exp_typ") or ""), _s(party.get("exp_typ") or ""))
                    invoices = party.get("inv") or ([party] if party.get("inum") else [])
                    for inv in invoices:
                        for itm in (inv.get("itms") or [inv]):
                            all_rows.append([
                                month, exp_typ, None,
                                _s(inv.get("inum")), _s(inv.get("idt")), _n(inv.get("val")),
                                _s(inv.get("port_code")), _s(inv.get("sbnum")), _s(inv.get("sbdt")),
                                _nf(itm, "rt"),
                                _nf(itm, "txval", "invtxval"),
                                _nf(itm, "iamt", "inviamt"),
                                _nf(itm, "csamt", "invcsamt"), "GSTR-1"
                            ])

            elif key == "nil":
                if not headers: headers = ["Month", "Type", "Value", "Return Type", "Nil Rated", "Exempted", "Non-GST Supply"]
                num_cols = {3, 5, 6, 7}
                _nil_typ_map = {
                    "INTRB2B":  "Intra-State (Registered)",
                    "INTRB2C":  "Intra-State (Unregistered)",
                    "INTRB2CS": "Intra-State B2CS",
                    "EXPTB2B":  "Inter-State (Registered)",
                    "EXPTB2C":  "Inter-State (Unregistered)",
                }
                items = sec_data.get("inv") if isinstance(sec_data, dict) else sec_data
                for r in (items if isinstance(items, list) else []):
                    sply_ty = _s(r.get("sply_ty") or "")
                    sply_label = _nil_typ_map.get(sply_ty, sply_ty)
                    nil_amt  = _n(r.get("nil_amt"))
                    expt_amt = _n(r.get("expt_amt"))
                    ngsup    = _n(r.get("ngsup_amt"))
                    total    = (nil_amt or 0) + (expt_amt or 0) + (ngsup or 0)
                    all_rows.append([month, sply_label, total, "GSTR-1", nil_amt, expt_amt, ngsup])

            elif key == "hsnsum":
                if not headers: headers = ["Month", "HSN", "Record Type", "Description", "UQC", "Total Quantity", "Taxable Value", "GST %", "IGST", "CGST", "SGST", "Cess", "Total Value", "Return Type"]
                num_cols = {6, 7, 9, 10, 11, 12}

                seen_hsn_sigs = set()
                def _hsn_row(item, rec_type):
                    # Deduplicate: same HSN, rate, and quantity/taxable value in the same month
                    # Include rec_type and uqc to distinguish valid separate records (e.g. B2B vs B2C)
                    sig = f"{month}|{rec_type}|{_s(item.get('hsn_sc'))}|{_n(item.get('rt'))}|{_n(item.get('qty'))}|{_n(item.get('txval'))}|{_s(item.get('uqc'))}"
                    if sig in seen_hsn_sigs: return
                    seen_hsn_sigs.add(sig)

                    desc = _s(item.get("desc") or item.get("user_desc") or "")
                    all_rows.append([
                        month, _s(item.get("hsn_sc")), rec_type, desc,
                        _s(item.get("uqc")), _n(item.get("qty")),
                        _nf(item, "txval", "invtxval"), _nf(item, "rt"),
                        _nf(item, "iamt", "inviamt"), _nf(item, "camt", "invcamt"), _nf(item, "samt", "invsamt"), _nf(item, "csamt", "invcsamt"),
                        "-", "GSTR-1"
                    ])

                for r in records:
                    if "hsn_b2b" in r or "hsn_b2c" in r:
                        # ZIP format: separate B2B and B2C lists
                        for item in r.get("hsn_b2b", []):
                            _hsn_row(item, "B2B")
                        for item in r.get("hsn_b2c", []):
                            _hsn_row(item, "B2C")
                    else:
                        # API format: flat record with optional typ field
                        _hsn_row(r, _s(r.get("hsn_typ") or r.get("typ") or ""))

            elif key == "at":
                if not records: continue
                if not headers: headers = ["Month", "Gross Advance Received", "GST %", "IGST", "CGST", "SGST", "Cess", "Place of Supply", "Return Type"]
                num_cols = {2, 4, 5, 6, 7}
                for rec in records:
                    for itm in (rec.get("itms") or [rec]):
                        all_rows.append([
                            month, _n(itm.get("ad_amt")), _n(itm.get("rt")),
                            _n(itm.get("iamt")), _n(itm.get("camt")),
                            _n(itm.get("samt")), _n(itm.get("csamt")),
                            _state_name(_s(rec.get("pos"))), "GSTR-1"
                        ])

            elif key == "txp":
                if not records: continue
                if not headers: headers = ["Month", "Advance adjusted", "Rate", "IGST", "CGST", "SGST", "Cess", "Place Of Supply", "Return Type"]
                num_cols = {2, 4, 5, 6, 7}
                for rec in records:
                    for itm in (rec.get("itms") or [rec]):
                        all_rows.append([
                            month, _n(itm.get("ad_amt")), _n(itm.get("rt")),
                            _n(itm.get("iamt")), _n(itm.get("camt")),
                            _n(itm.get("samt")), _n(itm.get("csamt")),
                            _state_name(_s(rec.get("pos"))), "GSTR-1"
                        ])

            elif key == "dociss":
                if not headers: headers = ["Nature of Document", "Sr. No. From", "Sr. No. To", "Total Number", "Cancelled", "Net Issued", "Return Type"]
                num_cols = {4, 5, 6}
                if records:
                    for rec in records:
                        doc_typ = rec.get("doc_num") or rec.get("doc_typ") or rec.get("ty")
                        docs = rec.get("docs") or rec.get("doc") or []
                        for d_item in docs:
                            all_rows.append([
                                _doc_name(doc_typ), _s(d_item.get("from")), _s(d_item.get("to")),
                                _n(d_item.get("totnum")), _n(d_item.get("cancel")),
                                _n(d_item.get("net_issue")), "GSTR-1"
                            ])
                else:
                    # Fallback from Dashboard Summary
                    summary_str = data.get("summary", "")
                    if summary_str:
                        try:
                            s = json.loads(summary_str) if isinstance(summary_str, str) else summary_str
                            for sec in s.get("data", {}).get("sec_sum", []):
                                if sec.get("sec_nm") == "DOC_ISSUE":
                                    all_rows.append([
                                        "Total Documents (Summary)", "-", "-",
                                        _n(sec.get("ttl_doc_issued")),
                                        _n(sec.get("ttl_doc_cancelled")),
                                        _n(sec.get("net_doc_issued")),
                                        "GSTR-1"
                                    ])
                        except: pass

            elif key == "ecom":
                if not headers: headers = ["Month", "GSTIN of e-Commerce Operator", "Trade Name", "ECO Liable to Pay Tax \nu/s 9(5)", "Net Value of Supplies", "IGST", "CGST", "SGST", "Cess"]
                num_cols = {5, 6, 7, 8, 9}
                for rec in records:
                    all_rows.append([
                        month, _s(rec.get("etin")), _s(rec.get("trdnm")), "",
                        _n(rec.get("txval")), _n(rec.get("iamt")), _n(rec.get("camt")),
                        _n(rec.get("samt")), _n(rec.get("csamt"))
                    ])

        ws = wb.create_sheet(sheet_name)
        if all_rows and headers:
            _write_portal_sheet(ws, headers, all_rows, num_cols, profile=prof, report_name=report_title, fy=fy_str)
        else:
            # Write header-only sheet matching portal layout
            if not headers:
                headers = _default_headers.get(key, [])
            _write_portal_sheet(ws, headers, [], num_cols, profile=prof, report_name=report_title, fy=fy_str)

    if not wb.sheetnames:
        _empty_sheet(wb)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

def gstr1_to_excel(data: dict, out_path: str, profile: dict = None):
    """Wrapper to use consolidated logic even for single month if desired, or keep legacy."""
    if isinstance(data, list):
        return gstr1_consolidated_to_excel(data, out_path, profile)
    return gstr1_consolidated_to_excel([data], out_path, profile)
