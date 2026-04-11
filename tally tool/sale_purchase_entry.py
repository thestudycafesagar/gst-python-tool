"""
TallySalesPro - Professional Sales Voucher & Master Creator for TallyPrime
Converts Excel sales data to Tally XML (Accounting & Inventory voucher modes)
Also creates Ledger Masters and Stock Item Masters.
Author: Studycafe | Built with CustomTkinter
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
import openpyxl
import os
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta
import urllib.request
from urllib.error import HTTPError, URLError
import json
import threading
import re
import html

# ─── Theme ──────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

COLORS = {
    "bg_dark": ("#F0F4F8", "#0F172A"),
    "bg_card": ("#FFFFFF", "#1E293B"),
    "bg_card_hover": ("#E2E8F0", "#334155"),
    "bg_input": ("#F1F5F9", "#1E293B"),
    "accent": ("#2563EB", "#3B82F6"),
    "accent_hover": ("#1D4ED8", "#2563EB"),
    "success": ("#059669", "#10B981"),
    "warning": ("#D97706", "#F59E0B"),
    "error": ("#DC2626", "#EF4444"),
    "text_primary": ("#0F172A", "#F1F5F9"),
    "text_secondary": ("#475569", "#CBD5E1"),
    "text_muted": ("#64748B", "#94A3B8"),
    "border": ("#E2E8F0", "#334155"),
    "table_header": ("#1E293B", "#0F172A"),
}

ACCENT = COLORS["accent"]
ACCENT_HOVER = COLORS["accent_hover"]
SUCCESS = COLORS["success"]
DANGER = COLORS["error"]
SURFACE = COLORS["bg_card"]
SURFACE2 = COLORS["bg_input"]
TEXT_PRIMARY = COLORS["text_primary"]
TEXT_MUTED = COLORS["text_muted"]
CARD_BG = COLORS["bg_dark"]
SUSPENSE_LEDGER = "Suspense A/c"

# ─── Tally XML Helpers ──────────────────────────────────────────────────────

def xml_escape(s: str) -> str:
    if not s:
        return ""
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;").replace("'","&apos;")

def fmt_amt(num: float) -> str:
    return f"{num:.2f}"

def tally_date(dt) -> str:
    today = datetime.today().strftime("%Y%m%d")

    if dt in (None, ""):
        return today
    if isinstance(dt, datetime):
        return dt.strftime("%Y%m%d")
    if isinstance(dt, date):
        return dt.strftime("%Y%m%d")

    # Excel serial date support.
    if isinstance(dt, (int, float)) and not isinstance(dt, bool):
        try:
            if float(dt) > 1000:
                return (datetime(1899, 12, 30) + timedelta(days=float(dt))).strftime("%Y%m%d")
        except Exception:
            pass

    text = str(dt).strip()
    if not text:
        return today

    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]

    # Already in compact numeric form.
    if re.fullmatch(r"\d{8}", text):
        if 1900 <= int(text[:4]) <= 2100:
            return text
        return f"{text[4:8]}{text[2:4]}{text[:2]}"

    candidates = [text]
    if " " in text:
        candidates.append(text.split(" ", 1)[0])

    formats = (
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%d-%b-%Y",
        "%d-%b-%y",
        "%d-%B-%Y",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
    )
    for candidate in candidates:
        for fmt in formats:
            try:
                return datetime.strptime(candidate, fmt).strftime("%Y%m%d")
            except ValueError:
                continue

    return today

def push_to_tally(xml_str: str, host: str = "localhost", port: int = 9000) -> str:
    """Send XML to TallyPrime HTTP API and return response."""
    url = f"http://{host}:{port}"
    req = urllib.request.Request(url, data=xml_str.encode("utf-8"),
                                 headers={"Content-Type":"application/xml"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def _row_get(row: dict, key: str, default=None):
    """Read a row value by key while tolerating header spacing/case differences."""
    if key in row:
        value = row.get(key)
        return default if value is None else value

    target = re.sub(r"\s+", "", str(key or "")).lower()
    for raw_key, raw_value in row.items():
        normalized = re.sub(r"\s+", "", str(raw_key or "")).lower()
        if normalized == target:
            return default if raw_value is None else raw_value
    return default


def _row_text(row: dict, key: str, default: str = "") -> str:
    value = _row_get(row, key, default)
    if value is None:
        return default
    return str(value).strip()


def _row_float(row: dict, key: str, default: float = 0.0) -> float:
    value = _row_get(row, key, default)
    if value in (None, ""):
        return float(default)
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def _row_voucher_number(row: dict, default: str = "") -> str:
    return (
        _row_text(row, "VoucherNo")
        or _row_text(row, "InvoiceNo")
        or _row_text(row, "BillNo")
        or default
    )


def _row_invoice_reference(row: dict, default: str = "") -> str:
    return (
        _row_text(row, "InvoiceNo")
        or _row_text(row, "SupplierInvoiceNo")
        or _row_text(row, "BillNo")
        or _row_text(row, "VoucherNo")
        or default
    )


def _ledger_or_suspense(value: str, fallback: str = SUSPENSE_LEDGER) -> str:
    text = str(value or "").strip()
    return text or fallback


def _company_static_block(company: str) -> str:
    selected = str(company or "").strip()
    if not selected:
        return ""
    return f"   <STATICVARIABLES><SVCURRENTCOMPANY>{xml_escape(selected)}</SVCURRENTCOMPANY></STATICVARIABLES>"


def _normalize_company_name(value) -> str:
    text = html.unescape(str(value or ""))
    text = text.replace("\x00", "")
    text = re.sub(r"[\x01-\x1F\x7F]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _company_key(value) -> str:
    return _normalize_company_name(value).upper()


def _is_valid_company_name(value) -> bool:
    name = _normalize_company_name(value)
    if not name:
        return False
    return not name.isdigit()


def _build_tally_url(host: str, port: str) -> str:
    host_text = (host or "localhost").strip()
    port_text = (port or "9000").strip()
    if host_text.startswith("http://"):
        host_text = host_text[7:]
    elif host_text.startswith("https://"):
        host_text = host_text[8:]
    host_text = host_text.strip("/") or "localhost"
    if "/" in host_text:
        host_text = host_text.split("/", 1)[0]
    if not port_text.isdigit():
        raise ValueError("Port must be numeric.")
    return f"http://{host_text}:{port_text}"


def _post_tally_xml(tally_url: str, xml_payload: str, timeout: float = 15.0) -> str:
    request = urllib.request.Request(
        tally_url,
        data=xml_payload.encode("utf-8"),
        headers={"Content-Type": "application/xml"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read().decode("utf-8", errors="replace")


def _check_tally_connection(tally_url: str, timeout: float = 5.0) -> dict:
    probe_xml = (
        "<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>"
        "<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Companies</REPORTNAME>"
        "</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>"
    )
    try:
        _post_tally_xml(tally_url, probe_xml, timeout=timeout)
        return {"connected": True}
    except HTTPError as exc:
        return {"connected": False, "error": f"HTTP {exc.code}"}
    except URLError:
        return {"connected": False, "error": "ConnectionError"}
    except Exception as exc:
        return {"connected": False, "error": str(exc)}


def _extract_company_names(response_text: str) -> set:
    names = set()
    try:
        root = ET.fromstring(response_text)
        for node in root.iter():
            tag = str(node.tag or "").upper()
            txt = _normalize_company_name(node.text)
            attr_name = _normalize_company_name(node.attrib.get("NAME") or "")
            if tag in {"COMPANYNAME", "SVCURRENTCOMPANY", "CURRENTCOMPANY"} and _is_valid_company_name(txt):
                names.add(txt)
            if "COMPANY" in tag and _is_valid_company_name(attr_name):
                names.add(attr_name)
            if tag == "COMPANY" and _is_valid_company_name(txt):
                names.add(txt)
    except ET.ParseError:
        pass

    patterns = [
        r'COMPANY[^>]*NAME="([^"]+)"',
        r"<COMPANYNAME>(.*?)</COMPANYNAME>",
        r"<SVCURRENTCOMPANY>(.*?)</SVCURRENTCOMPANY>",
        r"<COMPANY[^>]*>.*?<NAME>(.*?)</NAME>",
    ]
    for pattern in patterns:
        for match in re.findall(pattern, response_text, flags=re.IGNORECASE | re.DOTALL):
            value = _normalize_company_name(match)
            if _is_valid_company_name(value):
                names.add(value)
    return names


def _fetch_tally_companies(tally_url: str, timeout: float = 15.0) -> dict:
    requests_xml = [
        (
            "report-list-companies",
            "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>"
            "<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Companies</REPORTNAME>"
            "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
            "</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>",
        ),
        (
            "collection-company",
            "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
            "<TYPE>Collection</TYPE><ID>Company Collection</ID></HEADER><BODY><DESC>"
            "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
            "<TDL><TDLMESSAGE><COLLECTION NAME='Company Collection'>"
            "<TYPE>Company</TYPE><FETCH>Name</FETCH><NATIVEMETHOD>Name</NATIVEMETHOD>"
            "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>",
        ),
    ]

    companies = set()
    errors = []
    for label, xml_payload in requests_xml:
        try:
            response_text = _post_tally_xml(tally_url, xml_payload, timeout=timeout)
            companies.update(_extract_company_names(response_text))
        except HTTPError as exc:
            errors.append(f"{label}: HTTP {exc.code}")
        except URLError:
            errors.append(f"{label}: ConnectionError")
        except Exception as exc:
            errors.append(f"{label}: {exc}")

    sorted_companies = sorted(companies, key=lambda x: _company_key(x))
    if sorted_companies:
        return {"success": True, "companies": sorted_companies}

    err = "; ".join(errors) if errors else "No companies returned by Tally."
    return {"success": False, "error": err, "companies": []}


def _safe_int(text, default=0) -> int:
    try:
        return int(float(str(text).strip()))
    except (TypeError, ValueError):
        return default


def _parse_tally_response_details(response_text: str) -> dict:
    details = {
        "success": False,
        "created": 0,
        "altered": 0,
        "deleted": 0,
        "lastvchid": 0,
        "lastmid": 0,
        "combined": 0,
        "ignored": 0,
        "errors": 0,
        "cancelled": 0,
        "exceptions": 0,
        "line_errors": [],
        "error": "",
    }

    try:
        root = ET.fromstring(response_text)
    except ET.ParseError as exc:
        details["error"] = f"Could not parse Tally response: {exc}"
        return details

    def _count(tag_name: str) -> int:
        node = root.find(f".//{tag_name}")
        return _safe_int(node.text if node is not None else 0)

    details["created"] = _count("CREATED")
    details["altered"] = _count("ALTERED")
    details["deleted"] = _count("DELETED")
    details["lastvchid"] = _count("LASTVCHID")
    details["lastmid"] = _count("LASTMID")
    details["combined"] = _count("COMBINED")
    details["ignored"] = _count("IGNORED")
    details["errors"] = _count("ERRORS")
    details["cancelled"] = _count("CANCELLED")
    details["exceptions"] = _count("EXCEPTIONS")

    line_errors = []
    seen = set()
    for node in root.findall(".//LINEERROR"):
        text = (node.text or "").strip()
        if text and text not in seen:
            seen.add(text)
            line_errors.append(text)
    details["line_errors"] = line_errors

    details["success"] = (
        details["errors"] == 0
        and details["exceptions"] == 0
        and not details["line_errors"]
    )
    if not details["success"] and details["line_errors"]:
        details["error"] = details["line_errors"][0]
    return details


def _extract_missing_ledgers_from_line_errors(line_errors: list) -> list:
    missing = []
    seen = set()
    patterns = [
        re.compile(r"Ledger\s+'([^']+)'\s+does\s+not\s+exist", re.IGNORECASE),
        re.compile(r'Ledger\s+"([^"]+)"\s+does\s+not\s+exist', re.IGNORECASE),
    ]

    for err in line_errors or []:
        text = str(err or "")
        for pattern in patterns:
            for raw_name in pattern.findall(text):
                name = str(raw_name or "").strip()
                key = name.casefold()
                if not name or key in seen:
                    continue
                seen.add(key)
                missing.append(name)
    return missing


def _collect_auto_voucher_ledgers(rows: list, mode: str) -> list:
    """Collect non-party ledgers from voucher rows so missing ledgers can be pre-created."""
    entries = {}
    is_purchase_mode = mode in {"purchase_accounting", "purchase_item"}

    def add_entry(name: str, parent: str, tax_type: str = "", gst_rate: str = ""):
        ledger_name = str(name or "").strip()
        if not ledger_name:
            return

        key = re.sub(r"\s+", " ", ledger_name).strip().casefold()
        if not key:
            return

        candidate = {
            "Name": ledger_name,
            "Parent": parent,
            "GSTApplicable": "",
            "GSTIN": "",
            "StateOfSupply": "",
            "TypeOfTaxation": tax_type or "",
            "GSTRate": str(gst_rate).strip() if gst_rate not in (None, "") else "",
        }

        existing = entries.get(key)
        if existing is None:
            entries[key] = candidate
            return

        existing_parent = str(existing.get("Parent", "")).strip().casefold()
        candidate_parent = str(candidate.get("Parent", "")).strip().casefold()
        if existing_parent != "duties & taxes" and candidate_parent == "duties & taxes":
            entries[key] = candidate
            return

        if not existing.get("Parent") and candidate.get("Parent"):
            existing["Parent"] = candidate["Parent"]
        if not existing.get("TypeOfTaxation") and candidate.get("TypeOfTaxation"):
            existing["TypeOfTaxation"] = candidate["TypeOfTaxation"]
        if not existing.get("GSTRate") and candidate.get("GSTRate"):
            existing["GSTRate"] = candidate["GSTRate"]

    for r in rows or []:
        if is_purchase_mode:
            purchase_ledger = (
                _row_text(r, "PurchaseLedger")
                or _row_text(r, "PurchaseAccount")
                or _row_text(r, "Purchase Ledger")
                or _row_text(r, "ExpenseLedger")
                or _row_text(r, "SalesLedger")
            )
            add_entry(purchase_ledger, "Purchase Accounts")
        else:
            sales_ledger = (
                _row_text(r, "SalesLedger")
                or _row_text(r, "SalesAccount")
                or _row_text(r, "Sales Ledger")
                or _row_text(r, "IncomeLedger")
            )
            add_entry(sales_ledger, "Sales Accounts")

        add_entry(_row_text(r, "CGSTLedger"), "Duties & Taxes", "Central Tax", _row_text(r, "CGSTRate"))
        add_entry(_row_text(r, "SGSTLedger"), "Duties & Taxes", "State Tax", _row_text(r, "SGSTRate"))
        add_entry(_row_text(r, "IGSTLedger"), "Duties & Taxes", "Integrated Tax", _row_text(r, "IGSTRate"))

    return list(entries.values())


def _build_missing_ledger_defs(line_errors: list, rows: list, mode: str) -> list:
    missing_names = _extract_missing_ledgers_from_line_errors(line_errors)
    if not missing_names:
        return []

    is_purchase_mode = mode in {"purchase_accounting", "purchase_item"}

    party_keys = set()
    purchase_keys = set()
    sales_keys = set()
    tax_type_map = {}
    tax_rate_map = {}

    for r in rows or []:
        party_name = _row_text(r, "PartyLedger")
        if party_name:
            party_keys.add(party_name.casefold())

        purchase_ledger = (
            _row_text(r, "PurchaseLedger")
            or _row_text(r, "PurchaseAccount")
            or _row_text(r, "Purchase Ledger")
            or _row_text(r, "ExpenseLedger")
            or _row_text(r, "SalesLedger")
        )
        if purchase_ledger:
            purchase_keys.add(purchase_ledger.casefold())

        sales_ledger = (
            _row_text(r, "SalesLedger")
            or _row_text(r, "SalesAccount")
            or _row_text(r, "Sales Ledger")
            or _row_text(r, "IncomeLedger")
        )
        if sales_ledger:
            sales_keys.add(sales_ledger.casefold())

        for ledger_col, rate_col, tax_type in (
            ("CGSTLedger", "CGSTRate", "Central Tax"),
            ("SGSTLedger", "SGSTRate", "State Tax"),
            ("IGSTLedger", "IGSTRate", "Integrated Tax"),
        ):
            tax_name = _row_text(r, ledger_col)
            if not tax_name:
                continue
            key = tax_name.casefold()
            tax_type_map.setdefault(key, tax_type)
            rate_val = _row_text(r, rate_col)
            if rate_val and key not in tax_rate_map:
                tax_rate_map[key] = rate_val

    entries = []
    seen = set()
    for ledger_name in missing_names:
        key = ledger_name.casefold()
        if key in seen:
            continue
        seen.add(key)

        tax_type = ""
        gst_rate = ""
        if key in party_keys:
            parent = "Sundry Creditors" if is_purchase_mode else "Sundry Debtors"
        elif key in purchase_keys:
            parent = "Purchase Accounts"
        elif key in sales_keys:
            parent = "Sales Accounts"
        elif key in tax_type_map or "gst" in key or "tax" in key:
            parent = "Duties & Taxes"
            tax_type = tax_type_map.get(key, "")
            gst_rate = tax_rate_map.get(key, "")
            if not tax_type:
                if "igst" in key:
                    tax_type = "Integrated Tax"
                elif "cgst" in key:
                    tax_type = "Central Tax"
                elif "sgst" in key:
                    tax_type = "State Tax"
        else:
            parent = "Purchase Accounts" if is_purchase_mode else "Sales Accounts"

        entries.append(
            {
                "Name": ledger_name,
                "Parent": parent,
                "GSTApplicable": "",
                "GSTIN": "",
                "StateOfSupply": "",
                "TypeOfTaxation": tax_type,
                "GSTRate": str(gst_rate).strip() if gst_rate not in (None, "") else "",
            }
        )

    return entries


def _extract_numeric_voucher_no(value):
    text = str(value or "").strip()
    if not text:
        return None
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    if text.isdigit():
        return int(text)
    return None


def _extract_voucher_numbers(response_text: str) -> list:
    found = []
    seen = set()

    def add_value(raw):
        num = _extract_numeric_voucher_no(raw)
        if num is None or num in seen:
            return
        seen.add(num)
        found.append(num)

    try:
        root = ET.fromstring(response_text)
        for node in root.iter():
            tag = str(node.tag or "").upper()
            if "VOUCHERNUM" in tag:
                add_value(node.text)
            for key, raw_val in node.attrib.items():
                if "VOUCHERNUM" in str(key or "").upper():
                    add_value(raw_val)
    except ET.ParseError:
        pass

    for match in re.findall(r"<VOUCHERNUMBER>(.*?)</VOUCHERNUMBER>", response_text, flags=re.IGNORECASE | re.DOTALL):
        add_value(match)

    return found


def _fetch_next_voucher_number(tally_url: str, company_name: str = "", voucher_type: str = "Sales", timeout: float = 15.0) -> dict:
    selected_company = _normalize_company_name(company_name)
    escaped_voucher_type = xml_escape(voucher_type or "Sales")

    static_vars = "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    static_vars += f"<SVVOUCHERTYPENAME>{escaped_voucher_type}</SVVOUCHERTYPENAME>"
    if selected_company:
        static_vars += f"<SVCURRENTCOMPANY>{xml_escape(selected_company)}</SVCURRENTCOMPANY>"
    static_vars += "</STATICVARIABLES>"

    request_variants = [
        (
            "report-list-vouchers",
            "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>"
            "<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Vouchers</REPORTNAME>"
            f"{static_vars}</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>",
        ),
        (
            "collection-vouchers",
            "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
            "<TYPE>Collection</TYPE><ID>Voucher Number Collection</ID></HEADER><BODY><DESC>"
            f"{static_vars}"
            "<TDL><TDLMESSAGE><COLLECTION NAME='Voucher Number Collection'>"
            "<TYPE>Voucher</TYPE><FETCH>VoucherNumber,VoucherTypeName,Date</FETCH>"
            "<FILTERS>OnlySales</FILTERS>"
            "</COLLECTION>"
            "<SYSTEM TYPE='Formulae' NAME='OnlySales'>$$IsSales:$VoucherTypeName</SYSTEM>"
            "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>",
        ),
    ]

    all_numbers = []
    had_response = False
    errors = []

    for label, xml_payload in request_variants:
        try:
            response_text = _post_tally_xml(tally_url, xml_payload, timeout=timeout)
            had_response = True
            all_numbers.extend(_extract_voucher_numbers(response_text))
        except Exception as exc:
            errors.append(f"{label}: {exc}")

    if all_numbers:
        last_no = max(all_numbers)
        return {
            "success": True,
            "last_number": last_no,
            "next_number": last_no + 1,
            "error": "",
        }

    if had_response:
        # No existing numeric vouchers found in selected type/company.
        return {
            "success": True,
            "last_number": 0,
            "next_number": 1,
            "error": "",
        }

    return {
        "success": False,
        "last_number": 0,
        "next_number": 0,
        "error": "; ".join(errors) if errors else "Could not fetch voucher number from Tally.",
    }


# ═══════════════════════════════════════════════════════════════════════════
#  GENERATE SALES XML  –  ACCOUNTING MODE  (mirrors original VBA logic)
# ═══════════════════════════════════════════════════════════════════════════

def generate_accounting_xml(rows: list, company: str, use_today_date: bool = False, start_voucher_number=None) -> str:
    """rows = list of dicts with keys matching Excel columns."""
    lines = []
    a = lines.append
    company_static = _company_static_block(company)
    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    for idx, r in enumerate(rows):
        source_date = datetime.today() if use_today_date else _row_get(r, "Date", "")
        dt       = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno      = xml_escape(vno_raw)
        party_raw = _ledger_or_suspense(_row_text(r, "PartyLedger"))
        sales_raw = _ledger_or_suspense(_row_text(r, "SalesLedger"))
        party = xml_escape(party_raw)
        sales = xml_escape(sales_raw)
        taxable  = _row_float(r, "TaxableValue", 0.0)
        cgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "CGSTLedger")))
        cgst_r   = _row_float(r, "CGSTRate", 0.0)
        sgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "SGSTLedger")))
        sgst_r   = _row_float(r, "SGSTRate", 0.0)
        igst_led = xml_escape(_ledger_or_suspense(_row_text(r, "IGSTLedger")))
        igst_r   = _row_float(r, "IGSTRate", 0.0)
        narr     = xml_escape(_row_text(r, "Narration"))
        party_gstin = xml_escape(_row_text(r, "PartyGSTIN") or _row_text(r, "GSTIN"))
        place_of_supply = xml_escape(_row_text(r, "PlaceOfSupply"))

        cgst_amt = round(taxable * cgst_r / 100, 2) if cgst_r > 0 else 0
        sgst_amt = round(taxable * sgst_r / 100, 2) if sgst_r > 0 else 0
        igst_amt = round(taxable * igst_r / 100, 2) if igst_r > 0 else 0
        total    = taxable + cgst_amt + sgst_amt + igst_amt

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a('    <VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">')
        a(f'     <DATE>{dt}</DATE>')
        a('     <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>')
        a(f'     <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>')
        a(f'     <PARTYLEDGERNAME>{party}</PARTYLEDGERNAME>')
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>')
        if party_gstin:
            a(f'     <PARTYGSTIN>{party_gstin}</PARTYGSTIN>')
        if place_of_supply:
            a(f'     <PLACEOFSUPPLY>{place_of_supply}</PLACEOFSUPPLY>')
        if narr:
            a(f'     <NARRATION>{narr}</NARRATION>')

        # Party – Debit
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{party}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>-{fmt_amt(total)}</AMOUNT>')
        a('     </LEDGERENTRIES.LIST>')

        # Sales – Credit
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{sales}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>{fmt_amt(taxable)}</AMOUNT>')
        a('     </LEDGERENTRIES.LIST>')

        # CGST
        if cgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{cgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(cgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        # SGST
        if sgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{sgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(sgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        # IGST
        if igst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{igst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(igst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')

        a('    </VOUCHER>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════
#  GENERATE SALES XML  –  ITEM / INVENTORY MODE
# ═══════════════════════════════════════════════════════════════════════════

def generate_item_xml(
    rows: list,
    company: str,
    use_today_date: bool = False,
    start_voucher_number=None,
    fallback_sales_ledger: str = SUSPENSE_LEDGER,
) -> str:
    """
    Item-mode sales voucher. Each row needs additional columns:
      ItemName, Quantity, Rate, Per (unit), GodownName (optional)
    Uses ALLINVENTORYENTRIES.LIST for stock items +
         LEDGERENTRIES.LIST for accounting legs (party, tax ledgers).
    """
    lines = []
    a = lines.append
    company_static = _company_static_block(company)
    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    def _name_key(value: str) -> str:
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    for idx, r in enumerate(rows):
        source_date = datetime.today() if use_today_date else _row_get(r, "Date", "")
        dt       = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno      = xml_escape(vno_raw)
        party = xml_escape(_ledger_or_suspense(_row_text(r, "PartyLedger")))
        taxable  = _row_float(r, "TaxableValue", 0.0)

        # Support common source variants while preserving current UI/flow.
        item_name_raw = (
            _row_text(r, "ItemName")
            or _row_text(r, "Item")
            or _row_text(r, "StockItem")
            or _row_text(r, "ProductName")
            or _row_text(r, "SalesLedger")
        )
        if not item_name_raw:
            raise ValueError(f"Item row {idx + 1}: item name is missing (ItemName/Item/SalesLedger).")
        item_name = xml_escape(item_name_raw)
        item_name_key = _name_key(item_name_raw)

        qty = _row_float(r, "Quantity", 0.0)
        if qty <= 0:
            qty = _row_float(r, "Qty", 0.0)
        if qty <= 0:
            qty = _row_float(r, "Unit", 0.0)
        if qty <= 0:
            raise ValueError(f"Item row {idx + 1}: quantity is missing/zero (Quantity/Qty/Unit).")

        rate = _row_float(r, "Rate", 0.0)
        if rate <= 0 and taxable > 0 and qty > 0:
            rate = taxable / qty
        per_unit_raw = (
            _row_text(r, "Per", "")
            or _row_text(r, "UOM", "")
            or _row_text(r, "Unit", "")
            or "Nos"
        )
        per_unit = xml_escape(_normalize_stock_unit_name(per_unit_raw) or "Nos")
        godown   = xml_escape(_row_text(r, "GodownName", "Main Location") or "Main Location")

        explicit_sales_ledger = (
            _row_text(r, "SalesAccount")
            or _row_text(r, "Sales Ledger")
            or _row_text(r, "IncomeLedger")
        )
        default_sales_ledger = _ledger_or_suspense(fallback_sales_ledger)
        if _name_key(default_sales_ledger) == item_name_key:
            for candidate in ("Sales Account", "Sales", "Sales A/c", "Sales Ledger"):
                if _name_key(candidate) != item_name_key:
                    default_sales_ledger = candidate
                    break

        sales_ledger_raw = explicit_sales_ledger or _row_text(r, "SalesLedger") or default_sales_ledger
        # If source ledger equals stock item name, switch to fallback sales ledger.
        if _name_key(sales_ledger_raw) == item_name_key:
            sales_ledger_raw = default_sales_ledger
        sales_ledger_raw = _ledger_or_suspense(sales_ledger_raw, default_sales_ledger)
        if _name_key(sales_ledger_raw) == item_name_key:
            raise ValueError(
                f"Item row {idx + 1}: sales ledger cannot be same as item '{item_name_raw}'. "
                "Provide SalesAccount/IncomeLedger in Excel or use a valid fallback sales ledger."
            )
        sales = xml_escape(sales_ledger_raw)

        cgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "CGSTLedger")))
        cgst_r   = _row_float(r, "CGSTRate", 0.0)
        sgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "SGSTLedger")))
        sgst_r   = _row_float(r, "SGSTRate", 0.0)
        igst_led = xml_escape(_ledger_or_suspense(_row_text(r, "IGSTLedger")))
        igst_r   = _row_float(r, "IGSTRate", 0.0)
        narr     = xml_escape(_row_text(r, "Narration"))
        party_gstin = xml_escape(_row_text(r, "PartyGSTIN") or _row_text(r, "GSTIN"))
        place_of_supply = xml_escape(_row_text(r, "PlaceOfSupply"))
        hsn_code = xml_escape(_row_text(r, "HSNCode"))

        item_amt  = round(qty * rate, 2) if qty and rate else taxable
        cgst_amt = round(taxable * cgst_r / 100, 2) if cgst_r > 0 else 0
        sgst_amt = round(taxable * sgst_r / 100, 2) if sgst_r > 0 else 0
        igst_amt = round(taxable * igst_r / 100, 2) if igst_r > 0 else 0
        total    = taxable + cgst_amt + sgst_amt + igst_amt

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a('    <VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">')
        a(f'     <DATE>{dt}</DATE>')
        a('     <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>')
        a(f'     <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>')
        a(f'     <PARTYLEDGERNAME>{party}</PARTYLEDGERNAME>')
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>')
        if party_gstin:
            a(f'     <PARTYGSTIN>{party_gstin}</PARTYGSTIN>')
        if place_of_supply:
            a(f'     <PLACEOFSUPPLY>{place_of_supply}</PLACEOFSUPPLY>')
        if narr:
            a(f'     <NARRATION>{narr}</NARRATION>')

        # ── Inventory entry ──
        a('     <ALLINVENTORYENTRIES.LIST>')
        a(f'      <STOCKITEMNAME>{item_name}</STOCKITEMNAME>')
        a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
        a(f'      <RATE>{fmt_amt(rate)}/{per_unit}</RATE>')
        a(f'      <AMOUNT>{fmt_amt(item_amt)}</AMOUNT>')
        a(f'      <ACTUALQTY>{fmt_amt(qty)} {per_unit}</ACTUALQTY>')
        a(f'      <BILLEDQTY>{fmt_amt(qty)} {per_unit}</BILLEDQTY>')
        a('      <BATCHALLOCATIONS.LIST>')
        a(f'       <GODOWNNAME>{godown}</GODOWNNAME>')
        a(f'       <AMOUNT>{fmt_amt(item_amt)}</AMOUNT>')
        a(f'       <ACTUALQTY>{fmt_amt(qty)} {per_unit}</ACTUALQTY>')
        a(f'       <BILLEDQTY>{fmt_amt(qty)} {per_unit}</BILLEDQTY>')
        a('      </BATCHALLOCATIONS.LIST>')
        a('      <ACCOUNTINGALLOCATIONS.LIST>')
        a(f'       <LEDGERNAME>{sales}</LEDGERNAME>')
        a('       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
        a(f'       <AMOUNT>{fmt_amt(item_amt)}</AMOUNT>')
        a('      </ACCOUNTINGALLOCATIONS.LIST>')
        a('     </ALLINVENTORYENTRIES.LIST>')

        # ── Ledger entries (Party DR) ──
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{party}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>-{fmt_amt(total)}</AMOUNT>')
        a('      <BILLALLOCATIONS.LIST>')
        a(f'       <NAME>{vno}</NAME>')
        a('       <BILLTYPE>New Ref</BILLTYPE>')
        a(f'       <AMOUNT>-{fmt_amt(total)}</AMOUNT>')
        a('      </BILLALLOCATIONS.LIST>')
        a('     </LEDGERENTRIES.LIST>')

        # ── Tax ledgers ──
        if cgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{cgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(cgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        if sgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{sgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(sgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        if igst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{igst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(igst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')

        a('    </VOUCHER>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')
    return "\n".join(lines)


def generate_purchase_accounting_xml(
    rows: list,
    company: str,
    use_today_date: bool = False,
    start_voucher_number=None,
) -> str:
    """Purchase accounting invoice XML (mirror of sales with debit/credit reversed)."""
    lines = []
    a = lines.append
    company_static = _company_static_block(company)
    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    for idx, r in enumerate(rows):
        source_date = datetime.today() if use_today_date else _row_get(r, "Date", "")
        dt = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno = xml_escape(vno_raw)
        supplier_invoice_raw = _row_invoice_reference(r, vno_raw)
        supplier_invoice = xml_escape(supplier_invoice_raw)

        party = xml_escape(_ledger_or_suspense(_row_text(r, "PartyLedger")))
        purchase_raw = (
            _row_text(r, "PurchaseLedger")
            or _row_text(r, "PurchaseAccount")
            or _row_text(r, "Purchase Ledger")
            or _row_text(r, "ExpenseLedger")
            or _row_text(r, "SalesLedger")
        )
        purchase_raw = _ledger_or_suspense(purchase_raw)
        purchase = xml_escape(purchase_raw)

        taxable = _row_float(r, "TaxableValue", 0.0)
        cgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "CGSTLedger")))
        cgst_r = _row_float(r, "CGSTRate", 0.0)
        sgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "SGSTLedger")))
        sgst_r = _row_float(r, "SGSTRate", 0.0)
        igst_led = xml_escape(_ledger_or_suspense(_row_text(r, "IGSTLedger")))
        igst_r = _row_float(r, "IGSTRate", 0.0)
        narr = xml_escape(_row_text(r, "Narration"))
        party_gstin = xml_escape(_row_text(r, "PartyGSTIN") or _row_text(r, "GSTIN"))
        place_of_supply = xml_escape(_row_text(r, "PlaceOfSupply"))

        cgst_amt = round(taxable * cgst_r / 100, 2) if cgst_r > 0 else 0
        sgst_amt = round(taxable * sgst_r / 100, 2) if sgst_r > 0 else 0
        igst_amt = round(taxable * igst_r / 100, 2) if igst_r > 0 else 0
        total = taxable + cgst_amt + sgst_amt + igst_amt

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a('    <VOUCHER VCHTYPE="Purchase" ACTION="Create" OBJVIEW="Invoice Voucher View">')
        a(f'     <DATE>{dt}</DATE>')
        a('     <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>')
        a(f'     <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>')
        a(f'     <PARTYLEDGERNAME>{party}</PARTYLEDGERNAME>')
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>')
        if supplier_invoice:
            a(f'     <REFERENCE>{supplier_invoice}</REFERENCE>')
        if party_gstin:
            a(f'     <PARTYGSTIN>{party_gstin}</PARTYGSTIN>')
        if place_of_supply:
            a(f'     <PLACEOFSUPPLY>{place_of_supply}</PLACEOFSUPPLY>')
        if narr:
            a(f'     <NARRATION>{narr}</NARRATION>')

        # Party - Credit
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{party}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>{fmt_amt(total)}</AMOUNT>')
        a('      <BILLALLOCATIONS.LIST>')
        a(f'       <NAME>{supplier_invoice or vno}</NAME>')
        a('       <BILLTYPE>New Ref</BILLTYPE>')
        a(f'       <AMOUNT>{fmt_amt(total)}</AMOUNT>')
        a('      </BILLALLOCATIONS.LIST>')
        a('     </LEDGERENTRIES.LIST>')

        # Purchase - Debit
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{purchase}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>-{fmt_amt(taxable)}</AMOUNT>')
        a('     </LEDGERENTRIES.LIST>')

        if cgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{cgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(cgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        if sgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{sgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(sgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        if igst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{igst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(igst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')

        a('    </VOUCHER>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')
    return "\n".join(lines)


def generate_purchase_item_xml(
    rows: list,
    company: str,
    use_today_date: bool = False,
    start_voucher_number=None,
    fallback_purchase_ledger: str = SUSPENSE_LEDGER,
) -> str:
    """Purchase item invoice XML (inventory + accounting allocations)."""
    lines = []
    a = lines.append
    company_static = _company_static_block(company)
    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    def _name_key(value: str) -> str:
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    for idx, r in enumerate(rows):
        source_date = datetime.today() if use_today_date else _row_get(r, "Date", "")
        dt = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno = xml_escape(vno_raw)
        supplier_invoice_raw = _row_invoice_reference(r, vno_raw)
        supplier_invoice = xml_escape(supplier_invoice_raw)

        party = xml_escape(_ledger_or_suspense(_row_text(r, "PartyLedger")))
        taxable = _row_float(r, "TaxableValue", 0.0)

        item_name_raw = (
            _row_text(r, "ItemName")
            or _row_text(r, "Item")
            or _row_text(r, "StockItem")
            or _row_text(r, "ProductName")
            or _row_text(r, "PurchaseLedger")
            or _row_text(r, "SalesLedger")
        )
        if not item_name_raw:
            raise ValueError(f"Purchase item row {idx + 1}: item name is missing.")
        item_name = xml_escape(item_name_raw)
        item_name_key = _name_key(item_name_raw)

        qty = _row_float(r, "Quantity", 0.0)
        if qty <= 0:
            qty = _row_float(r, "Qty", 0.0)
        if qty <= 0:
            qty = _row_float(r, "Unit", 0.0)
        if qty <= 0:
            raise ValueError(f"Purchase item row {idx + 1}: quantity is missing/zero.")

        rate = _row_float(r, "Rate", 0.0)
        if rate <= 0 and taxable > 0 and qty > 0:
            rate = taxable / qty
        per_unit_raw = (
            _row_text(r, "Per", "")
            or _row_text(r, "UOM", "")
            or _row_text(r, "Unit", "")
            or "Nos"
        )
        per_unit = xml_escape(_normalize_stock_unit_name(per_unit_raw) or "Nos")
        godown = xml_escape(_row_text(r, "GodownName", "Main Location") or "Main Location")

        explicit_purchase_ledger = (
            _row_text(r, "PurchaseAccount")
            or _row_text(r, "Purchase Ledger")
            or _row_text(r, "ExpenseLedger")
            or _row_text(r, "PurchaseLedger")
        )
        default_purchase_ledger = _ledger_or_suspense(fallback_purchase_ledger)
        if _name_key(default_purchase_ledger) == item_name_key:
            for candidate in ("Purchase Account", "Purchase", "Purchase A/c", "Purchase Ledger"):
                if _name_key(candidate) != item_name_key:
                    default_purchase_ledger = candidate
                    break

        purchase_ledger_raw = (
            explicit_purchase_ledger
            or _row_text(r, "PurchaseLedger")
            or _row_text(r, "SalesLedger")
            or default_purchase_ledger
        )
        if _name_key(purchase_ledger_raw) == item_name_key:
            purchase_ledger_raw = default_purchase_ledger
        purchase_ledger_raw = _ledger_or_suspense(purchase_ledger_raw, default_purchase_ledger)
        if _name_key(purchase_ledger_raw) == item_name_key:
            raise ValueError(
                f"Purchase item row {idx + 1}: purchase ledger cannot match item '{item_name_raw}'."
            )
        purchase_ledger = xml_escape(purchase_ledger_raw)

        cgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "CGSTLedger")))
        cgst_r = _row_float(r, "CGSTRate", 0.0)
        sgst_led = xml_escape(_ledger_or_suspense(_row_text(r, "SGSTLedger")))
        sgst_r = _row_float(r, "SGSTRate", 0.0)
        igst_led = xml_escape(_ledger_or_suspense(_row_text(r, "IGSTLedger")))
        igst_r = _row_float(r, "IGSTRate", 0.0)
        narr = xml_escape(_row_text(r, "Narration"))
        party_gstin = xml_escape(_row_text(r, "PartyGSTIN") or _row_text(r, "GSTIN"))
        place_of_supply = xml_escape(_row_text(r, "PlaceOfSupply"))

        item_amt = round(qty * rate, 2) if qty and rate else taxable
        cgst_amt = round(taxable * cgst_r / 100, 2) if cgst_r > 0 else 0
        sgst_amt = round(taxable * sgst_r / 100, 2) if sgst_r > 0 else 0
        igst_amt = round(taxable * igst_r / 100, 2) if igst_r > 0 else 0
        total = taxable + cgst_amt + sgst_amt + igst_amt

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a('    <VOUCHER VCHTYPE="Purchase" ACTION="Create" OBJVIEW="Invoice Voucher View">')
        a(f'     <DATE>{dt}</DATE>')
        a('     <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>')
        a(f'     <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>')
        a(f'     <PARTYLEDGERNAME>{party}</PARTYLEDGERNAME>')
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>')
        if supplier_invoice:
            a(f'     <REFERENCE>{supplier_invoice}</REFERENCE>')
        if party_gstin:
            a(f'     <PARTYGSTIN>{party_gstin}</PARTYGSTIN>')
        if place_of_supply:
            a(f'     <PLACEOFSUPPLY>{place_of_supply}</PLACEOFSUPPLY>')
        if narr:
            a(f'     <NARRATION>{narr}</NARRATION>')

        a('     <ALLINVENTORYENTRIES.LIST>')
        a(f'      <STOCKITEMNAME>{item_name}</STOCKITEMNAME>')
        a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
        a(f'      <RATE>{fmt_amt(rate)}/{per_unit}</RATE>')
        a(f'      <AMOUNT>-{fmt_amt(item_amt)}</AMOUNT>')
        a(f'      <ACTUALQTY>{fmt_amt(qty)} {per_unit}</ACTUALQTY>')
        a(f'      <BILLEDQTY>{fmt_amt(qty)} {per_unit}</BILLEDQTY>')
        a('      <BATCHALLOCATIONS.LIST>')
        a(f'       <GODOWNNAME>{godown}</GODOWNNAME>')
        a(f'       <AMOUNT>-{fmt_amt(item_amt)}</AMOUNT>')
        a(f'       <ACTUALQTY>{fmt_amt(qty)} {per_unit}</ACTUALQTY>')
        a(f'       <BILLEDQTY>{fmt_amt(qty)} {per_unit}</BILLEDQTY>')
        a('      </BATCHALLOCATIONS.LIST>')
        a('      <ACCOUNTINGALLOCATIONS.LIST>')
        a(f'       <LEDGERNAME>{purchase_ledger}</LEDGERNAME>')
        a('       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
        a(f'       <AMOUNT>-{fmt_amt(item_amt)}</AMOUNT>')
        a('      </ACCOUNTINGALLOCATIONS.LIST>')
        a('     </ALLINVENTORYENTRIES.LIST>')

        # Party - Credit
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{party}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>{fmt_amt(total)}</AMOUNT>')
        a('      <BILLALLOCATIONS.LIST>')
        a(f'       <NAME>{supplier_invoice or vno}</NAME>')
        a('       <BILLTYPE>New Ref</BILLTYPE>')
        a(f'       <AMOUNT>{fmt_amt(total)}</AMOUNT>')
        a('      </BILLALLOCATIONS.LIST>')
        a('     </LEDGERENTRIES.LIST>')

        # Taxes - Debit
        if cgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{cgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(cgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        if sgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{sgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(sgst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')
        if igst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{igst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(igst_amt)}</AMOUNT>')
            a('     </LEDGERENTRIES.LIST>')

        a('    </VOUCHER>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════
#  GENERATE LEDGER MASTER XML
# ═══════════════════════════════════════════════════════════════════════════

def generate_ledger_xml(ledgers: list, company: str) -> str:
    """
    ledgers = list of dict: Name, Parent (group), GSTApplicable,
    GSTIN, StateOfSupply, TypeOfTaxation (e.g. 'Central Tax','State Tax','Integrated Tax')
    """
    lines = []
    a = lines.append
    company_static = _company_static_block(company)
    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>All Masters</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    for led in ledgers:
        name   = xml_escape(led["Name"])
        parent = xml_escape(led.get("Parent","Sundry Debtors"))
        gst_app = led.get("GSTApplicable","")
        gstin   = xml_escape(led.get("GSTIN",""))
        state   = xml_escape(led.get("StateOfSupply",""))
        tax_type= xml_escape(led.get("TypeOfTaxation",""))
        gst_rate= led.get("GSTRate","")

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a(f'    <LEDGER NAME="{name}" ACTION="Create">')
        a(f'     <NAME>{name}</NAME>')
        a(f'     <PARENT>{parent}</PARENT>')
        if gst_app:
            a(f'     <GSTAPPLICABLE>{xml_escape(gst_app)}</GSTAPPLICABLE>')
        if gstin:
            a(f'     <PARTYGSTIN>{gstin}</PARTYGSTIN>')
        if state:
            a(f'     <LEDSTATENAME>{state}</LEDSTATENAME>')
        if parent.lower() in ("duties & taxes","duties and taxes","duty"):
            if tax_type:
                a(f'     <TAXTYPE>{tax_type}</TAXTYPE>')
            if gst_rate:
                a(f'     <GSTRATE>{gst_rate}</GSTRATE>')
        a('    </LEDGER>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════
#  GENERATE STOCK ITEM MASTER XML
# ═══════════════════════════════════════════════════════════════════════════

def generate_stockitem_xml(items: list, company: str) -> str:
    """
    items = list of dict: Name, Parent (stock group), Unit, HSNCode,
    GSTRate, GSTApplicable, Description, OpeningQty, OpeningRate, OpeningValue
    """
    lines = []
    a = lines.append
    company_static = _company_static_block(company)

    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>All Masters</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    for item in items:
        name   = xml_escape(item["Name"])
        parent_raw = _normalize_stock_group_name(item.get("Parent", "Primary"))
        parent = xml_escape(parent_raw)
        unit_raw = _normalize_stock_unit_name(item.get("Unit", "Nos"))
        unit = xml_escape(unit_raw)
        hsn    = xml_escape(item.get("HSNCode",""))
        gst_r  = item.get("GSTRate","")
        gst_a  = item.get("GSTApplicable","Applicable")
        desc   = xml_escape(item.get("Description",""))

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a(f'    <STOCKITEM NAME="{name}" ACTION="Create">')
        a(f'     <NAME>{name}</NAME>')

        if parent_raw and parent_raw.casefold() != "primary":
            a(f'     <PARENT>{parent}</PARENT>')

        # ✅ FIX STARTS HERE
        a(f'     <BASEUNITS>{unit}</BASEUNITS>')
        a('     <ISADDITIONALUNITS>NO</ISADDITIONALUNITS>')
        # ❌ REMOVED: <ADDITIONALUNITS>
        # ✅ FIX ENDS HERE

        if hsn:
            a(f'     <GSTDETAILS.LIST>')
            a(f'      <HSNCODE>{hsn}</HSNCODE>')
            a(f'      <TAXABILITY>Taxable</TAXABILITY>')
            if gst_r:
                a(f'      <STATEWISEDETAILS.LIST>')
                a(f'       <RATEDETAILS.LIST>')
                a(f'        <GSTRATE>{gst_r}</GSTRATE>')
                a(f'       </RATEDETAILS.LIST>')
                a(f'      </STATEWISEDETAILS.LIST>')
            a(f'     </GSTDETAILS.LIST>')

        if gst_a:
            a(f'     <GSTAPPLICABLE>{xml_escape(gst_a)}</GSTAPPLICABLE>')

        if desc:
            a(f'     <DESCRIPTION>{desc}</DESCRIPTION>')

        a('    </STOCKITEM>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')

    return "\n".join(lines)

def generate_stockgroup_xml(
    group_names: list,
    company: str,
    default_parent: str = "",
    force_primary_parent: bool = False,
) -> str:
    """Create/alter stock groups so stock item parent groups exist before item import."""
    lines = []
    a = lines.append
    company_static = _company_static_block(company)
    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>All Masters</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    seen = set()
    for raw_name in group_names or []:
        group_name = str(raw_name or "").strip()
        if not group_name:
            continue
        key = group_name.casefold()
        if key in seen:
            continue
        seen.add(key)
        if key == "primary":
            continue

        name = xml_escape(group_name)
        parent_name = str(default_parent or "").strip()

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a(f'    <STOCKGROUP NAME="{name}" ACTION="Create Alter">')
        a(f'     <NAME>{name}</NAME>')
        # For many Tally setups, omitting PARENT creates group under Primary.
        # Keep explicit parent only when it's not the root Primary group.
        if parent_name and (parent_name.casefold() != "primary" or force_primary_parent):
            a(f'     <PARENT>{xml_escape(parent_name)}</PARENT>')
        a('    </STOCKGROUP>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')
    return "\n".join(lines)


def _normalize_stock_group_name(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return "Primary"

    # These are ledger groups, not inventory stock groups.
    ledger_like_groups = {
        "indirect income",
        "direct income",
        "indirect expenses",
        "direct expenses",
        "sales accounts",
        "purchase accounts",
        "sundry debtors",
        "sundry creditors",
        "duties & taxes",
        "duties and taxes",
        "bank accounts",
        "cash-in-hand",
        "cash in hand",
    }
    if text.casefold() in ledger_like_groups:
        return "Primary"
    return text


def _normalize_stock_unit_name(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return "Nos"

    # Keep user-entered unit unless it is a common alias.
    aliases = {
        "no": "Nos",
        "no.": "Nos",
        "nos": "Nos",
        "nos.": "Nos",
        "number": "Nos",
        "numbers": "Nos",
        "piece": "pcs",
        "pieces": "pcs",
    }
    return aliases.get(text.casefold(), text)


def _collect_required_stock_groups(items: list) -> list:
    required = []
    seen = set()
    for item in items or []:
        group_name = _normalize_stock_group_name(item.get("Parent", ""))
        if not group_name or group_name.casefold() == "primary":
            continue
        key = group_name.casefold()
        if key in seen:
            continue
        seen.add(key)
        required.append(group_name)
    return required


def _collect_required_stock_units(items: list) -> list:
    required = []
    seen = set()
    for item in items or []:
        unit_name = _normalize_stock_unit_name(item.get("Unit", "Nos"))
        if not unit_name:
            continue
        key = unit_name.casefold()
        if key in seen:
            continue
        seen.add(key)
        # "Not Applicable" behaves like built-in selection in many setups.
        if key in {"not applicable", "n/a", "na"}:
            continue
        required.append(unit_name)
    return required


def generate_unit_xml(unit_names: list, company: str) -> str:
    """Create/alter simple unit masters required for stock items."""
    lines = []
    a = lines.append
    company_static = _company_static_block(company)
    a('<?xml version="1.0" encoding="UTF-8"?>')
    a('<ENVELOPE>')
    a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>')
    a(' <BODY><IMPORTDATA>')
    a('  <REQUESTDESC><REPORTNAME>All Masters</REPORTNAME>')
    if company_static:
        a(company_static)
    a('  </REQUESTDESC>')
    a('  <REQUESTDATA>')

    seen = set()
    for raw in unit_names or []:
        unit_name = _normalize_stock_unit_name(raw)
        if not unit_name:
            continue
        key = unit_name.casefold()
        if key in seen:
            continue
        seen.add(key)
        if key in {"not applicable", "n/a", "na"}:
            continue

        unit_esc = xml_escape(unit_name)
        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a(f'    <UNIT NAME="{unit_esc}" ACTION="Create Alter">')
        a(f'     <NAME>{unit_esc}</NAME>')
        a('     <ISSIMPLEUNIT>Yes</ISSIMPLEUNIT>')
        a('     <DECIMALPLACES>2</DECIMALPLACES>')
        a(f'     <FORMALNAME>{unit_esc}</FORMALNAME>')
        a('    </UNIT>')
        a('   </TALLYMESSAGE>')

    a('  </REQUESTDATA>')
    a(' </IMPORTDATA></BODY>')
    a('</ENVELOPE>')
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════
#  READ EXCEL
# ═══════════════════════════════════════════════════════════════════════════

def read_excel(filepath: str, sheet: str = None) -> tuple:
    """Returns (headers: list, rows: list[dict])"""
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb[sheet] if sheet else wb.active
    header_rows = ws.iter_rows(min_row=1, max_row=1, values_only=True)
    first_row = next(header_rows, None) or []
    headers = [str(c or "").strip() for c in first_row]
    rows = []
    for vals in ws.iter_rows(min_row=2, values_only=True):
        vals = list(vals[:len(headers)])
        if len(vals) < len(headers):
            vals.extend([None] * (len(headers) - len(vals)))
        if all(v is None for v in vals):
            continue
        rows.append(dict(zip(headers, vals)))
    wb.close()
    return headers, rows


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════

class TallySalesApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("TallySalesPro — Sales Voucher & Master Creator")
        self.geometry("1120x720")
        self.minsize(960, 640)
        self.configure(fg_color=COLORS["bg_dark"])

        self.loaded_rows = []
        self.loaded_headers = []
        self.file_path_var = ctk.StringVar()
        self.company_placeholder = "Auto (Loaded Company)"
        self.company_var = ctk.StringVar(value=self.company_placeholder)
        self.tally_host_var = ctk.StringVar(value="localhost")
        self.tally_port_var = ctk.StringVar(value="9000")
        self.use_today_date_var = ctk.BooleanVar(value=False)
        self.status_var = ctk.StringVar(value="Ready")
        self.connection_status_var = ctk.StringVar(value="Connection: Not checked")
        self.company_status_var = ctk.StringVar(value="Companies: Not fetched")
        self.fetched_companies = []
        self._company_fetch_running = False
        self.debug_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tally_push_debug.log")
        self._voucher_info_labels = {}
        self._loaded_rows_by_mode = {}
        self._loaded_headers_by_mode = {}
        self._voucher_load_running = {}
        self._voucher_browse_buttons = {}
        self._voucher_save_buttons = {}
        self._voucher_push_buttons = {}
        self._voucher_template_buttons = {}
        self._push_running = False
        self._push_overlay = None
        self._push_message_var = ctk.StringVar(value="")
        self.template_names_by_mode = {
            "accounting": "Template_Sales Accounting Voucher.xlsx",
            "item": "Template_Sales Item Invoice.xlsx",
            "purchase_accounting": "Template_Purchase Accounting Voucher.xlsx",
            "purchase_item": "Template_Purchase Item Voucher.xlsx",
        }

        self._build_ui()

    def _resolve_theme_color(self, key: str):
        value = COLORS.get(key)
        if isinstance(value, tuple):
            mode = ctk.get_appearance_mode().lower()
            return value[1] if mode == "dark" else value[0]
        return value

    def _apply_ttk_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("default")
        except tk.TclError:
            pass

        tree_bg = self._resolve_theme_color("bg_card")
        field_bg = self._resolve_theme_color("bg_input")
        heading_bg = self._resolve_theme_color("table_header")
        text_fg = self._resolve_theme_color("text_primary")
        border_fg = self._resolve_theme_color("border")
        selected_bg = self._resolve_theme_color("accent")

        style.configure(
            "Treeview",
            background=tree_bg,
            foreground=text_fg,
            fieldbackground=field_bg,
            bordercolor=border_fg,
            font=("Segoe UI", 10),
            rowheight=26,
        )
        style.configure(
            "Treeview.Heading",
            background=heading_bg,
            foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", selected_bg)],
            foreground=[("selected", "#FFFFFF")],
        )

    def set_theme(self, mode: str):
        try:
            ctk.set_appearance_mode(mode)
        except Exception:
            pass
        self._apply_ttk_styles()

    # ── UI BUILDING ─────────────────────────────────────────────────────

    def _build_ui(self):
        self._apply_ttk_styles()

        settings_card = ctk.CTkFrame(
            self,
            fg_color=COLORS["bg_card"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=12,
        )
        settings_card.pack(fill="x", padx=16, pady=(10, 8))

        row_1 = ctk.CTkFrame(settings_card, fg_color="transparent")
        row_1.pack(fill="x", padx=14, pady=(12, 6))
        ctk.CTkLabel(row_1, text="Host", font=("Segoe UI", 10), text_color=COLORS["text_secondary"]).pack(side="left")
        ctk.CTkEntry(
            row_1,
            textvariable=self.tally_host_var,
            width=140,
            height=32,
            fg_color=COLORS["bg_input"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        ).pack(side="left", padx=(6, 12))
        ctk.CTkLabel(row_1, text="Port", font=("Segoe UI", 10), text_color=COLORS["text_secondary"]).pack(side="left")
        ctk.CTkEntry(
            row_1,
            textvariable=self.tally_port_var,
            width=90,
            height=32,
            fg_color=COLORS["bg_input"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        ).pack(side="left", padx=(6, 12))

        self.connection_test_btn = ctk.CTkButton(
            row_1,
            text="Test Connection",
            width=130,
            height=32,
            font=("Segoe UI", 10, "bold"),
            fg_color=COLORS["warning"],
            hover_color="#B45309",
            text_color="#FFFFFF",
            corner_radius=8,
            command=self._check_tally_connection_thread,
        )
        self.connection_test_btn.pack(side="right")

        row_2 = ctk.CTkFrame(settings_card, fg_color="transparent")
        row_2.pack(fill="x", padx=14, pady=(0, 6))
        ctk.CTkLabel(row_2, text="Target Company", font=("Segoe UI", 10), text_color=COLORS["text_secondary"]).pack(side="left")
        self.company_combo = ctk.CTkComboBox(
            row_2,
            values=[self.company_placeholder],
            variable=self.company_var,
            width=380,
            height=34,
            fg_color=COLORS["bg_input"],
            border_color=COLORS["border"],
            button_color=COLORS["accent"],
            button_hover_color=COLORS["accent_hover"],
            font=("Segoe UI", 10),
        )
        self.company_combo.set(self.company_placeholder)
        self.company_combo.pack(side="left", padx=(10, 8), fill="x", expand=True)
        self.company_refresh_btn = ctk.CTkButton(
            row_2,
            text="Refresh",
            width=96,
            height=34,
            font=("Segoe UI", 10, "bold"),
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"],
            corner_radius=8,
            command=self._fetch_tally_companies_thread,
        )
        self.company_refresh_btn.pack(side="right")

        status_row = ctk.CTkFrame(settings_card, fg_color="transparent")
        status_row.pack(fill="x", padx=14, pady=(0, 10))
        self.connection_status_label = ctk.CTkLabel(
            status_row,
            textvariable=self.connection_status_var,
            font=("Segoe UI", 10),
            text_color=COLORS["text_muted"],
        )
        self.connection_status_label.pack(anchor="w")
        self.company_status_label = ctk.CTkLabel(
            status_row,
            textvariable=self.company_status_var,
            font=("Segoe UI", 10),
            text_color=COLORS["text_muted"],
        )
        self.company_status_label.pack(anchor="w")

        self.today_date_checkbox = ctk.CTkCheckBox(
            settings_card,
            text="Use Today Date For Vouchers (ignore Excel Date)",
            variable=self.use_today_date_var,
            font=("Segoe UI", 10, "bold"),
            text_color=COLORS["text_secondary"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            border_color=COLORS["border"],
        )
        self.today_date_checkbox.pack(anchor="w", padx=14, pady=(0, 10))

        self.tabs = ctk.CTkTabview(
            self,
            corner_radius=10,
            fg_color=COLORS["bg_card"],
            border_width=1,
            border_color=COLORS["border"],
            segmented_button_fg_color=COLORS["bg_input"],
            segmented_button_selected_color=COLORS["accent"],
            segmented_button_selected_hover_color=COLORS["accent_hover"],
            segmented_button_unselected_color=COLORS["bg_input"],
            segmented_button_unselected_hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_primary"],
            text_color_disabled=COLORS["text_muted"],
        )
        self.tabs.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        self.tab_acct = self.tabs.add("📋 Sales Accounting Invoice")
        self.tab_item = self.tabs.add("📦 Sales Item Invoice")
        self.tab_purchase_acct = self.tabs.add("🧾 Purchase Accounting Invoice")
        self.tab_purchase_item = self.tabs.add("🛒 Purchase Item Invoice")
        self.tab_ledger = self.tabs.add("🏦 Create Ledgers")
        self.tab_stock = self.tabs.add("📁 Create Stock Items")

        self._build_voucher_tab(self.tab_acct, mode="accounting")
        self._build_voucher_tab(self.tab_item, mode="item")
        self._build_voucher_tab(self.tab_purchase_acct, mode="purchase_accounting")
        self._build_voucher_tab(self.tab_purchase_item, mode="purchase_item")
        self._build_ledger_tab()
        self._build_stock_tab()

        status_bar = ctk.CTkFrame(self, fg_color=COLORS["bg_card"], corner_radius=0, height=32)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)
        ctk.CTkLabel(
            status_bar,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            text_color=COLORS["text_muted"],
        ).pack(side="left", padx=16)

        self.after(200, lambda: self._fetch_tally_companies_thread(silent=True))

    def _get_tally_url(self):
        return _build_tally_url(self.tally_host_var.get(), self.tally_port_var.get())

    def _append_debug_log(self, mode, target_company, xml_payload, response_text, parsed, note=""):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [
            "=" * 96,
            f"[{timestamp}] mode={mode} company={target_company or 'Loaded company in Tally'} note={note}",
            (
                "summary: "
                f"created={parsed.get('created', 0)} "
                f"altered={parsed.get('altered', 0)} "
                f"deleted={parsed.get('deleted', 0)} "
                f"ignored={parsed.get('ignored', 0)} "
                f"errors={parsed.get('errors', 0)} "
                f"exceptions={parsed.get('exceptions', 0)}"
            ),
        ]
        line_errors = parsed.get("line_errors") or []
        if line_errors:
            lines.append("line_errors:")
            for err in line_errors:
                lines.append(f"- {err}")
        if parsed.get("error"):
            lines.append(f"parsed_error: {parsed.get('error')}")

        lines.append("response:")
        lines.append(response_text[:12000])
        lines.append("xml:")
        lines.append(xml_payload[:12000])
        lines.append("\n")

        with open(self.debug_log_path, "a", encoding="utf-8") as log_file:
            log_file.write("\n".join(lines))

    def _get_selected_company(self):
        selected = _normalize_company_name(self.company_var.get())
        if not selected or _company_key(selected) == _company_key(self.company_placeholder):
            if len(getattr(self, "fetched_companies", [])) == 1:
                return self.fetched_companies[0]
            return ""
        return selected

    def _set_company_dropdown(self, companies, keep_selection=True):
        current = _normalize_company_name(self.company_var.get()) if keep_selection else ""
        cleaned = []
        seen = set()
        for name in companies or []:
            normalized = _normalize_company_name(name)
            if not _is_valid_company_name(normalized):
                continue
            key = _company_key(normalized)
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(normalized)

        cleaned = sorted(cleaned, key=lambda x: _company_key(x))
        values = [self.company_placeholder] + cleaned
        self.company_combo.configure(values=values)

        if current and _company_key(current) in {_company_key(x) for x in cleaned}:
            self.company_combo.set(current)
            self.company_var.set(current)
        else:
            self.company_combo.set(self.company_placeholder)
            self.company_var.set(self.company_placeholder)

        self.fetched_companies = cleaned
        self.company_status_var.set(f"Companies: {len(cleaned)} available")
        self.company_status_label.configure(text_color=COLORS["text_muted"])

    def _fetch_tally_companies_thread(self, silent=False):
        if self._company_fetch_running:
            return
        try:
            tally_url = self._get_tally_url()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc))
            return

        self._company_fetch_running = True
        self.company_refresh_btn.configure(state="disabled", text="Fetching...")
        if not silent:
            self.company_status_var.set("Companies: Fetching...")
            self.company_status_label.configure(text_color=COLORS["warning"])

        def worker():
            result = _fetch_tally_companies(tally_url, timeout=15)

            def done():
                self._company_fetch_running = False
                self.company_refresh_btn.configure(state="normal", text="Refresh")
                if result.get("success"):
                    companies = result.get("companies", [])
                    self._set_company_dropdown(companies, keep_selection=True)
                    self.status_var.set(f"Fetched {len(companies)} company(s) from Tally")
                else:
                    err = str(result.get("error") or "Unknown error")
                    self.company_status_var.set("Companies: Fetch failed")
                    self.company_status_label.configure(text_color=COLORS["error"])
                    self.status_var.set("Company fetch failed")
                    if not silent:
                        messagebox.showwarning("Company Fetch Failed", f"Could not fetch companies from Tally.\n\n{err}")

            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _check_tally_connection_thread(self):
        try:
            tally_url = self._get_tally_url()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc))
            return

        self.connection_test_btn.configure(state="disabled", text="Checking...")
        self.connection_status_var.set("Connection: Checking...")
        self.connection_status_label.configure(text_color=COLORS["warning"])

        def worker():
            result = _check_tally_connection(tally_url, timeout=8)

            def done():
                self.connection_test_btn.configure(state="normal", text="Test Connection")
                if result.get("connected"):
                    self.connection_status_var.set(f"Connection: Connected ({tally_url})")
                    self.connection_status_label.configure(text_color=COLORS["success"])
                    self.status_var.set("Connected to Tally")
                    self._fetch_tally_companies_thread(silent=True)
                else:
                    err = str(result.get("error") or "Unknown error")
                    self.connection_status_var.set("Connection: Offline")
                    self.connection_status_label.configure(text_color=COLORS["error"])
                    self.status_var.set(f"Connection failed: {err}")
                    messagebox.showwarning("Connection Failed", f"Could not connect to Tally.\n\n{err}")

            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    # ── VOUCHER TAB (shared for accounting & item) ──────────────────────

    def _build_voucher_tab(self, parent, mode="accounting"):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        # File load row
        load_frame = ctk.CTkFrame(parent, fg_color="transparent")
        load_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))

        fp_var = ctk.StringVar()
        ctk.CTkEntry(load_frame, textvariable=fp_var, placeholder_text="Select Excel file (.xlsx / .xlsm)...",
                      width=500, state="readonly").pack(side="left", padx=(0,8))

        def browse():
            if self._voucher_load_running.get(mode):
                self.status_var.set("Please wait, file is still loading...")
                return
            f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm *.xls")])
            if f:
                fp_var.set(f)
                self._load_preview(f, tree, mode)
        browse_btn = ctk.CTkButton(
            load_frame,
            text="Browse",
            command=browse,
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
        )
        browse_btn.pack(side="left", padx=(0,8))
        self._voucher_browse_buttons[mode] = browse_btn
        self._voucher_load_running[mode] = False

        if mode in {"item", "purchase_item"}:
            ctk.CTkLabel(load_frame, text="⚠ Requires: ItemName, Quantity, Rate, and Per or Unit/UOM columns",
                          font=("Segoe UI", 11), text_color=COLORS["warning"]).pack(side="left")

        # Preview table
        tree_frame = ctk.CTkFrame(
            parent,
            fg_color=COLORS["bg_dark"],
            corner_radius=8,
            border_width=1,
            border_color=COLORS["border"],
        )
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree = ttk.Treeview(tree_frame, show="headings",
                             yscrollcommand=tree_scroll_y.set,
                             xscrollcommand=tree_scroll_x.set)
        tree_scroll_y.config(command=tree.yview)
        tree_scroll_x.config(command=tree.xview)
        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)

        # Action buttons
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 10))

        save_btn = ctk.CTkButton(
            btn_frame,
            text="💾  Save XML File",
            fg_color=SUCCESS,
            hover_color="#15803D",
            width=160,
            command=lambda: self._generate(mode, "save", fp_var.get()),
        )
        save_btn.pack(side="left", padx=(0,10))
        self._voucher_save_buttons[mode] = save_btn

        template_btn = ctk.CTkButton(
            btn_frame,
            text="📥  Download Template",
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"],
            width=170,
            command=lambda: self._download_template_for_mode(mode),
        )
        template_btn.pack(side="left", padx=(0,10))
        self._voucher_template_buttons[mode] = template_btn

        push_btn = ctk.CTkButton(
            btn_frame,
            text="🚀  Push to Tally",
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            width=160,
            command=lambda: self._generate(mode, "push", fp_var.get()),
        )
        push_btn.pack(side="left", padx=(0,10))
        self._voucher_push_buttons[mode] = push_btn

        lbl = ctk.CTkLabel(btn_frame, text="", font=("Segoe UI", 11), text_color=TEXT_MUTED)
        lbl.pack(side="left", padx=10)
        self._voucher_info_labels[mode] = lbl

    def _set_voucher_loading_state(self, mode: str, is_loading: bool):
        self._voucher_load_running[mode] = is_loading
        state = "disabled" if is_loading else "normal"
        browse_text = "Loading..." if is_loading else "Browse"
        browse_btn = self._voucher_browse_buttons.get(mode)
        if browse_btn:
            browse_btn.configure(state=state, text=browse_text)
        save_btn = self._voucher_save_buttons.get(mode)
        if save_btn:
            save_btn.configure(state=state)
        template_btn = self._voucher_template_buttons.get(mode)
        if template_btn:
            template_btn.configure(state=state)
        push_btn = self._voucher_push_buttons.get(mode)
        if push_btn:
            push_btn.configure(state=state)

    def _set_push_loading_state(self, is_loading: bool, message: str = ""):
        self._push_running = is_loading
        state = "disabled" if is_loading else "normal"

        for btn in self._voucher_browse_buttons.values():
            btn.configure(state=state)
        for btn in self._voucher_save_buttons.values():
            btn.configure(state=state)
        for btn in self._voucher_template_buttons.values():
            btn.configure(state=state)
        for btn in self._voucher_push_buttons.values():
            btn.configure(state=state)

        if is_loading:
            self._push_message_var.set(message or "Posting to Tally...")
            if self._push_overlay is None or not self._push_overlay.winfo_exists():
                overlay = ctk.CTkToplevel(self)
                overlay.title("Please Wait")
                overlay.geometry("420x170")
                overlay.transient(self)
                overlay.grab_set()
                overlay.resizable(False, False)

                frame = ctk.CTkFrame(overlay, fg_color=COLORS["bg_card"], corner_radius=10)
                frame.pack(fill="both", expand=True, padx=12, pady=12)

                ctk.CTkLabel(
                    frame,
                    text="Pushing Entries To Tally",
                    font=("Segoe UI", 15, "bold"),
                    text_color=COLORS["text_primary"],
                ).pack(pady=(12, 8))

                ctk.CTkLabel(
                    frame,
                    textvariable=self._push_message_var,
                    font=("Segoe UI", 11),
                    text_color=COLORS["text_muted"],
                    wraplength=360,
                    justify="center",
                ).pack(pady=(0, 10))

                progress = ctk.CTkProgressBar(frame, mode="indeterminate", width=320)
                progress.pack(pady=(0, 10))
                progress.start()

                overlay.protocol("WM_DELETE_WINDOW", lambda: None)
                self._push_overlay = overlay
            else:
                self._push_overlay.deiconify()
                self._push_overlay.lift()
        else:
            if self._push_overlay is not None and self._push_overlay.winfo_exists():
                try:
                    self._push_overlay.grab_release()
                except Exception:
                    pass
                self._push_overlay.destroy()
            self._push_overlay = None
            self._push_message_var.set("")

        self.update_idletasks()

    def _get_template_definition(self, mode: str) -> dict:
        templates = {
            "accounting": {
                "sheet_name": "Sheet1",
                "headers": [
                    "Date",
                    "VoucherNo",
                    "GSTIN",
                    "PartyLedger",
                    "SalesLedger",
                    "TaxableValue",
                    "CGSTLedger",
                    "CGSTRate",
                    "SGSTLedger",
                    "SGSTRate",
                    "IGSTLedger",
                    "IGSTRate",
                    "Narration",
                ],
                "sample_rows": [],
            },
            "item": {
                "sheet_name": "Sheet1",
                "headers": [
                    "Date",
                    "VoucherNo",
                    "GSTIN",
                    "PartyLedger",
                    "Sales Ledger",
                    "Item Name",
                    "Unit",
                    "Quantity",
                    "Rate",
                    "TaxableValue",
                    "CGSTLedger",
                    "CGSTRate",
                    "SGSTLedger",
                    "SGSTRate",
                    "IGSTLedger",
                    "IGSTRate",
                    "Narration",
                ],
                "sample_rows": [],
            },
            "purchase_accounting": {
                "sheet_name": "Sheet1",
                "headers": [
                    "Date",
                    "VoucherNo",
                    "GSTIN",
                    "PartyLedger",
                    "Purchase Ledger",
                    "TaxableValue",
                    "CGSTLedger",
                    "CGSTRate",
                    "SGSTLedger",
                    "SGSTRate",
                    "IGSTLedger",
                    "IGSTRate",
                    "Narration",
                ],
                "sample_rows": [],
            },
            "purchase_item": {
                "sheet_name": "Sheet1",
                "headers": [
                    "Date",
                    "Invoice No",
                    "PartyLedger",
                    "Purchase Ledger",
                    "Item Name",
                    "Unit",
                    "Quantity",
                    "Rate",
                    "TaxableValue",
                    "CGSTLedger",
                    "CGSTRate",
                    "SGSTLedger",
                    "SGSTRate",
                    "IGSTLedger",
                    "IGSTRate",
                    "Narration",
                ],
                "sample_rows": [],
            },
        }
        return templates.get(mode, {})

    def _download_template_for_mode(self, mode: str):
        template_name = self.template_names_by_mode.get(mode, "Template.xlsx")
        template_definition = self._get_template_definition(mode)
        if not template_definition:
            messagebox.showwarning("Template Not Found", "No template is configured for this mode.")
            return

        out_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=template_name,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not out_path:
            return

        try:
            headers = list(template_definition.get("headers", []))
            sample_rows = list(template_definition.get("sample_rows", []))
            sheet_name = str(template_definition.get("sheet_name") or "Template")

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet_name

            for col_index, header in enumerate(headers, 1):
                ws.cell(row=1, column=col_index, value=header)

            for row_index, row_values in enumerate(sample_rows, 2):
                for col_index, value in enumerate(row_values, 1):
                    ws.cell(row=row_index, column=col_index, value=value)

            wb.save(out_path)
            wb.close()
            self.status_var.set(f"Template saved: {os.path.basename(out_path)}")
            messagebox.showinfo("Template Saved", f"Template saved successfully!\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Template Error", str(exc))

    def _load_preview(self, filepath, tree, mode):
        self._set_voucher_loading_state(mode, True)
        info_label = self._voucher_info_labels.get(mode)
        if info_label:
            info_label.configure(text="Loading Excel preview...")
        self.status_var.set(f"Loading: {os.path.basename(filepath)}")

        def worker():
            try:
                headers, rows = read_excel(filepath)
                preview_rows = []
                for r in rows[:300]:
                    values = []
                    for h in headers:
                        cell_value = _row_get(r, h, "")
                        values.append("" if cell_value is None else str(cell_value))
                    preview_rows.append(values)
                result = {
                    "ok": True,
                    "headers": headers,
                    "rows": rows,
                    "preview_rows": preview_rows,
                }
            except Exception as exc:
                result = {"ok": False, "error": str(exc)}

            def done():
                self._set_voucher_loading_state(mode, False)
                if not result.get("ok"):
                    if info_label:
                        info_label.configure(text="")
                    self.status_var.set("Ready")
                    messagebox.showerror("Error", str(result.get("error", "Unknown error")))
                    return

                headers = result["headers"]
                rows = result["rows"]
                self.loaded_headers = headers
                self.loaded_rows = rows
                self._loaded_headers_by_mode[mode] = headers
                self._loaded_rows_by_mode[mode] = rows

                tree.delete(*tree.get_children())
                tree["columns"] = headers
                for h in headers:
                    tree.heading(h, text=h)
                    tree.column(h, width=100, minwidth=60)

                for values in result["preview_rows"]:
                    tree.insert("", "end", values=values)

                info = f"✅ Loaded {len(rows)} rows, {len(headers)} columns"
                if info_label:
                    info_label.configure(text=info)
                self.status_var.set(f"Loaded: {os.path.basename(filepath)} — {len(rows)} rows")

            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _generate(self, mode, action, filepath):
        _ = filepath
        if self._voucher_load_running.get(mode):
            messagebox.showinfo("Please Wait", "Excel file is still loading. Please try again in a moment.")
            return

        if action == "push" and self._push_running:
            self.status_var.set("Push already in progress. Please wait...")
            return

        mode_rows = self._loaded_rows_by_mode.get(mode)
        rows_to_use = mode_rows if mode_rows is not None else self.loaded_rows
        company = self._get_selected_company()
        use_today_date = bool(self.use_today_date_var.get())
        if not rows_to_use:
            messagebox.showwarning("No Data", "Load an Excel file first.")
            return

        if action == "push" and not company and len(getattr(self, "fetched_companies", [])) > 1:
            self.status_var.set("Select target company before push")
            messagebox.showwarning(
                "Select Target Company",
                "Multiple companies were detected in Tally.\n"
                "Please select the target company from the dropdown before pushing.",
            )
            return

        mode_to_voucher_type = {
            "accounting": "Sales",
            "item": "Sales",
            "purchase_accounting": "Purchase",
            "purchase_item": "Purchase",
        }
        mode_to_file_name = {
            "accounting": "Tally_Sales_Accounting.xml",
            "item": "Tally_Sales_Item.xml",
            "purchase_accounting": "Tally_Purchase_Accounting.xml",
            "purchase_item": "Tally_Purchase_Item.xml",
        }

        def build_voucher_xml(selected_mode: str, today_flag: bool, voucher_start, rows_data):
            if selected_mode == "accounting":
                return generate_accounting_xml(
                    rows_data,
                    company,
                    use_today_date=today_flag,
                    start_voucher_number=voucher_start,
                )
            if selected_mode == "item":
                return generate_item_xml(
                    rows_data,
                    company,
                    use_today_date=today_flag,
                    start_voucher_number=voucher_start,
                )
            if selected_mode == "purchase_accounting":
                return generate_purchase_accounting_xml(
                    rows_data,
                    company,
                    use_today_date=today_flag,
                    start_voucher_number=voucher_start,
                )
            if selected_mode == "purchase_item":
                return generate_purchase_item_xml(
                    rows_data,
                    company,
                    use_today_date=today_flag,
                    start_voucher_number=voucher_start,
                )
            raise ValueError(f"Unsupported mode: {selected_mode}")

        try:
            if action == "save":
                xml = build_voucher_xml(mode, use_today_date, None, rows_to_use)
                out = filedialog.asksaveasfilename(defaultextension=".xml",
                        filetypes=[("XML","*.xml")],
                        initialfile=mode_to_file_name.get(mode, "Tally_Voucher.xml"))
                if out:
                    with open(out, "w", encoding="utf-8") as f:
                        f.write(xml)
                    self.status_var.set(f"Saved XML: {out}")
                    messagebox.showinfo("Success", f"XML saved successfully!\n{out}")
            else:
                tally_url = self._get_tally_url()
                host, port_text = tally_url.rsplit(":", 1)
                host = host.replace("http://", "", 1)
                port_value = int(port_text)
                target_company = company or "Loaded company in Tally"
                date_mode = "today date" if use_today_date else "excel date"
                voucher_type = mode_to_voucher_type.get(mode, "Sales")
                rows_snapshot = list(rows_to_use)

                self._set_push_loading_state(True, f"Preparing vouchers for {target_company}...")
                self.status_var.set(f"Posting to Tally ({target_company}, {date_mode})...")

                def worker():
                    result = {
                        "ok": False,
                        "target_company": target_company,
                        "parsed": {},
                        "message": "",
                        "detail": "",
                    }
                    try:
                        effective_today_date = use_today_date

                        auto_ledger_defs = _collect_auto_voucher_ledgers(rows_snapshot, mode)
                        if auto_ledger_defs:
                            self.after(0, lambda: self._push_message_var.set("Creating required ledgers in Tally..."))
                            auto_ledger_xml = generate_ledger_xml(auto_ledger_defs, company)
                            auto_ledger_resp = push_to_tally(auto_ledger_xml, host, port_value)
                            auto_ledger_parsed = _parse_tally_response_details(auto_ledger_resp)
                            self._append_debug_log(
                                "auto-ledger",
                                target_company,
                                auto_ledger_xml,
                                auto_ledger_resp,
                                auto_ledger_parsed,
                                note=f"mode={mode}, ledgers={len(auto_ledger_defs)}",
                            )

                        self.after(0, lambda: self._push_message_var.set("Fetching next voucher number from Tally..."))
                        next_voucher = None
                        voucher_note = "excel_voucher_no"
                        voucher_result = _fetch_next_voucher_number(
                            tally_url,
                            company_name=company,
                            voucher_type=voucher_type,
                            timeout=15,
                        )
                        if voucher_result.get("success"):
                            next_voucher = voucher_result.get("next_number")
                            voucher_note = f"auto_vno_start={next_voucher}"
                        else:
                            voucher_note = f"auto_vno_fetch_failed={voucher_result.get('error', 'Unknown')}"

                        vno_label = f"vno start {next_voucher}" if next_voucher is not None else "vno from excel"
                        self.after(
                            0,
                            lambda: self.status_var.set(
                                f"Posting to Tally ({target_company}, {date_mode}, {vno_label})..."
                            ),
                        )
                        self.after(0, lambda: self._push_message_var.set("Posting voucher data to Tally..."))

                        xml = build_voucher_xml(mode, use_today_date, next_voucher, rows_snapshot)
                        resp = push_to_tally(xml, host, port_value)
                        parsed = _parse_tally_response_details(resp)
                        self._append_debug_log(
                            mode,
                            target_company,
                            xml,
                            resp,
                            parsed,
                            note=f"voucher_type={voucher_type}, date_mode={date_mode}, {voucher_note}",
                        )

                        if parsed.get("success"):
                            result = {
                                "ok": True,
                                "target_company": target_company,
                                "parsed": parsed,
                                "message": (
                                    "Posted to Tally successfully.\n\n"
                                    f"Target Company: {target_company}\n"
                                    f"Created: {parsed.get('created', 0)}\n"
                                    f"Altered: {parsed.get('altered', 0)}\n"
                                    f"Ignored: {parsed.get('ignored', 0)}"
                                ),
                            }
                        else:
                            if not use_today_date:
                                self.after(0, lambda: self._push_message_var.set("Retrying with today date..."))
                                retry_xml = build_voucher_xml(mode, True, next_voucher, rows_snapshot)
                                retry_resp = push_to_tally(retry_xml, host, port_value)
                                retry_parsed = _parse_tally_response_details(retry_resp)
                                self._append_debug_log(
                                    mode,
                                    target_company,
                                    retry_xml,
                                    retry_resp,
                                    retry_parsed,
                                    note=f"voucher_type={voucher_type}, auto_retry_today_date, {voucher_note}",
                                )
                                if retry_parsed.get("success"):
                                    result = {
                                        "ok": True,
                                        "target_company": target_company,
                                        "parsed": retry_parsed,
                                        "message": (
                                            "Initial push failed with Excel date; auto-retry with today date succeeded.\n\n"
                                            f"Target Company: {target_company}\n"
                                            f"Created: {retry_parsed.get('created', 0)}\n"
                                            f"Altered: {retry_parsed.get('altered', 0)}\n"
                                            f"Debug Log: {self.debug_log_path}"
                                        ),
                                    }
                                else:
                                    parsed = retry_parsed
                                    effective_today_date = True

                            if not result.get("ok"):
                                missing_ledger_defs = _build_missing_ledger_defs(
                                    parsed.get("line_errors") or [],
                                    rows_snapshot,
                                    mode,
                                )
                                if missing_ledger_defs:
                                    self.after(0, lambda: self._push_message_var.set("Creating missing ledgers and retrying..."))
                                    missing_ledger_xml = generate_ledger_xml(missing_ledger_defs, company)
                                    missing_ledger_resp = push_to_tally(missing_ledger_xml, host, port_value)
                                    missing_ledger_parsed = _parse_tally_response_details(missing_ledger_resp)
                                    self._append_debug_log(
                                        "missing-ledger",
                                        target_company,
                                        missing_ledger_xml,
                                        missing_ledger_resp,
                                        missing_ledger_parsed,
                                        note=f"mode={mode}, ledgers={len(missing_ledger_defs)}",
                                    )

                                    blockers = []
                                    for err in missing_ledger_parsed.get("line_errors", []):
                                        err_low = err.lower()
                                        if "already exists" in err_low or "already present" in err_low:
                                            continue
                                        blockers.append(err)

                                    if not blockers:
                                        self.after(0, lambda: self._push_message_var.set("Retrying voucher post after ledger creation..."))
                                        post_ledger_retry_xml = build_voucher_xml(
                                            mode,
                                            effective_today_date,
                                            next_voucher,
                                            rows_snapshot,
                                        )
                                        post_ledger_retry_resp = push_to_tally(post_ledger_retry_xml, host, port_value)
                                        post_ledger_retry_parsed = _parse_tally_response_details(post_ledger_retry_resp)
                                        self._append_debug_log(
                                            mode,
                                            target_company,
                                            post_ledger_retry_xml,
                                            post_ledger_retry_resp,
                                            post_ledger_retry_parsed,
                                            note=(
                                                f"voucher_type={voucher_type}, auto_retry_missing_ledgers, {voucher_note}"
                                            ),
                                        )
                                        if post_ledger_retry_parsed.get("success"):
                                            result = {
                                                "ok": True,
                                                "target_company": target_company,
                                                "parsed": post_ledger_retry_parsed,
                                                "message": (
                                                    "Auto-created missing ledgers and posted to Tally successfully.\n\n"
                                                    f"Target Company: {target_company}\n"
                                                    f"Created: {post_ledger_retry_parsed.get('created', 0)}\n"
                                                    f"Altered: {post_ledger_retry_parsed.get('altered', 0)}\n"
                                                    f"Ignored: {post_ledger_retry_parsed.get('ignored', 0)}"
                                                ),
                                            }
                                        else:
                                            parsed = post_ledger_retry_parsed

                            if not result.get("ok"):
                                detail = parsed.get("error") or "Unknown Tally exception"
                                line_errors = parsed.get("line_errors") or []
                                if line_errors:
                                    detail = line_errors[0]
                                result = {
                                    "ok": False,
                                    "target_company": target_company,
                                    "parsed": parsed,
                                    "detail": detail,
                                }

                    except Exception as exc:
                        result = {
                            "ok": False,
                            "target_company": target_company,
                            "parsed": {
                                "created": 0,
                                "altered": 0,
                                "deleted": 0,
                                "ignored": 0,
                                "errors": 1,
                                "exceptions": 1,
                                "line_errors": [],
                                "error": str(exc),
                            },
                            "detail": str(exc),
                        }
                        try:
                            self._append_debug_log(mode, company or "", "", str(exc), result["parsed"], note="python_exception")
                        except Exception:
                            pass

                    def done():
                        self._set_push_loading_state(False)
                        if result.get("ok"):
                            parsed_local = result.get("parsed", {})
                            self.status_var.set(f"Posted to Tally ({result.get('target_company', target_company)})")
                            messagebox.showinfo("Tally Response", result.get("message", "Posted to Tally."))
                            return

                        parsed_local = result.get("parsed", {})
                        detail = result.get("detail") or result.get("error") or "Unknown Tally exception"
                        self.status_var.set("Push failed (see debug log)")
                        messagebox.showerror(
                            "Push Failed",
                            "Tally returned an exception/error while importing voucher.\n\n"
                            f"Target Company: {result.get('target_company', target_company)}\n"
                            f"Errors: {parsed_local.get('errors', 0)}\n"
                            f"Exceptions: {parsed_local.get('exceptions', 0)}\n"
                            f"Detail: {detail}\n\n"
                            f"Debug Log Saved: {self.debug_log_path}",
                        )

                    self.after(0, done)

                threading.Thread(target=worker, daemon=True).start()
                return
        except ValueError as e:
            messagebox.showerror("Invalid Settings", str(e))
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            try:
                parsed = {
                    "created": 0,
                    "altered": 0,
                    "deleted": 0,
                    "ignored": 0,
                    "errors": 1,
                    "exceptions": 1,
                    "line_errors": [],
                    "error": str(e),
                }
                self._append_debug_log(mode, company or "", "", str(e), parsed, note="python_exception")
            except Exception:
                pass
            messagebox.showerror("Error", str(e))

    # ── LEDGER CREATION TAB ─────────────────────────────────────────────

    def _build_ledger_tab(self):
        parent = self.tab_ledger

        info = ctk.CTkLabel(parent, text="Create Party / Sales / Tax Ledgers in TallyPrime",
                             font=("Segoe UI", 13, "bold"), text_color=TEXT_PRIMARY)
        info.pack(pady=(10,5))

        main = ctk.CTkFrame(parent, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=10, pady=5)
        main.grid_columnconfigure(0, weight=1, minsize=300)
        main.grid_columnconfigure(1, weight=2)
        main.grid_rowconfigure(0, weight=1)

        # Left: Form
        form = ctk.CTkScrollableFrame(
            main,
            fg_color=COLORS["bg_input"],
            corner_radius=10,
            width=360,
            border_width=1,
            border_color=COLORS["border"],
        )
        form.grid(row=0, column=0, sticky="nsew", padx=(0,10), pady=0)
        form.grid_columnconfigure(0, weight=1)

        fields = {}
        configs = [
            ("Ledger Name *", "led_name", "e.g. ABC Traders"),
            ("Parent Group *", "led_parent", "Sundry Debtors / Sales Accounts / Duties & Taxes"),
            ("GST Applicable", "led_gst_app", "Applicable / Not Applicable"),
            ("GSTIN", "led_gstin", "e.g. 07AAACR1718Q1ZZ"),
            ("State", "led_state", "e.g. Delhi"),
            ("Tax Type", "led_tax_type", "Central Tax / State Tax / Integrated Tax"),
            ("GST Rate %", "led_gst_rate", "e.g. 9"),
        ]
        for label, key, placeholder in configs:
            ctk.CTkLabel(form, text=label, font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
            e = ctk.CTkEntry(
                form,
                placeholder_text=placeholder,
                fg_color=COLORS["bg_card"],
                border_color=COLORS["border"],
                text_color=COLORS["text_primary"],
            )
            e.pack(fill="x", padx=12, pady=(0,2))
            fields[key] = e

        self._ledger_list = []
        ledger_edit_index = None

        def add_ledger():
            nonlocal ledger_edit_index
            name = fields["led_name"].get().strip()
            parent_grp = fields["led_parent"].get().strip()
            if not name or not parent_grp:
                messagebox.showwarning("Required","Ledger Name and Parent Group are required.")
                return
            entry = {
                "Name": name, "Parent": parent_grp,
                "GSTApplicable": fields["led_gst_app"].get().strip(),
                "GSTIN": fields["led_gstin"].get().strip(),
                "StateOfSupply": fields["led_state"].get().strip(),
                "TypeOfTaxation": fields["led_tax_type"].get().strip(),
                "GSTRate": fields["led_gst_rate"].get().strip(),
            }
            row_values = (name, parent_grp, entry["GSTApplicable"], entry["GSTIN"], entry["GSTRate"])

            if ledger_edit_index is None:
                self._ledger_list.append(entry)
                led_tree.insert("", "end", values=row_values)
            else:
                self._ledger_list[ledger_edit_index] = entry
                item_ids = led_tree.get_children()
                if ledger_edit_index < len(item_ids):
                    led_tree.item(item_ids[ledger_edit_index], values=row_values)
                ledger_edit_index = None
                add_ledger_btn.configure(text="➕  Add to Queue")

            for v in fields.values():
                v.delete(0, "end")
            self._led_count_label.configure(text=f"{len(self._ledger_list)} ledger(s) queued")

        def edit_selected_ledger():
            nonlocal ledger_edit_index
            selected = led_tree.selection()
            if not selected:
                messagebox.showwarning("Select Row", "Select a queued ledger row to edit.")
                return

            item_id = selected[0]
            idx = led_tree.index(item_id)
            if idx >= len(self._ledger_list):
                messagebox.showwarning("Selection Error", "Selected row is out of sync with queue.")
                return

            entry = self._ledger_list[idx]
            fields["led_name"].delete(0, "end")
            fields["led_name"].insert(0, entry.get("Name", ""))
            fields["led_parent"].delete(0, "end")
            fields["led_parent"].insert(0, entry.get("Parent", ""))
            fields["led_gst_app"].delete(0, "end")
            fields["led_gst_app"].insert(0, entry.get("GSTApplicable", ""))
            fields["led_gstin"].delete(0, "end")
            fields["led_gstin"].insert(0, entry.get("GSTIN", ""))
            fields["led_state"].delete(0, "end")
            fields["led_state"].insert(0, entry.get("StateOfSupply", ""))
            fields["led_tax_type"].delete(0, "end")
            fields["led_tax_type"].insert(0, entry.get("TypeOfTaxation", ""))
            fields["led_gst_rate"].delete(0, "end")
            fields["led_gst_rate"].insert(0, entry.get("GSTRate", ""))

            ledger_edit_index = idx
            add_ledger_btn.configure(text="💾  Update Selected")

        add_ledger_btn = ctk.CTkButton(
            form,
            text="➕  Add to Queue",
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=add_ledger,
        )
        add_ledger_btn.pack(fill="x", padx=12, pady=(10,6))
        ctk.CTkButton(
            form,
            text="✏️  Edit Selected",
            fg_color="#94A3B8",
            hover_color="#64748B",
            text_color="#FFFFFF",
            command=edit_selected_ledger,
        ).pack(fill="x", padx=12, pady=(0,12))

        # Right: Queue table + actions
        right = ctk.CTkFrame(main, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(0, weight=1)

        led_tree = ttk.Treeview(right, columns=("Name","Parent","GST","GSTIN","Rate"),
                                 show="headings", height=8)
        for c in ("Name","Parent","GST","GSTIN","Rate"):
            led_tree.heading(c, text=c)
            led_tree.column(c, width=120)
        led_tree.grid(row=0, column=0, sticky="nsew", pady=(0,8))

        self._led_count_label = ctk.CTkLabel(right, text="0 ledger(s) queued",
                                              font=("Segoe UI", 11), text_color=TEXT_MUTED)
        self._led_count_label.grid(row=1, column=0, sticky="w")

        btn_row = ctk.CTkFrame(right, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", pady=(5,0))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)

        def clear_ledgers():
            nonlocal ledger_edit_index
            self._ledger_list.clear()
            led_tree.delete(*led_tree.get_children())
            self._led_count_label.configure(text="0 ledger(s) queued")
            ledger_edit_index = None
            add_ledger_btn.configure(text="➕  Add to Queue")

        def export_ledgers(action):
            company = self._get_selected_company()
            if not self._ledger_list:
                messagebox.showwarning("Empty","Add at least one ledger.")
                return
            xml = generate_ledger_xml(self._ledger_list, company)
            if action == "save":
                out = filedialog.asksaveasfilename(defaultextension=".xml",
                        initialfile="Tally_Ledgers.xml", filetypes=[("XML","*.xml")])
                if out:
                    with open(out,"w",encoding="utf-8") as f: f.write(xml)
                    messagebox.showinfo("Saved", f"Ledger XML saved!\n{out}")
            else:
                try:
                    tally_url = self._get_tally_url()
                    host, port_text = tally_url.rsplit(":", 1)
                    host = host.replace("http://", "", 1)
                    target_company = company or "Loaded company in Tally"
                    self.status_var.set(f"Posting ledgers to Tally ({target_company})...")
                    resp = push_to_tally(xml, host, int(port_text))
                    self.status_var.set(f"Ledgers posted to Tally ({target_company})")
                    messagebox.showinfo("Tally Response", f"Target Company: {target_company}\n\n{resp[:1000]}")
                except ValueError as e:
                    messagebox.showerror("Invalid Settings", str(e))
                except Exception as e:
                    messagebox.showerror("Error", str(e))

        ctk.CTkButton(btn_row, text="💾 Save XML", fg_color=SUCCESS, hover_color="#15803D",
                       command=lambda: export_ledgers("save")).grid(row=0, column=0, sticky="ew", padx=(0,6), pady=(0,6))
        ctk.CTkButton(btn_row, text="🚀 Push to Tally", fg_color=ACCENT, hover_color=ACCENT_HOVER,
                       command=lambda: export_ledgers("push")).grid(row=0, column=1, sticky="ew", padx=(6,0), pady=(0,6))
        ctk.CTkButton(btn_row, text="🗑 Clear", fg_color=DANGER, hover_color="#B91C1C",
                       command=clear_ledgers).grid(row=1, column=0, sticky="ew", padx=(0,6))

        # Load from Excel
        def load_ledgers_excel():
            f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")])
            if not f: return
            try:
                _, rows = read_excel(f)
                for r in rows:
                    entry = {
                        "Name": str(r.get("Name","") or r.get("LedgerName","") or ""),
                        "Parent": str(r.get("Parent","") or r.get("ParentGroup","") or "Sundry Debtors"),
                        "GSTApplicable": str(r.get("GSTApplicable","") or ""),
                        "GSTIN": str(r.get("GSTIN","") or ""),
                        "StateOfSupply": str(r.get("State","") or r.get("StateOfSupply","") or ""),
                        "TypeOfTaxation": str(r.get("TaxType","") or r.get("TypeOfTaxation","") or ""),
                        "GSTRate": str(r.get("GSTRate","") or ""),
                    }
                    if entry["Name"]:
                        self._ledger_list.append(entry)
                        led_tree.insert("","end", values=(entry["Name"], entry["Parent"],
                                        entry["GSTApplicable"], entry["GSTIN"], entry["GSTRate"]))
                self._led_count_label.configure(text=f"{len(self._ledger_list)} ledger(s) queued")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(
            btn_row,
            text="📂 Load Excel",
            fg_color="#94A3B8",
            hover_color="#64748B",
            text_color="#FFFFFF",
            command=load_ledgers_excel,
        ).grid(row=1, column=1, sticky="ew", padx=(6,0))

    # ── STOCK ITEM CREATION TAB ─────────────────────────────────────────

    def _build_stock_tab(self):
        parent = self.tab_stock

        info = ctk.CTkLabel(parent, text="Create Stock Items (Inventory Masters) in TallyPrime",
                             font=("Segoe UI", 13, "bold"), text_color=TEXT_PRIMARY)
        info.pack(pady=(10,5))

        main = ctk.CTkFrame(parent, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=10, pady=5)
        main.grid_columnconfigure(0, weight=1, minsize=300)
        main.grid_columnconfigure(1, weight=2)
        main.grid_rowconfigure(0, weight=1)

        form = ctk.CTkScrollableFrame(
            main,
            fg_color=COLORS["bg_input"],
            corner_radius=10,
            width=360,
            border_width=1,
            border_color=COLORS["border"],
        )
        form.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        form.grid_columnconfigure(0, weight=1)

        fields = {}
        configs = [
            ("Item Name *", "item_name", "e.g. Laptop Dell Inspiron"),
            ("Stock Group", "item_parent", "Primary"),
            ("Unit", "item_unit", "Nos / Pcs / Kg / Ltr"),
            ("HSN/SAC Code", "item_hsn", "e.g. 84713010"),
            ("GST Rate %", "item_gst_rate", "e.g. 18"),
            ("Description", "item_desc", "Optional description"),
        ]
        for label, key, placeholder in configs:
            ctk.CTkLabel(form, text=label, font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
            e = ctk.CTkEntry(
                form,
                placeholder_text=placeholder,
                fg_color=COLORS["bg_card"],
                border_color=COLORS["border"],
                text_color=COLORS["text_primary"],
            )
            e.pack(fill="x", padx=12, pady=(0,2))
            fields[key] = e

        self._stock_list = []
        stock_edit_index = None

        def add_item():
            nonlocal stock_edit_index
            name = fields["item_name"].get().strip()
            if not name:
                messagebox.showwarning("Required","Item Name is required.")
                return
            parent_group = _normalize_stock_group_name(fields["item_parent"].get())
            unit_name = _normalize_stock_unit_name(fields["item_unit"].get())
            entry = {
                "Name": name,
                "Parent": parent_group,
                "Unit": unit_name,
                "HSNCode": fields["item_hsn"].get().strip(),
                "GSTRate": fields["item_gst_rate"].get().strip(),
                "GSTApplicable": "Applicable" if fields["item_gst_rate"].get().strip() else "",
                "Description": fields["item_desc"].get().strip(),
            }
            row_values = (name, entry["Parent"], entry["Unit"], entry["HSNCode"], entry["GSTRate"])

            if stock_edit_index is None:
                self._stock_list.append(entry)
                stk_tree.insert("", "end", values=row_values)
            else:
                self._stock_list[stock_edit_index] = entry
                item_ids = stk_tree.get_children()
                if stock_edit_index < len(item_ids):
                    stk_tree.item(item_ids[stock_edit_index], values=row_values)
                stock_edit_index = None
                add_item_btn.configure(text="➕  Add to Queue")

            for v in fields.values():
                v.delete(0, "end")
            self._stk_count_label.configure(text=f"{len(self._stock_list)} item(s) queued")

        def edit_selected_item():
            nonlocal stock_edit_index
            selected = stk_tree.selection()
            if not selected:
                messagebox.showwarning("Select Row", "Select a queued stock item row to edit.")
                return

            item_id = selected[0]
            idx = stk_tree.index(item_id)
            if idx >= len(self._stock_list):
                messagebox.showwarning("Selection Error", "Selected row is out of sync with queue.")
                return

            entry = self._stock_list[idx]
            fields["item_name"].delete(0, "end")
            fields["item_name"].insert(0, entry.get("Name", ""))
            fields["item_parent"].delete(0, "end")
            fields["item_parent"].insert(0, entry.get("Parent", ""))
            fields["item_unit"].delete(0, "end")
            fields["item_unit"].insert(0, entry.get("Unit", ""))
            fields["item_hsn"].delete(0, "end")
            fields["item_hsn"].insert(0, entry.get("HSNCode", ""))
            fields["item_gst_rate"].delete(0, "end")
            fields["item_gst_rate"].insert(0, entry.get("GSTRate", ""))
            fields["item_desc"].delete(0, "end")
            fields["item_desc"].insert(0, entry.get("Description", ""))

            stock_edit_index = idx
            add_item_btn.configure(text="💾  Update Selected")

        add_item_btn = ctk.CTkButton(
            form,
            text="➕  Add to Queue",
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=add_item,
        )
        add_item_btn.pack(fill="x", padx=12, pady=(10,6))
        ctk.CTkButton(
            form,
            text="✏️  Edit Selected",
            fg_color="#94A3B8",
            hover_color="#64748B",
            text_color="#FFFFFF",
            command=edit_selected_item,
        ).pack(fill="x", padx=12, pady=(0,12))

        right = ctk.CTkFrame(main, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(0, weight=1)

        stk_tree = ttk.Treeview(right, columns=("Name","Group","Unit","HSN","GST%"),
                                 show="headings", height=8)
        for c in ("Name","Group","Unit","HSN","GST%"):
            stk_tree.heading(c, text=c)
            stk_tree.column(c, width=120)
        stk_tree.grid(row=0, column=0, sticky="nsew", pady=(0,8))

        self._stk_count_label = ctk.CTkLabel(right, text="0 item(s) queued",
                                              font=("Segoe UI", 11), text_color=TEXT_MUTED)
        self._stk_count_label.grid(row=1, column=0, sticky="w")

        btn_row = ctk.CTkFrame(right, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", pady=(5,0))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)

        def clear_stock():
            nonlocal stock_edit_index
            self._stock_list.clear()
            stk_tree.delete(*stk_tree.get_children())
            self._stk_count_label.configure(text="0 item(s) queued")
            stock_edit_index = None
            add_item_btn.configure(text="➕  Add to Queue")

        def export_stock(action):
            company = self._get_selected_company()
            if not self._stock_list:
                messagebox.showwarning("Empty","Add at least one stock item.")
                return
            normalized_items = []
            for item in self._stock_list:
                normalized = dict(item)
                normalized["Parent"] = _normalize_stock_group_name(normalized.get("Parent", ""))
                normalized["Unit"] = _normalize_stock_unit_name(normalized.get("Unit", "Nos"))
                normalized_items.append(normalized)

            xml = generate_stockitem_xml(normalized_items, company)
            if action == "save":
                out = filedialog.asksaveasfilename(defaultextension=".xml",
                        initialfile="Tally_StockItems.xml", filetypes=[("XML","*.xml")])
                if out:
                    with open(out,"w",encoding="utf-8") as f: f.write(xml)
                    messagebox.showinfo("Saved", f"Stock Item XML saved!\n{out}")
            else:
                try:
                    tally_url = self._get_tally_url()
                    host, port_text = tally_url.rsplit(":", 1)
                    host = host.replace("http://", "", 1)
                    target_company = company or "Loaded company in Tally"

                    # Ensure required units exist before creating stock items.
                    required_units = _collect_required_stock_units(normalized_items)
                    if required_units:
                        unit_xml = generate_unit_xml(required_units, company)
                        unit_resp = push_to_tally(unit_xml, host, int(port_text))
                        unit_parsed = _parse_tally_response_details(unit_resp)
                        self._append_debug_log(
                            "stock-unit",
                            target_company,
                            unit_xml,
                            unit_resp,
                            unit_parsed,
                            note=f"units={','.join(required_units)}",
                        )
                        unit_blockers = []
                        for err in unit_parsed.get("line_errors", []):
                            err_text = err.lower()
                            if "already exists" in err_text or "already present" in err_text:
                                continue
                            unit_blockers.append(err)
                        if unit_blockers:
                            messagebox.showerror(
                                "Unit Error",
                                "Could not create required unit masters in Tally.\n\n"
                                f"Target Company: {target_company}\n"
                                f"Detail: {unit_blockers[0]}",
                            )
                            return

                    # Ensure parent stock groups exist before creating stock items.
                    required_groups = _collect_required_stock_groups(normalized_items)
                    if required_groups:
                        group_blockers = ["Unknown stock group create failure"]
                        for parent_mode, force_parent in (("", False), ("Primary", True)):
                            group_xml = generate_stockgroup_xml(
                                required_groups,
                                company,
                                default_parent=parent_mode,
                                force_primary_parent=force_parent,
                            )
                            group_resp = push_to_tally(group_xml, host, int(port_text))
                            group_parsed = _parse_tally_response_details(group_resp)
                            self._append_debug_log(
                                "stock-group",
                                target_company,
                                group_xml,
                                group_resp,
                                group_parsed,
                                note=f"parent_mode={parent_mode or 'root'}, groups={','.join(required_groups)}",
                            )
                            blockers = []
                            for err in group_parsed.get("line_errors", []):
                                err_text = err.lower()
                                if "already exists" in err_text or "already present" in err_text:
                                    continue
                                blockers.append(err)
                            if not blockers:
                                group_blockers = []
                                break
                            group_blockers = blockers

                        if group_blockers:
                            messagebox.showerror(
                                "Stock Group Error",
                                "Could not create required stock groups in Tally.\n\n"
                                f"Target Company: {target_company}\n"
                                f"Detail: {group_blockers[0]}",
                            )
                            return

                    self.status_var.set(f"Posting stock items to Tally ({target_company})...")
                    resp = push_to_tally(xml, host, int(port_text))
                    parsed = _parse_tally_response_details(resp)
                    self._append_debug_log("stock-master", target_company, xml, resp, parsed, note="stock_item_push")

                    if parsed.get("success"):
                        self.status_var.set(f"Stock items posted to Tally ({target_company})")
                        messagebox.showinfo(
                            "Tally Response",
                            "Stock items posted successfully.\n\n"
                            f"Target Company: {target_company}\n"
                            f"Created: {parsed.get('created', 0)}\n"
                            f"Altered: {parsed.get('altered', 0)}\n"
                            f"Ignored: {parsed.get('ignored', 0)}",
                        )
                    else:
                        detail = parsed.get("error") or "Unknown Tally error"
                        if parsed.get("line_errors"):
                            detail = parsed["line_errors"][0]
                        self.status_var.set("Stock item push failed")
                        messagebox.showerror(
                            "Push Failed",
                            "Tally returned an exception/error while importing stock items.\n\n"
                            f"Target Company: {target_company}\n"
                            f"Errors: {parsed.get('errors', 0)}\n"
                            f"Exceptions: {parsed.get('exceptions', 0)}\n"
                            f"Detail: {detail}",
                        )
                except ValueError as e:
                    messagebox.showerror("Invalid Settings", str(e))
                except Exception as e:
                    messagebox.showerror("Error", str(e))

        ctk.CTkButton(btn_row, text="💾 Save XML", fg_color=SUCCESS, hover_color="#15803D",
                       command=lambda: export_stock("save")).grid(row=0, column=0, sticky="ew", padx=(0,6), pady=(0,6))
        ctk.CTkButton(btn_row, text="🚀 Push to Tally", fg_color=ACCENT, hover_color=ACCENT_HOVER,
                       command=lambda: export_stock("push")).grid(row=0, column=1, sticky="ew", padx=(6,0), pady=(0,6))
        ctk.CTkButton(btn_row, text="🗑 Clear", fg_color=DANGER, hover_color="#B91C1C",
                       command=clear_stock).grid(row=1, column=0, sticky="ew", padx=(0,6))

        def load_stock_excel():
            f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")])
            if not f: return
            try:
                _, rows = read_excel(f)
                for r in rows:
                    parent_group = _normalize_stock_group_name(
                        r.get("Parent", "") or r.get("StockGroup", "") or "Primary"
                    )
                    unit_name = _normalize_stock_unit_name(r.get("Unit", "") or "Nos")
                    entry = {
                        "Name": str(r.get("Name","") or r.get("ItemName","") or ""),
                        "Parent": parent_group,
                        "Unit": unit_name,
                        "HSNCode": str(r.get("HSNCode","") or r.get("HSN","") or ""),
                        "GSTRate": str(r.get("GSTRate","") or ""),
                        "GSTApplicable": "Applicable",
                        "Description": str(r.get("Description","") or ""),
                    }
                    if entry["Name"]:
                        self._stock_list.append(entry)
                        stk_tree.insert("","end", values=(entry["Name"], entry["Parent"],
                                        entry["Unit"], entry["HSNCode"], entry["GSTRate"]))
                self._stk_count_label.configure(text=f"{len(self._stock_list)} item(s) queued")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(
            btn_row,
            text="📂 Load Excel",
            fg_color="#94A3B8",
            hover_color="#64748B",
            text_color="#FFFFFF",
            command=load_stock_excel,
        ).grid(row=1, column=1, sticky="ew", padx=(6,0))


# ═══════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = TallySalesApp()
    app.mainloop()