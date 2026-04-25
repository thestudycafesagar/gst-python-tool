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
import webbrowser

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

LEDGER_PARENT_OPTIONS = [
    "Sundry Debtors",
    "Sundry Creditors",
    "Sales Accounts",
    "Purchase Accounts",
    "Duties & Taxes",
    "Direct Incomes",
    "Indirect Incomes",
    "Direct Expenses",
    "Indirect Expenses",
    "Fixed Assets",
]

LEDGER_GST_APPLICABLE_OPTIONS = [
    "Applicable",
    "Not Applicable",
]

LEDGER_STATE_OPTIONS = [
    "Not Applicable",
    "Andaman & Nicobar Islands",
    "Andhra Pradesh",
    "Arunachal Pradesh",
    "Assam",
    "Bihar",
    "Chandigarh",
    "Chhattisgarh",
    "Dadra & Nagar Haveli and Daman & Diu",
    "Delhi",
    "Goa",
    "Gujarat",
    "Haryana",
    "Himachal Pradesh",
    "Jammu & Kashmir",
    "Jharkhand",
    "Karnataka",
    "Kerala",
    "Ladakh",
    "Lakshadweep",
    "Madhya Pradesh",
    "Maharashtra",
    "Manipur",
    "Meghalaya",
    "Mizoram",
    "Nagaland",
    "Odisha",
    "Puducherry",
    "Punjab",
    "Rajasthan",
    "Sikkim",
    "Tamil Nadu",
    "Telangana",
    "Tripura",
    "Uttarakhand",
    "Uttar Pradesh",
    "West Bengal",
]

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


def _normalize_manual_date_to_tally(date_text: str) -> str:
    text = str(date_text or "").strip()
    if not text:
        raise ValueError("Custom date is empty.")

    compact = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d{8}", compact):
        for fmt in ("%Y%m%d", "%d%m%Y"):
            try:
                return datetime.strptime(compact, fmt).strftime("%Y%m%d")
            except ValueError:
                continue

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y%m%d")
        except ValueError:
            continue

    raise ValueError("Invalid custom date format. Use DD/MM/YYYY, DD-MM-YYYY, or YYYY-MM-DD.")

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


def _row_text_any(row: dict, keys: list, default: str = "") -> str:
    for key in keys or []:
        value = _row_text(row, key, "")
        if value:
            return value
    return default


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


def _is_effectively_blank_ledger(value: str) -> bool:
    text = str(value or "").strip()
    if not text:
        return True

    key = text.casefold()
    if key in {
        "na",
        "n/a",
        "none",
        "null",
        "nil",
        "not applicable",
        "* not applicable",
        "-",
        "--",
    }:
        return True

    compact = text.replace(",", "")
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", compact):
        try:
            return float(compact) == 0.0
        except ValueError:
            return False

    return False


def _ledger_or_suspense(value: str, fallback: str = SUSPENSE_LEDGER) -> str:
    text = str(value or "").strip()
    if _is_effectively_blank_ledger(text):
        text = ""
    return text or fallback


def _is_party_parent(parent: str) -> bool:
    key = re.sub(r"\s+", " ", str(parent or "")).strip().casefold()
    return key in {"sundry debtors", "sundry creditors"}


def _is_duties_parent(parent: str) -> bool:
    key = re.sub(r"\s+", " ", str(parent or "")).strip().casefold()
    return key in {"duties & taxes", "duties and taxes", "duty"}


def _current_fy_start() -> str:
    today = date.today()
    fy_start_year = today.year if today.month >= 4 else today.year - 1
    return f"{fy_start_year}0401"


def _state_name_from_gstin(gstin: str) -> str:
    gstin_text = str(gstin or "").strip().upper()
    state_map = {
        "01": "Jammu And Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh",
        "05": "Uttarakhand", "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh",
        "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
        "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
        "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh",
        "24": "Gujarat", "25": "Daman And Diu", "26": "Dadra And Nagar Haveli And Daman And Diu",
        "27": "Maharashtra", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala",
        "33": "Tamil Nadu", "34": "Puducherry", "35": "Andaman And Nicobar Islands", "36": "Telangana",
        "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory", "99": "Centre Jurisdiction",
    }
    return state_map.get(gstin_text[:2], "")


def _normalize_gst_applicable(value: str, gstin: str = "") -> str:
    raw = str(value or "").strip()
    key = raw.casefold()
    if key in {"applicable", "yes", "y", "true", "1", "registered", "regular", "gst applicable"}:
        return "Applicable"
    if key in {"not applicable", "no", "n", "false", "0", "na", "n/a", "notapplicable"}:
        return "Not Applicable"
    if gstin:
        return "Applicable"
    return raw


def _normalize_gst_registration_type(value: str, gstin: str = "", gst_applicable: str = "") -> str:
    raw = str(value or "").strip()
    if not raw:
        if gstin or str(gst_applicable).strip().casefold() == "applicable":
            return "Regular"
        return ""

    key = raw.casefold()
    mapping = {
        "regular": "Regular",
        "registered": "Regular",
        "composition": "Composition",
        "consumer": "Consumer",
        "unregistered": "Unregistered",
        "sez": "SEZ",
        "sez unit": "SEZ",
        "sez developer": "SEZ",
        "overseas": "Overseas",
    }
    return mapping.get(key, raw)


def _normalize_state_for_ledger(value: str) -> str:
    text = str(value or "").strip()
    if text.casefold() in {"not applicable", "* not applicable", "na", "n/a"}:
        return ""
    return text


def _company_static_block(company: str) -> str:
    selected = str(company or "").strip()
    if not selected:
        return ""
    return f"   <STATICVARIABLES><SVCURRENTCOMPANY>{xml_escape(selected)}</SVCURRENTCOMPANY></STATICVARIABLES>"


def _collect_party_context(
    row: dict,
    party_ledger: str,
    allow_place_of_supply_column: bool = True,
) -> dict:
    party_ledger_raw = _ledger_or_suspense(party_ledger)
    party_name_raw = _row_text_any(
        row,
        [
            "PartyName",
            "BuyerName",
            "SupplierName",
            "BillToName",
            "Party",
        ],
        default=party_ledger_raw,
    )
    mailing_name_raw = _row_text_any(
        row,
        [
            "PartyMailingName",
            "MailingName",
            "BillingName",
            "Supplier",
            "Bill To Name",
        ],
        default=party_name_raw,
    )
    gstin_raw = _row_text_any(
        row,
        [
            "PartyGSTIN",
            "GSTIN",
            "GSTIN/UIN",
            "GSTIN UIN",
            "Party GSTIN",
            "SupplierGSTIN",
            "Supplier GSTIN",
            "GST No",
            "GST Number",
        ],
    ).upper()

    # Party's own state — look only at direct state columns, NOT PlaceOfSupply
    # (PlaceOfSupply is a delivery/billing concept for the sales side).
    state_raw = _row_text_any(
        row,
        [
            "PartyState",
            "State",
            "StateName",
            "State Name",
        ],
    )
    state_raw = _normalize_state_for_ledger(state_raw)
    if not state_raw and gstin_raw:
        state_raw = _state_name_from_gstin(gstin_raw)

    # PlaceOfSupply is valid for sales delivery context. For purchases,
    # do not let Excel POS override party state/GSTIN-derived state.
    if allow_place_of_supply_column:
        place_raw = _row_text_any(
            row,
            [
                "PlaceOfSupply",
                "Place Of Supply",
                "Place of Supply",
                "POS",
                "StateOfSupply",
            ],
            default=state_raw,
        )
    else:
        place_raw = state_raw
    place_raw = _normalize_state_for_ledger(place_raw)
    if not place_raw and gstin_raw:
        place_raw = _state_name_from_gstin(gstin_raw)

    country_raw = _row_text_any(
        row,
        [
            "PartyCountry",
            "Country",
            "Country Name",
            "CountryOfResidence",
        ],
        default="India",
    )
    pincode_raw = _row_text_any(
        row,
        [
            "PartyPincode",
            "Pincode",
            "PinCode",
            "PIN",
            "PIN Code",
            "PostalCode",
            "Postal Code",
        ],
    )
    address1_raw = _row_text_any(
        row,
        [
            "PartyAddress1",
            "PartyAddressLine1",
            "Address1",
            "Address Line 1",
            "Address Line1",
            "AddressLine1",
            "BillToAddress",
            "Bill To Address",
            "Address",
        ],
    )
    address2_raw = _row_text_any(
        row,
        [
            "PartyAddress2",
            "PartyAddressLine2",
            "Address2",
            "Address Line 2",
            "Address Line2",
            "AddressLine2",
        ],
    )

    gst_app_raw = _row_text_any(
        row,
        [
            "GSTApplicable",
            "GST Applicable",
            "IsGSTApplicable",
            "GST",
        ],
    )
    reg_type_raw = _row_text_any(
        row,
        [
            "GSTRegistrationType",
            "GST Registration Type",
            "GST Reg Type",
            "RegistrationType",
            "Registration Type",
            "RegType",
            "Reg Type",
        ],
    )
    reg_type = _normalize_gst_registration_type(reg_type_raw, gstin=gstin_raw, gst_applicable=gst_app_raw)
    if reg_type.casefold() == "regular" and not gstin_raw:
        reg_type = ""

    return {
        "party_ledger": party_ledger_raw,
        "party_name": party_name_raw,
        "mailing_name": mailing_name_raw,
        "gstin": gstin_raw,
        "state": state_raw,
        "place_of_supply": place_raw,
        "country": country_raw or "India",
        "pincode": pincode_raw,
        "address1": address1_raw,
        "address2": address2_raw,
        "registration_type": reg_type,
    }


def _append_invoice_party_context_xml(
    add_line,
    party_context: dict,
    include_basic_buyer: bool = False,
    include_state: bool = True,
    include_place_of_supply: bool = True,
    place_of_supply_override=None,
) -> None:
    party_name = xml_escape(party_context.get("party_name", ""))
    mailing_name = xml_escape(party_context.get("mailing_name", "") or party_context.get("party_name", ""))
    party_gstin = xml_escape(party_context.get("gstin", ""))
    party_state = xml_escape(party_context.get("state", "")) if include_state else ""
    if place_of_supply_override is None:
        place_source = party_context.get("place_of_supply", "")
    else:
        place_source = place_of_supply_override
    place_of_supply = xml_escape(place_source) if include_place_of_supply else ""
    country = xml_escape(party_context.get("country", "") or "India")
    pincode = xml_escape(party_context.get("pincode", ""))
    address1 = xml_escape(party_context.get("address1", ""))
    address2 = xml_escape(party_context.get("address2", ""))
    reg_type_raw = str(party_context.get("registration_type", "") or "").strip()
    party_gstin_raw = str(party_context.get("gstin", "") or "").strip()
    country_raw = str(party_context.get("country", "") or "India").strip()

    # Tally marks vouchers as uncertain when GST registration is omitted.
    # Always send an explicit registration type for invoice vouchers.
    if not reg_type_raw:
        if party_gstin_raw:
            reg_type_raw = "Regular"
        elif country_raw and country_raw.casefold() not in {"india"}:
            reg_type_raw = "Overseas"
        else:
            reg_type_raw = "Unregistered"

    reg_type = xml_escape(reg_type_raw)

    if address1 or address2:
        add_line('     <ADDRESS.LIST TYPE="String">')
        if address1:
            add_line(f'      <ADDRESS>{address1}</ADDRESS>')
        if address2:
            add_line(f'      <ADDRESS>{address2}</ADDRESS>')
        add_line('     </ADDRESS.LIST>')

    if reg_type:
        add_line(f'     <GSTREGISTRATIONTYPE>{reg_type}</GSTREGISTRATIONTYPE>')
        if reg_type.casefold() == "regular":
            add_line('     <VATDEALERTYPE>Regular</VATDEALERTYPE>')

    if party_state:
        add_line(f'     <STATENAME>{party_state}</STATENAME>')
        add_line(f'     <PARTYSTATENAME>{party_state}</PARTYSTATENAME>')

    add_line(f'     <COUNTRYOFRESIDENCE>{country}</COUNTRYOFRESIDENCE>')

    if party_gstin:
        add_line(f'     <PARTYGSTIN>{party_gstin}</PARTYGSTIN>')

    if place_of_supply:
        add_line(f'     <PLACEOFSUPPLY>{place_of_supply}</PLACEOFSUPPLY>')

    if party_name:
        add_line(f'     <PARTYNAME>{party_name}</PARTYNAME>')
        add_line(f'     <BASICBASEPARTYNAME>{party_name}</BASICBASEPARTYNAME>')
        if include_basic_buyer:
            add_line(f'     <BASICBUYERNAME>{party_name}</BASICBUYERNAME>')

    if mailing_name:
        add_line(f'     <PARTYMAILINGNAME>{mailing_name}</PARTYMAILINGNAME>')

    if pincode:
        add_line(f'     <PARTYPINCODE>{pincode}</PARTYPINCODE>')


def _state_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def _is_gstin_like(value: str) -> bool:
    text = str(value or "").strip().upper()
    return bool(re.fullmatch(r"\d{2}[A-Z0-9]{13}", text))


def _pick_company_gst_registration(registrations: list, preferred_state: str = "") -> dict:
    rows = list(registrations or [])
    if not rows:
        return {}

    preferred_key = _state_key(preferred_state)

    def score(entry: dict) -> tuple:
        state_key = _state_key(entry.get("state", ""))
        name = str(entry.get("name", "") or "").strip()
        gstin = str(entry.get("gstin", "") or "").strip().upper()

        return (
            1 if (preferred_key and preferred_key == state_key) else 0,
            1 if (name and not _is_gstin_like(name)) else 0,
            1 if bool(gstin) else 0,
        )

    return max(rows, key=score)


def _resolve_company_gst_state(registrations: list, preferred_state: str = "") -> str:
    selected = _pick_company_gst_registration(registrations, preferred_state=preferred_state)
    if not selected:
        return ""
    return str(selected.get("state", "") or "").strip()


def _append_company_gst_context_xml(
    add_line,
    party_context: dict,
    company_gst_registrations: list = None,
    prefer_party_state: bool = True,
) -> None:
    registrations = list(company_gst_registrations or [])
    if not registrations:
        return

    preferred_state = ""
    if prefer_party_state:
        preferred_state = (
            str(party_context.get("place_of_supply", "") or "").strip()
            or str(party_context.get("state", "") or "").strip()
        )
    selected = _pick_company_gst_registration(registrations, preferred_state=preferred_state)
    if not selected:
        return

    reg_name_raw = str(selected.get("name", "") or "").strip()
    reg_gstin_raw = str(selected.get("gstin", "") or "").strip().upper()
    reg_state_raw = str(selected.get("state", "") or "").strip()

    if not reg_name_raw and reg_state_raw:
        reg_name_raw = f"{reg_state_raw} Registration"
    if not reg_name_raw and reg_gstin_raw:
        reg_name_raw = reg_gstin_raw

    reg_name = xml_escape(reg_name_raw)
    reg_gstin = xml_escape(reg_gstin_raw)
    reg_state = xml_escape(reg_state_raw)

    if reg_name and reg_gstin:
        add_line(
            f'     <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="{reg_gstin}">{reg_name}</GSTREGISTRATION>'
        )
        add_line(f'     <CMPGSTIN>{reg_gstin}</CMPGSTIN>')
        add_line('     <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>')
    if reg_state:
        add_line(f'     <CMPGSTSTATE>{reg_state}</CMPGSTSTATE>')


def _append_tax_object_allocation_xml(add_line, tax_classification: str) -> None:
    tax_class = xml_escape(str(tax_classification or "").strip())
    if not tax_class:
        return
    # Many Tally companies do not define Tax Classification masters named
    # exactly IGST/CGST/SGST. Emitting those names causes import exception.
    if tax_class.casefold() in {"igst", "cgst", "sgst", "utgst", "cess"}:
        return
    add_line('      <TAXOBJECTALLOCATIONS.LIST>')
    add_line('       <TAXOBJECTALLOCATIONS>')
    add_line('        <TAXTYPE>GST</TAXTYPE>')
    add_line('        <TAXABILITY>Taxable</TAXABILITY>')
    add_line(f'        <TAXCLASSIFICATIONNAME>{tax_class}</TAXCLASSIFICATIONNAME>')
    add_line('       </TAXOBJECTALLOCATIONS>')
    add_line('      </TAXOBJECTALLOCATIONS.LIST>')


def _gst_transaction_type(reg_type: str, gstin: str = "") -> str:
    """Return the GSTTRANSACTIONTYPE value Tally needs to avoid 'Uncertain' in GSTR.
    Rules:
      - Regular registered party with GSTIN  -> 'Tax Invoice' (B2B)
      - Composition/Consumer/Unregistered    -> 'Unregistered'
      - Overseas / non-India                 -> 'Overseas'
      - SEZ                                  -> 'SEZ exports with payment'
      - Empty / unknown                      -> 'Unregistered'
    """
    key = str(reg_type or "").strip().casefold()
    if key in {"regular", "registered"} and gstin:
        return "Tax Invoice"
    if key in {"composition", "consumer", "unregistered", ""}:
        return "Unregistered"
    if key in {"overseas"}:
        return "Overseas"
    if key in {"sez", "sez unit", "sez developer"}:
        return "SEZ exports with payment"
    # fallback
    if gstin:
        return "Tax Invoice"
    return "Unregistered"


def _pick_tax_ledger_name(row: dict, ledger_keys: list, rate_value: float, default_name: str) -> str:
    tax_ledger_raw = _row_text_any(row, ledger_keys, "")
    if _is_effectively_blank_ledger(tax_ledger_raw):
        tax_ledger_raw = ""
    if rate_value > 0 and not tax_ledger_raw:
        tax_ledger_raw = default_name
    return _ledger_or_suspense(tax_ledger_raw)


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


def _extract_ledger_names(response_text: str) -> set:
    names = set()

    try:
        root = ET.fromstring(response_text)
        for node in root.iter():
            tag = str(node.tag or "").upper()
            if tag == "LEDGER":
                attr_name = _normalize_company_name(node.attrib.get("NAME") or "")
                if attr_name:
                    names.add(attr_name)
                for child in list(node):
                    child_tag = str(child.tag or "").upper()
                    child_text = _normalize_company_name(child.text)
                    if child_tag == "NAME" and child_text:
                        names.add(child_text)
    except ET.ParseError:
        pass

    for match in re.findall(r"<LEDGER\b[^>]*\bNAME=\"([^\"]+)\"", response_text, flags=re.IGNORECASE):
        value = _normalize_company_name(match)
        if value:
            names.add(value)

    return names


def _fetch_existing_ledger_names(tally_url: str, company_name: str = "", timeout: float = 15.0) -> dict:
    selected_company = _normalize_company_name(company_name)
    static_vars = "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    if selected_company:
        static_vars += f"<SVCURRENTCOMPANY>{xml_escape(selected_company)}</SVCURRENTCOMPANY>"
    static_vars += "</STATICVARIABLES>"

    request_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Ledger Name Lookup</ID></HEADER>"
        f"<BODY><DESC>{static_vars}<TDL><TDLMESSAGE>"
        "<COLLECTION NAME='Ledger Name Lookup'><TYPE>Ledger</TYPE>"
        "<FETCH>Name,Parent</FETCH><NATIVEMETHOD>Name</NATIVEMETHOD></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

    try:
        response_text = _post_tally_xml(tally_url, request_xml, timeout=timeout)
    except HTTPError as exc:
        return {"success": False, "error": f"HTTP {exc.code}", "ledgers": set()}
    except URLError:
        return {"success": False, "error": "ConnectionError", "ledgers": set()}
    except Exception as exc:
        return {"success": False, "error": str(exc), "ledgers": set()}

    ledgers = _extract_ledger_names(response_text)
    if ledgers:
        return {"success": True, "ledgers": ledgers}

    return {
        "success": False,
        "error": "No ledgers returned for selected company.",
        "ledgers": set(),
    }


def _extract_company_gst_registrations(response_text: str) -> list:
    registrations = []
    seen = set()

    try:
        root = ET.fromstring(response_text)
        for tax_unit in root.findall(".//TAXUNIT"):
            tax_type = str(tax_unit.attrib.get("TAXTYPE") or tax_unit.findtext("TAXTYPE") or "").strip().upper()
            if tax_type and tax_type != "GST":
                continue

            name_raw = str(tax_unit.attrib.get("NAME") or tax_unit.findtext("NAME") or "").strip()
            gstin_raw = str(
                tax_unit.attrib.get("TAXREGISTRATION")
                or tax_unit.findtext("GSTREGNUMBER")
                or tax_unit.findtext("GSTIN")
                or ""
            ).strip().upper()
            state_raw = str(tax_unit.findtext("STATENAME") or "").strip()

            if not gstin_raw and _is_gstin_like(name_raw):
                gstin_raw = name_raw.upper()
            if not state_raw and gstin_raw:
                state_raw = _state_name_from_gstin(gstin_raw)

            if not gstin_raw:
                continue

            name = _normalize_company_name(name_raw or gstin_raw)
            if not name:
                continue

            key = (name.casefold(), gstin_raw)
            if key in seen:
                continue
            seen.add(key)

            registrations.append(
                {
                    "name": name,
                    "gstin": gstin_raw,
                    "state": state_raw,
                }
            )
    except ET.ParseError:
        pass

    return registrations


def _fetch_company_gst_registrations(tally_url: str, company_name: str = "", timeout: float = 15.0) -> dict:
    selected_company = _normalize_company_name(company_name)
    static_vars = "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    if selected_company:
        static_vars += f"<SVCURRENTCOMPANY>{xml_escape(selected_company)}</SVCURRENTCOMPANY>"
    static_vars += "</STATICVARIABLES>"

    request_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Tax Unit Lookup</ID></HEADER>"
        f"<BODY><DESC>{static_vars}<TDL><TDLMESSAGE>"
        "<COLLECTION NAME='Tax Unit Lookup'><TYPE>TaxUnit</TYPE>"
        "<FETCH>Name,TaxType,TaxRegistration,GSTRegNumber,StateName,UseFor</FETCH>"
        "<NATIVEMETHOD>Name</NATIVEMETHOD></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

    try:
        response_text = _post_tally_xml(tally_url, request_xml, timeout=timeout)
    except HTTPError as exc:
        return {"success": False, "error": f"HTTP {exc.code}", "registrations": []}
    except URLError:
        return {"success": False, "error": "ConnectionError", "registrations": []}
    except Exception as exc:
        return {"success": False, "error": str(exc), "registrations": []}

    registrations = _extract_company_gst_registrations(response_text)
    if registrations:
        return {"success": True, "registrations": registrations}

    return {
        "success": False,
        "error": "No GST registrations returned for selected company.",
        "registrations": [],
    }


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


def _ledger_name_key(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def _filter_out_existing_ledgers(ledger_defs: list, existing_ledger_names) -> list:
    existing_keys = {
        _ledger_name_key(name)
        for name in (existing_ledger_names or set())
        if str(name or "").strip()
    }
    filtered = []
    seen = set()

    for entry in ledger_defs or []:
        name = str((entry or {}).get("Name", "") or "").strip()
        key = _ledger_name_key(name)
        if not key or key in seen or key in existing_keys:
            continue
        seen.add(key)
        filtered.append(entry)

    return filtered


def _extract_stock_item_balances(response_text: str) -> dict:
    balances = {}

    try:
        root = ET.fromstring(response_text)
    except ET.ParseError:
        return balances

    for stock_item in root.findall(".//STOCKITEM"):
        name = str(
            stock_item.attrib.get("NAME")
            or stock_item.findtext("NAME")
            or ""
        ).strip()
        if not name:
            continue
        base_unit = str(stock_item.findtext("BASEUNITS") or "").strip()
        closing_balance = str(stock_item.findtext("CLOSINGBALANCE") or "").strip()
        qty_match = re.search(r"-?\d+(?:\.\d+)?", closing_balance)
        balances[_ledger_name_key(name)] = {
            "name": name,
            "base_unit": base_unit,
            "closing_balance": closing_balance,
            "quantity": float(qty_match.group(0)) if qty_match else 0.0,
        }

    return balances


def _fetch_stock_item_balances(tally_url: str, company_name: str = "", timeout: float = 15.0) -> dict:
    selected_company = _normalize_company_name(company_name)
    static_vars = "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    if selected_company:
        static_vars += f"<SVCURRENTCOMPANY>{xml_escape(selected_company)}</SVCURRENTCOMPANY>"
    static_vars += "</STATICVARIABLES>"

    request_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Stock Item Balance Lookup</ID></HEADER>"
        f"<BODY><DESC>{static_vars}<TDL><TDLMESSAGE>"
        "<COLLECTION NAME='Stock Item Balance Lookup'><TYPE>StockItem</TYPE>"
        "<FETCH>Name,BaseUnits,ClosingBalance</FETCH><NATIVEMETHOD>Name</NATIVEMETHOD></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

    try:
        response_text = _post_tally_xml(tally_url, request_xml, timeout=timeout)
    except HTTPError as exc:
        return {"success": False, "error": f"HTTP {exc.code}", "items": {}}
    except URLError:
        return {"success": False, "error": "ConnectionError", "items": {}}
    except Exception as exc:
        return {"success": False, "error": str(exc), "items": {}}

    items = _extract_stock_item_balances(response_text)
    if items:
        return {"success": True, "items": items}

    return {"success": False, "error": "No stock items returned for selected company.", "items": {}}


def _find_conflicting_row_voucher_numbers(rows: list, existing_numbers) -> list:
    existing_keys = {
        _normalize_voucher_number_text(value)
        for value in (existing_numbers or set())
        if _normalize_voucher_number_text(value)
    }
    conflicts = []
    seen = set()

    for idx, row in enumerate(rows or []):
        voucher_no = _normalize_voucher_number_text(_row_voucher_number(row))
        if not voucher_no or voucher_no not in existing_keys or voucher_no in seen:
            continue
        seen.add(voucher_no)
        conflicts.append(f"row {idx + 1} voucher '{voucher_no}' already exists")

    return conflicts


def _diagnose_item_stock_issues(rows: list, stock_items: dict) -> list:
    issues = []

    for idx, row in enumerate(rows or []):
        item_name_raw = (
            _row_text(row, "ItemName")
            or _row_text(row, "Item")
            or _row_text(row, "StockItem")
            or _row_text(row, "ProductName")
            or _row_text(row, "SalesLedger")
        )
        if not item_name_raw:
            continue

        requested_qty = _row_float(row, "Quantity", 0.0)
        if requested_qty <= 0:
            requested_qty = _row_float(row, "Qty", 0.0)
        if requested_qty <= 0:
            requested_qty = _row_float(row, "Unit", 0.0)

        item = (stock_items or {}).get(_ledger_name_key(item_name_raw))
        if item is None:
            issues.append(f"row {idx + 1} stock item '{item_name_raw}' does not exist in Tally")
            continue

        if requested_qty > float(item.get("quantity", 0.0) or 0.0) + 1e-9:
            requested_unit = (
                _row_text(row, "Per", "")
                or _row_text(row, "UOM", "")
                or _row_text(row, "Unit", "")
                or item.get("base_unit", "")
            )
            requested_unit = _normalize_stock_unit_name(requested_unit) or str(item.get("base_unit", "") or "").strip()
            available_text = str(item.get("closing_balance", "") or "").strip()
            if not available_text:
                available_text = f"{fmt_amt(float(item.get('quantity', 0.0) or 0.0))} {item.get('base_unit', '')}".strip()
            issues.append(
                f"row {idx + 1} item '{item.get('name', item_name_raw)}' needs {fmt_amt(requested_qty)} {requested_unit} "
                f"but Tally stock is {available_text}"
            )

    return issues


def _infer_tax_type_from_ledger_name(ledger_name: str) -> str:
    key = _ledger_name_key(ledger_name)
    if not key:
        return ""
    if "igst" in key or "integrated tax" in key:
        return "Integrated Tax"
    if "cgst" in key or "central tax" in key:
        return "Central Tax"
    if "sgst" in key or "utgst" in key or "state tax" in key or "ut tax" in key or "state/ut tax" in key:
        return "State Tax"
    if "cess" in key:
        return "Cess"
    return ""


def _is_protected_gst_tax_ledger(ledger_name: str, parent_name: str, tax_type: str = "") -> bool:
    if not _is_duties_parent(parent_name):
        return False

    key = _ledger_name_key(ledger_name)
    tax_key = _ledger_name_key(tax_type)
    if not key and not tax_key:
        return False

    if tax_key in {"integrated tax", "central tax", "state tax", "state/ut tax", "ut tax", "cess"}:
        return True

    if key in {"igst", "cgst", "sgst", "utgst", "cess", "integrated tax", "central tax", "state tax"}:
        return True

    marker_tokens = (
        "igst",
        "cgst",
        "sgst",
        "utgst",
        "integrated tax",
        "central tax",
        "state tax",
        "ut tax",
        "state/ut tax",
        "cess",
        "gst",
    )
    return any(token in key for token in marker_tokens)


def _collect_party_ledger_definition(row: dict, is_purchase_mode: bool) -> dict:
    party_ledger_raw = _row_text(row, "PartyLedger")
    if not party_ledger_raw:
        return {}

    party_context = _collect_party_context(
        row,
        party_ledger_raw,
        allow_place_of_supply_column=not is_purchase_mode,
    )
    party_name_raw = str(party_context.get("party_ledger", "") or party_ledger_raw).strip()
    if not party_name_raw:
        return {}

    gstin_raw = str(party_context.get("gstin", "") or "").strip().upper()
    gst_app_raw = _row_text_any(
        row,
        [
            "GSTApplicable",
            "GST Applicable",
            "IsGSTApplicable",
            "GST",
        ],
    )
    gst_applicable = _normalize_gst_applicable(gst_app_raw, gstin=gstin_raw)
    if not gst_applicable:
        gst_applicable = "Applicable" if gstin_raw else "Not Applicable"
    if not gstin_raw and str(gst_applicable).strip().casefold() == "applicable":
        gst_applicable = "Not Applicable"

    reg_type_raw = _row_text_any(
        row,
        [
            "GSTRegistrationType",
            "GST Registration Type",
            "GST Reg Type",
            "RegistrationType",
            "Registration Type",
            "RegType",
            "Reg Type",
        ],
    )
    gst_reg_type = _normalize_gst_registration_type(
        reg_type_raw,
        gstin=gstin_raw,
        gst_applicable=gst_applicable,
    )
    if not gst_reg_type:
        gst_reg_type = str(party_context.get("registration_type", "") or "").strip()
    if gst_reg_type.casefold() == "regular" and not gstin_raw:
        gst_reg_type = ""

    state_raw = _normalize_state_for_ledger(
        str(party_context.get("state", "") or party_context.get("place_of_supply", "") or "")
    )
    if not state_raw and gstin_raw:
        state_raw = _state_name_from_gstin(gstin_raw)

    return {
        "Name": party_name_raw,
        "Parent": "Sundry Creditors" if is_purchase_mode else "Sundry Debtors",
        "GSTApplicable": gst_applicable,
        "GSTIN": gstin_raw,
        "StateOfSupply": state_raw,
        "Address1": str(party_context.get("address1", "") or "").strip(),
        "Address2": str(party_context.get("address2", "") or "").strip(),
        "MailingName": str(
            party_context.get("mailing_name", "")
            or party_context.get("party_name", "")
            or party_name_raw
        ).strip(),
        "Country": str(party_context.get("country", "") or "India").strip() or "India",
        "Pincode": str(party_context.get("pincode", "") or "").strip(),
        "Billwise": "Yes",
        "GSTRegistrationType": gst_reg_type,
        "TypeOfTaxation": "",
        "GSTRate": "",
    }


def _merge_ledger_definitions(existing: dict, candidate: dict) -> None:
    if not candidate:
        return

    parent_priority = {
        "": 0,
        "sales accounts": 1,
        "purchase accounts": 1,
        "sundry debtors": 2,
        "sundry creditors": 2,
        "duties & taxes": 3,
        "duties and taxes": 3,
    }

    existing_parent = str(existing.get("Parent", "") or "").strip()
    candidate_parent = str(candidate.get("Parent", "") or "").strip()
    existing_rank = parent_priority.get(existing_parent.casefold(), 1 if existing_parent else 0)
    candidate_rank = parent_priority.get(candidate_parent.casefold(), 1 if candidate_parent else 0)
    if candidate_rank > existing_rank:
        existing["Parent"] = candidate_parent
    elif not existing_parent and candidate_parent:
        existing["Parent"] = candidate_parent

    for field, value in candidate.items():
        if field in {"Name", "Parent"}:
            continue
        if existing.get(field) in (None, "") and value not in (None, ""):
            existing[field] = value


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
                key = _ledger_name_key(name)
                if not name or key in seen:
                    continue
                seen.add(key)
                missing.append(name)
    return missing


def _collect_auto_voucher_ledgers(rows: list, mode: str) -> list:
    """Collect voucher ledgers so missing ledgers can be pre-created."""
    entries = {}
    is_purchase_mode = mode in {"purchase_accounting", "purchase_item"}

    def add_entry(
        name: str,
        parent: str,
        tax_type: str = "",
        gst_rate: str = "",
        extra_fields: dict = None,
    ):
        ledger_name = str(name or "").strip()
        if _is_effectively_blank_ledger(ledger_name):
            return

        key = _ledger_name_key(ledger_name)
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

        if extra_fields:
            for field, value in extra_fields.items():
                if field in {"Name", "Parent"}:
                    continue
                if value in (None, ""):
                    continue
                candidate[field] = value

        existing = entries.get(key)
        if existing is None:
            entries[key] = candidate
            return

        _merge_ledger_definitions(existing, candidate)

    for r in rows or []:
        party_entry = _collect_party_ledger_definition(r, is_purchase_mode)
        if party_entry:
            add_entry(
                party_entry.get("Name", ""),
                party_entry.get("Parent", ""),
                extra_fields=party_entry,
            )

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

        cgst_rate_val = _row_float(r, "CGSTRate", 0.0)
        sgst_rate_val = _row_float(r, "SGSTRate", 0.0)
        igst_rate_val = _row_float(r, "IGSTRate", 0.0)

        cgst_ledger_name = _row_text_any(
            r,
            ["CGSTLedger", "CGST Ledger", "CentralTaxLedger", "Central Tax Ledger", "Central Tax"],
            "",
        )
        sgst_ledger_name = _row_text_any(
            r,
            ["SGSTLedger", "SGST Ledger", "StateTaxLedger", "State Tax Ledger", "State Tax", "UTGSTLedger", "UTGST Ledger"],
            "",
        )
        igst_ledger_name = _row_text_any(
            r,
            ["IGSTLedger", "IGST Ledger", "IntegratedTaxLedger", "Integrated Tax Ledger", "Integrated Tax"],
            "",
        )

        if _is_effectively_blank_ledger(cgst_ledger_name):
            cgst_ledger_name = ""
        if _is_effectively_blank_ledger(sgst_ledger_name):
            sgst_ledger_name = ""
        if _is_effectively_blank_ledger(igst_ledger_name):
            igst_ledger_name = ""

        if cgst_rate_val > 0 and not cgst_ledger_name:
            cgst_ledger_name = "CGST"
        if sgst_rate_val > 0 and not sgst_ledger_name:
            sgst_ledger_name = "SGST"
        if igst_rate_val > 0 and not igst_ledger_name:
            igst_ledger_name = "IGST"

        add_entry(cgst_ledger_name, "Duties & Taxes", "Central Tax", _row_text(r, "CGSTRate"))
        add_entry(sgst_ledger_name, "Duties & Taxes", "State Tax", _row_text(r, "SGSTRate"))
        add_entry(igst_ledger_name, "Duties & Taxes", "Integrated Tax", _row_text(r, "IGSTRate"))

    return list(entries.values())


def _build_missing_ledger_defs(line_errors: list, rows: list, mode: str) -> list:
    missing_names = _extract_missing_ledgers_from_line_errors(line_errors)
    if not missing_names:
        return []

    is_purchase_mode = mode in {"purchase_accounting", "purchase_item"}

    party_keys = set()
    party_defs = {}
    purchase_keys = set()
    sales_keys = set()
    tax_type_map = {}
    tax_rate_map = {}

    for r in rows or []:
        party_entry = _collect_party_ledger_definition(r, is_purchase_mode)
        if party_entry:
            party_key = _ledger_name_key(party_entry.get("Name", ""))
            if party_key:
                party_keys.add(party_key)
                existing_party = party_defs.get(party_key)
                if existing_party is None:
                    party_defs[party_key] = party_entry
                else:
                    _merge_ledger_definitions(existing_party, party_entry)

        purchase_ledger = (
            _row_text(r, "PurchaseLedger")
            or _row_text(r, "PurchaseAccount")
            or _row_text(r, "Purchase Ledger")
            or _row_text(r, "ExpenseLedger")
            or _row_text(r, "SalesLedger")
        )
        if purchase_ledger:
            purchase_keys.add(_ledger_name_key(purchase_ledger))

        sales_ledger = (
            _row_text(r, "SalesLedger")
            or _row_text(r, "SalesAccount")
            or _row_text(r, "Sales Ledger")
            or _row_text(r, "IncomeLedger")
        )
        if sales_ledger:
            sales_keys.add(_ledger_name_key(sales_ledger))

        tax_configs = [
            (
                "Central Tax",
                "CGST",
                ["CGSTLedger", "CGST Ledger", "CentralTaxLedger", "Central Tax Ledger", "Central Tax"],
                ["CGSTRate", "CGST Rate"],
            ),
            (
                "State Tax",
                "SGST",
                ["SGSTLedger", "SGST Ledger", "StateTaxLedger", "State Tax Ledger", "State Tax", "UTGSTLedger", "UTGST Ledger"],
                ["SGSTRate", "SGST Rate"],
            ),
            (
                "Integrated Tax",
                "IGST",
                ["IGSTLedger", "IGST Ledger", "IntegratedTaxLedger", "Integrated Tax Ledger", "Integrated Tax"],
                ["IGSTRate", "IGST Rate"],
            ),
        ]
        for tax_type, default_ledger, ledger_cols, rate_cols in tax_configs:
            tax_name = _row_text_any(r, ledger_cols, "")
            if _is_effectively_blank_ledger(tax_name):
                tax_name = ""
            rate_val = ""
            rate_num = 0.0
            for rate_col in rate_cols:
                candidate_rate_val = _row_text(r, rate_col, "")
                if candidate_rate_val:
                    rate_val = candidate_rate_val
                candidate_rate_num = _row_float(r, rate_col, 0.0)
                if candidate_rate_num > 0:
                    rate_num = candidate_rate_num
                    break

            if not tax_name and rate_num > 0:
                tax_name = default_ledger
            if not tax_name:
                continue

            key = _ledger_name_key(tax_name)
            tax_type_map.setdefault(key, tax_type)
            if rate_val and key not in tax_rate_map:
                tax_rate_map[key] = rate_val

    entries = []
    seen = set()
    for ledger_name in missing_names:
        if _is_effectively_blank_ledger(ledger_name):
            continue
        key = _ledger_name_key(ledger_name)
        if key in seen:
            continue
        seen.add(key)

        tax_type = ""
        gst_rate = ""
        if key in party_keys:
            party_entry = dict(party_defs.get(key) or {})
            if party_entry:
                party_entry["Name"] = party_entry.get("Name") or ledger_name
                entries.append(party_entry)
                continue
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
                tax_type = _infer_tax_type_from_ledger_name(ledger_name)
        else:
            parent = "Purchase Accounts" if is_purchase_mode else "Sales Accounts"

        entries.append(
            {
                "Name": ledger_name,
                "Parent": parent,
                "GSTApplicable": "",
                "GSTIN": "",
                "StateOfSupply": "",
                "Address1": "",
                "Address2": "",
                "MailingName": "",
                "Country": "India",
                "Pincode": "",
                "Billwise": "No",
                "GSTRegistrationType": "",
                "TypeOfTaxation": tax_type,
                "GSTRate": str(gst_rate).strip() if gst_rate not in (None, "") else "",
            }
        )

    return entries


def _extract_numeric_voucher_no(value):
    text = _normalize_voucher_number_text(value)
    if text.isdigit():
        return int(text)
    return None


def _normalize_voucher_number_text(value) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def _increment_voucher_number_text(value: str):
    text = _normalize_voucher_number_text(value)
    if not text:
        return None
    if text.isdigit():
        return str(int(text) + 1)

    match = re.search(r"(\d+)(?!.*\d)", text)
    if not match:
        return None

    digits = match.group(1)
    start, end = match.span(1)
    return f"{text[:start]}{str(int(digits) + 1).zfill(len(digits))}{text[end:]}"


def _voucher_number_with_offset(start_value, offset: int) -> str:
    text = _normalize_voucher_number_text(start_value)
    if not text:
        return ""
    if offset <= 0:
        return text
    if text.isdigit():
        return str(int(text) + offset)

    current = text
    for idx in range(offset):
        next_value = _increment_voucher_number_text(current)
        if not next_value:
            return f"{text}-{idx + 2}"
        current = next_value
    return current


def _extract_voucher_records(response_text: str) -> list:
    records = []

    try:
        root = ET.fromstring(response_text)
    except ET.ParseError:
        return records

    for voucher in root.findall(".//VOUCHER"):
        voucher_type_name = str(
            voucher.findtext("VOUCHERTYPENAME")
            or voucher.attrib.get("VCHTYPE")
            or ""
        ).strip()
        voucher_number = _normalize_voucher_number_text(voucher.findtext("VOUCHERNUMBER"))
        voucher_date = str(voucher.findtext("DATE") or "").strip()
        master_id = str(voucher.findtext("MASTERID") or "").strip()
        if not voucher_type_name and not voucher_number:
            continue
        records.append(
            {
                "voucher_type": voucher_type_name,
                "voucher_number": voucher_number,
                "date": voucher_date,
                "master_id": master_id,
            }
        )

    return records


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
    voucher_type_text = str(voucher_type or "Sales").strip() or "Sales"
    escaped_voucher_type = xml_escape(voucher_type_text)

    static_vars = "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    static_vars += f"<SVVOUCHERTYPENAME>{escaped_voucher_type}</SVVOUCHERTYPENAME>"
    if selected_company:
        static_vars += f"<SVCURRENTCOMPANY>{xml_escape(selected_company)}</SVCURRENTCOMPANY>"
    static_vars += "</STATICVARIABLES>"

    request_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Voucher Number Collection</ID></HEADER><BODY><DESC>"
        f"{static_vars}"
        "<TDL><TDLMESSAGE><COLLECTION NAME='Voucher Number Collection'>"
        "<TYPE>Voucher</TYPE><FETCH>VoucherNumber,VoucherTypeName,Date,MasterID</FETCH>"
        "<FILTERS>ExactVoucherType</FILTERS>"
        "</COLLECTION>"
        f"<SYSTEM TYPE='Formulae' NAME='ExactVoucherType'>$VoucherTypeName = &quot;{escaped_voucher_type}&quot;</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

    try:
        response_text = _post_tally_xml(tally_url, request_xml, timeout=timeout)
    except Exception as exc:
        return {
            "success": False,
            "last_number": "",
            "next_number": "",
            "existing_numbers": set(),
            "error": str(exc) or "Could not fetch voucher number from Tally.",
        }

    records = [
        record
        for record in _extract_voucher_records(response_text)
        if _ledger_name_key(record.get("voucher_type", "")) == _ledger_name_key(voucher_type_text)
    ]
    existing_numbers = {
        _normalize_voucher_number_text(record.get("voucher_number", ""))
        for record in records
        if _normalize_voucher_number_text(record.get("voucher_number", ""))
    }

    if records:
        latest_record = max(
            records,
            key=lambda record: (
                _safe_int(record.get("date", 0), 0),
                _safe_int(record.get("master_id", 0), 0),
            ),
        )
        last_number = _normalize_voucher_number_text(latest_record.get("voucher_number", ""))
        next_number = _increment_voucher_number_text(last_number)
        if not next_number:
            numeric_numbers = [
                _extract_numeric_voucher_no(record.get("voucher_number", ""))
                for record in records
            ]
            numeric_numbers = [num for num in numeric_numbers if num is not None]
            if numeric_numbers:
                next_number = str(max(numeric_numbers) + 1)
        if not next_number:
            next_number = "1"
        return {
            "success": True,
            "last_number": last_number,
            "next_number": next_number,
            "existing_numbers": existing_numbers,
            "error": "",
        }

    if response_text.strip():
        return {
            "success": True,
            "last_number": "",
            "next_number": "1",
            "existing_numbers": existing_numbers,
            "error": "",
        }

    return {
        "success": False,
        "last_number": "",
        "next_number": "",
        "existing_numbers": set(),
        "error": "Could not fetch voucher number from Tally.",
    }


# ═══════════════════════════════════════════════════════════════════════════
#  GENERATE SALES XML  –  ACCOUNTING MODE  (mirrors original VBA logic)
# ═══════════════════════════════════════════════════════════════════════════

def generate_accounting_xml(
    rows: list,
    company: str,
    use_today_date: bool = False,
    date_mode: str = "",
    custom_tally_date: str = "",
    start_voucher_number=None,
    company_gst_registrations: list = None,
) -> str:
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

    resolved_mode = str(date_mode or ("current" if use_today_date else "excel")).strip().lower()
    if resolved_mode not in {"current", "excel", "custom"}:
        resolved_mode = "current" if use_today_date else "excel"
    resolved_custom_date = _normalize_manual_date_to_tally(custom_tally_date) if resolved_mode == "custom" else ""

    resolved_mode = str(date_mode or ("current" if use_today_date else "excel")).strip().lower()
    if resolved_mode not in {"current", "excel", "custom"}:
        resolved_mode = "current" if use_today_date else "excel"
    resolved_custom_date = _normalize_manual_date_to_tally(custom_tally_date) if resolved_mode == "custom" else ""

    for idx, r in enumerate(rows):
        if resolved_mode == "current":
            source_date = datetime.today()
        elif resolved_mode == "custom":
            source_date = resolved_custom_date
        else:
            source_date = _row_get(r, "Date", "")
        dt       = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno      = xml_escape(vno_raw)
        invoice_ref_raw = _row_invoice_reference(r, vno_raw)
        invoice_ref = xml_escape(invoice_ref_raw)
        party_raw = _ledger_or_suspense(_row_text(r, "PartyLedger"))
        party_context = _collect_party_context(r, party_raw)
        sales_raw = _ledger_or_suspense(_row_text(r, "SalesLedger"))
        party = xml_escape(party_raw)
        sales = xml_escape(sales_raw)
        taxable  = _row_float(r, "TaxableValue", 0.0)
        cgst_r   = _row_float(r, "CGSTRate", 0.0)
        cgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["CGSTLedger", "CGST Ledger", "CentralTaxLedger", "Central Tax Ledger", "Central Tax"],
            cgst_r,
            "CGST",
        ))
        sgst_r   = _row_float(r, "SGSTRate", 0.0)
        sgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["SGSTLedger", "SGST Ledger", "StateTaxLedger", "State Tax Ledger", "State Tax", "UTGSTLedger", "UTGST Ledger"],
            sgst_r,
            "SGST",
        ))
        igst_r   = _row_float(r, "IGSTRate", 0.0)
        igst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["IGSTLedger", "IGST Ledger", "IntegratedTaxLedger", "Integrated Tax Ledger", "Integrated Tax"],
            igst_r,
            "IGST",
        ))
        narr     = xml_escape(_row_text(r, "Narration"))

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
        _append_invoice_party_context_xml(a, party_context, include_basic_buyer=True)
        _append_company_gst_context_xml(a, party_context, company_gst_registrations)
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>')
        a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>')
        # GSTTRANSACTIONTYPE is required to avoid 'Uncertain' in GSTR-1
        _gst_txn_type = _gst_transaction_type(
            party_context.get("registration_type", ""),
            party_context.get("gstin", ""),
        )
        a(f'     <GSTTRANSACTIONTYPE>{xml_escape(_gst_txn_type)}</GSTTRANSACTIONTYPE>')
        if invoice_ref:
            a(f'     <REFERENCE>{invoice_ref}</REFERENCE>')
        if narr:
            a(f'     <NARRATION>{narr}</NARRATION>')

        # Party – Debit (with bill allocation so Tally links to GSTR-1)
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{party}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>-{fmt_amt(total)}</AMOUNT>')
        a('      <BILLALLOCATIONS.LIST>')
        a(f'       <NAME>{invoice_ref or vno}</NAME>')
        a('       <BILLTYPE>New Ref</BILLTYPE>')
        a(f'       <AMOUNT>-{fmt_amt(total)}</AMOUNT>')
        a('      </BILLALLOCATIONS.LIST>')
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
            _append_tax_object_allocation_xml(a, "CGST")
            a('     </LEDGERENTRIES.LIST>')
        # SGST
        if sgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{sgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(sgst_amt)}</AMOUNT>')
            _append_tax_object_allocation_xml(a, "SGST")
            a('     </LEDGERENTRIES.LIST>')
        # IGST
        if igst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{igst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>{fmt_amt(igst_amt)}</AMOUNT>')
            _append_tax_object_allocation_xml(a, "IGST")
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
    date_mode: str = "",
    custom_tally_date: str = "",
    start_voucher_number=None,
    fallback_sales_ledger: str = SUSPENSE_LEDGER,
    company_gst_registrations: list = None,
    legacy_invoice_context: bool = False,
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

    resolved_mode = str(date_mode or ("current" if use_today_date else "excel")).strip().lower()
    if resolved_mode not in {"current", "excel", "custom"}:
        resolved_mode = "current" if use_today_date else "excel"
    resolved_custom_date = _normalize_manual_date_to_tally(custom_tally_date) if resolved_mode == "custom" else ""

    def _name_key(value: str) -> str:
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    for idx, r in enumerate(rows):
        if resolved_mode == "current":
            source_date = datetime.today()
        elif resolved_mode == "custom":
            source_date = resolved_custom_date
        else:
            source_date = _row_get(r, "Date", "")
        dt       = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno      = xml_escape(vno_raw)
        invoice_ref_raw = _row_invoice_reference(r, vno_raw)
        invoice_ref = xml_escape(invoice_ref_raw)
        party_raw = _ledger_or_suspense(_row_text(r, "PartyLedger"))
        party_context = _collect_party_context(r, party_raw)
        party = xml_escape(party_raw)
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

        cgst_r   = _row_float(r, "CGSTRate", 0.0)
        cgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["CGSTLedger", "CGST Ledger", "CentralTaxLedger", "Central Tax Ledger", "Central Tax"],
            cgst_r,
            "CGST",
        ))
        sgst_r   = _row_float(r, "SGSTRate", 0.0)
        sgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["SGSTLedger", "SGST Ledger", "StateTaxLedger", "State Tax Ledger", "State Tax", "UTGSTLedger", "UTGST Ledger"],
            sgst_r,
            "SGST",
        ))
        igst_r   = _row_float(r, "IGSTRate", 0.0)
        igst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["IGSTLedger", "IGST Ledger", "IntegratedTaxLedger", "Integrated Tax Ledger", "Integrated Tax"],
            igst_r,
            "IGST",
        ))
        narr     = xml_escape(_row_text(r, "Narration"))
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
        if legacy_invoice_context:
            party_gstin = xml_escape(party_context.get("gstin", ""))
            place_of_supply = xml_escape(party_context.get("place_of_supply", ""))
            if party_gstin:
                a(f'     <PARTYGSTIN>{party_gstin}</PARTYGSTIN>')
            if place_of_supply:
                a(f'     <PLACEOFSUPPLY>{place_of_supply}</PLACEOFSUPPLY>')
        else:
            _append_invoice_party_context_xml(a, party_context, include_basic_buyer=True)
            _append_company_gst_context_xml(a, party_context, company_gst_registrations)
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>')
        if not legacy_invoice_context:
            a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>')
            # GSTTRANSACTIONTYPE is required to avoid 'Uncertain' in GSTR-1
            _gst_txn_type = _gst_transaction_type(
                party_context.get("registration_type", ""),
                party_context.get("gstin", ""),
            )
            a(f'     <GSTTRANSACTIONTYPE>{xml_escape(_gst_txn_type)}</GSTTRANSACTIONTYPE>')
        if invoice_ref:
            a(f'     <REFERENCE>{invoice_ref}</REFERENCE>')
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

        # ── Ledger entries (Party DR) with bill allocation ──
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{party}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>-{fmt_amt(total)}</AMOUNT>')
        a('      <BILLALLOCATIONS.LIST>')
        a(f'       <NAME>{invoice_ref or vno}</NAME>')
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
    date_mode: str = "",
    custom_tally_date: str = "",
    start_voucher_number=None,
    company_gst_registrations: list = None,
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

    resolved_mode = str(date_mode or ("current" if use_today_date else "excel")).strip().lower()
    if resolved_mode not in {"current", "excel", "custom"}:
        resolved_mode = "current" if use_today_date else "excel"
    resolved_custom_date = _normalize_manual_date_to_tally(custom_tally_date) if resolved_mode == "custom" else ""
    purchase_company_state = _resolve_company_gst_state(company_gst_registrations or [], preferred_state="")

    for idx, r in enumerate(rows):
        if resolved_mode == "current":
            source_date = datetime.today()
        elif resolved_mode == "custom":
            source_date = resolved_custom_date
        else:
            source_date = _row_get(r, "Date", "")
        dt = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno = xml_escape(vno_raw)
        supplier_invoice_raw = _row_invoice_reference(r, vno_raw)
        supplier_invoice = xml_escape(supplier_invoice_raw)

        party_raw = _ledger_or_suspense(_row_text(r, "PartyLedger"))
        party_context = _collect_party_context(
            r,
            party_raw,
            allow_place_of_supply_column=False,
        )
        party = xml_escape(party_raw)
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
        cgst_r = _row_float(r, "CGSTRate", 0.0)
        cgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["CGSTLedger", "CGST Ledger", "CentralTaxLedger", "Central Tax Ledger", "Central Tax"],
            cgst_r,
            "CGST",
        ))
        sgst_r = _row_float(r, "SGSTRate", 0.0)
        sgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["SGSTLedger", "SGST Ledger", "StateTaxLedger", "State Tax Ledger", "State Tax", "UTGSTLedger", "UTGST Ledger"],
            sgst_r,
            "SGST",
        ))
        igst_r = _row_float(r, "IGSTRate", 0.0)
        igst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["IGSTLedger", "IGST Ledger", "IntegratedTaxLedger", "Integrated Tax Ledger", "Integrated Tax"],
            igst_r,
            "IGST",
        ))
        narr = xml_escape(_row_text(r, "Narration"))

        # TDS fields
        tds_ledger_raw = (
            _row_text(r, "TDSLedger")
            or _row_text(r, "TDS Ledger")
            or _row_text(r, "Tds Ledger")
        )
        tds_rate = _row_float(r, "TDSRate", 0.0) or _row_float(r, "TDS Rate", 0.0)
        tds_amount_raw = _row_float(r, "TDSAmount", 0.0) or _row_float(r, "TDS Amount", 0.0)
        if tds_ledger_raw and tds_amount_raw <= 0 and tds_rate > 0:
            tds_amount = round(taxable * tds_rate / 100, 2)
        else:
            tds_amount = abs(tds_amount_raw)
        tds_led = xml_escape(tds_ledger_raw)

        cgst_amt = round(taxable * cgst_r / 100, 2) if cgst_r > 0 else 0
        sgst_amt = round(taxable * sgst_r / 100, 2) if sgst_r > 0 else 0
        igst_amt = round(taxable * igst_r / 100, 2) if igst_r > 0 else 0
        total = taxable + cgst_amt + sgst_amt + igst_amt
        party_total = total + tds_amount if (tds_led and tds_amount > 0) else total

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a('    <VOUCHER VCHTYPE="Purchase" ACTION="Create" OBJVIEW="Invoice Voucher View">')
        a(f'     <DATE>{dt}</DATE>')
        a('     <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>')
        a(f'     <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>')
        a(f'     <PARTYLEDGERNAME>{party}</PARTYLEDGERNAME>')
        _append_invoice_party_context_xml(
            a,
            party_context,
            include_basic_buyer=False,
            include_place_of_supply=bool(purchase_company_state),
            place_of_supply_override=purchase_company_state,
        )
        _append_company_gst_context_xml(
            a,
            party_context,
            company_gst_registrations,
            prefer_party_state=False,
        )
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>')
        a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>')
        # GSTTRANSACTIONTYPE required to avoid 'Uncertain' in GSTR-3B
        _gst_txn_type = _gst_transaction_type(
            party_context.get("registration_type", ""),
            party_context.get("gstin", ""),
        )
        a(f'     <GSTTRANSACTIONTYPE>{xml_escape(_gst_txn_type)}</GSTTRANSACTIONTYPE>')
        if supplier_invoice:
            a(f'     <REFERENCE>{supplier_invoice}</REFERENCE>')
        if narr:
            a(f'     <NARRATION>{narr}</NARRATION>')

        # Party - Credit (net of TDS)
        a('     <LEDGERENTRIES.LIST>')
        a(f'      <LEDGERNAME>{party}</LEDGERNAME>')
        a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
        a(f'      <AMOUNT>{fmt_amt(party_total)}</AMOUNT>')
        a('      <BILLALLOCATIONS.LIST>')
        a(f'       <NAME>{supplier_invoice or vno}</NAME>')
        a('       <BILLTYPE>New Ref</BILLTYPE>')
        a(f'       <AMOUNT>{fmt_amt(party_total)}</AMOUNT>')
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
            _append_tax_object_allocation_xml(a, "CGST")
            a('     </LEDGERENTRIES.LIST>')
        if sgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{sgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(sgst_amt)}</AMOUNT>')
            _append_tax_object_allocation_xml(a, "SGST")
            a('     </LEDGERENTRIES.LIST>')
        if igst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{igst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(igst_amt)}</AMOUNT>')
            _append_tax_object_allocation_xml(a, "IGST")
            a('     </LEDGERENTRIES.LIST>')

        # TDS Payable - shown as negative deduction in Invoice Amount column
        if tds_led and tds_amount > 0:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{tds_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(tds_amount)}</AMOUNT>')
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
    date_mode: str = "",
    custom_tally_date: str = "",
    start_voucher_number=None,
    fallback_purchase_ledger: str = SUSPENSE_LEDGER,
    company_gst_registrations: list = None,
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

    resolved_mode = str(date_mode or ("current" if use_today_date else "excel")).strip().lower()
    if resolved_mode not in {"current", "excel", "custom"}:
        resolved_mode = "current" if use_today_date else "excel"
    resolved_custom_date = _normalize_manual_date_to_tally(custom_tally_date) if resolved_mode == "custom" else ""
    purchase_company_state = _resolve_company_gst_state(company_gst_registrations or [], preferred_state="")

    def _name_key(value: str) -> str:
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    for idx, r in enumerate(rows):
        if resolved_mode == "current":
            source_date = datetime.today()
        elif resolved_mode == "custom":
            source_date = resolved_custom_date
        else:
            source_date = _row_get(r, "Date", "")
        dt = tally_date(source_date)
        if start_voucher_number is not None:
            vno_raw = str(int(start_voucher_number) + idx)
        else:
            vno_raw = _row_voucher_number(r)
        vno = xml_escape(vno_raw)
        supplier_invoice_raw = _row_invoice_reference(r, vno_raw)
        supplier_invoice = xml_escape(supplier_invoice_raw)

        party_raw = _ledger_or_suspense(_row_text(r, "PartyLedger"))
        party_context = _collect_party_context(
            r,
            party_raw,
            allow_place_of_supply_column=False,
        )
        party = xml_escape(party_raw)
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

        cgst_r = _row_float(r, "CGSTRate", 0.0)
        cgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["CGSTLedger", "CGST Ledger", "CentralTaxLedger", "Central Tax Ledger", "Central Tax"],
            cgst_r,
            "CGST",
        ))
        sgst_r = _row_float(r, "SGSTRate", 0.0)
        sgst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["SGSTLedger", "SGST Ledger", "StateTaxLedger", "State Tax Ledger", "State Tax", "UTGSTLedger", "UTGST Ledger"],
            sgst_r,
            "SGST",
        ))
        igst_r = _row_float(r, "IGSTRate", 0.0)
        igst_led = xml_escape(_pick_tax_ledger_name(
            r,
            ["IGSTLedger", "IGST Ledger", "IntegratedTaxLedger", "Integrated Tax Ledger", "Integrated Tax"],
            igst_r,
            "IGST",
        ))
        narr = xml_escape(_row_text(r, "Narration"))

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
        _append_invoice_party_context_xml(
            a,
            party_context,
            include_basic_buyer=False,
            include_place_of_supply=bool(purchase_company_state),
            place_of_supply_override=purchase_company_state,
        )
        _append_company_gst_context_xml(
            a,
            party_context,
            company_gst_registrations,
            prefer_party_state=False,
        )
        a(f'     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>')
        a('     <ISINVOICE>Yes</ISINVOICE>')
        a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>')
        a('     <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>')
        a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>')
        # GSTTRANSACTIONTYPE required to avoid 'Uncertain' in GSTR-3B
        _gst_txn_type = _gst_transaction_type(
            party_context.get("registration_type", ""),
            party_context.get("gstin", ""),
        )
        a(f'     <GSTTRANSACTIONTYPE>{xml_escape(_gst_txn_type)}</GSTTRANSACTIONTYPE>')
        if supplier_invoice:
            a(f'     <REFERENCE>{supplier_invoice}</REFERENCE>')
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

        # Party - Credit (supplier owes us goods; positive amount = credit in Tally)
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

        # Taxes - Debit (ITC claim; ISDEEMEDPOSITIVE=Yes, negative amount = debit)
        if cgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{cgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(cgst_amt)}</AMOUNT>')
            _append_tax_object_allocation_xml(a, "CGST")
            a('     </LEDGERENTRIES.LIST>')
        if sgst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{sgst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(sgst_amt)}</AMOUNT>')
            _append_tax_object_allocation_xml(a, "SGST")
            a('     </LEDGERENTRIES.LIST>')
        if igst_amt:
            a('     <LEDGERENTRIES.LIST>')
            a(f'      <LEDGERNAME>{igst_led}</LEDGERNAME>')
            a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            a(f'      <AMOUNT>-{fmt_amt(igst_amt)}</AMOUNT>')
            _append_tax_object_allocation_xml(a, "IGST")
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

def generate_ledger_xml(ledgers: list, company: str, alter_existing: bool = False) -> str:
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
        name_raw = str(led.get("Name", "") or "").strip()
        if not name_raw:
            continue
        parent_raw = str(led.get("Parent", "Sundry Debtors") or "Sundry Debtors").strip()
        gstin_raw = str(led.get("GSTIN", "") or "").strip().upper()
        state_raw = _normalize_state_for_ledger(
            str(led.get("StateOfSupply", "") or "").strip()
            or str(led.get("PlaceOfSupply", "") or "").strip()
            or str(led.get("State", "") or "").strip()
        )
        address1_raw = (
            str(led.get("Address1", "") or "").strip()
            or str(led.get("AddressLine1", "") or "").strip()
            or str(led.get("Address Line 1", "") or "").strip()
        )
        address2_raw = (
            str(led.get("Address2", "") or "").strip()
            or str(led.get("AddressLine2", "") or "").strip()
            or str(led.get("Address Line 2", "") or "").strip()
        )
        pincode_raw = (
            str(led.get("Pincode", "") or "").strip()
            or str(led.get("PinCode", "") or "").strip()
            or str(led.get("PIN", "") or "").strip()
            or str(led.get("PostalCode", "") or "").strip()
        )
        country_raw = (
            str(led.get("Country", "") or "").strip()
            or str(led.get("CountryOfResidence", "") or "").strip()
            or "India"
        )
        mailing_name_raw = (
            str(led.get("MailingName", "") or "").strip()
            or str(led.get("PartyMailingName", "") or "").strip()
            or name_raw
        )
        if not state_raw and gstin_raw:
            state_raw = _state_name_from_gstin(gstin_raw)

        is_party = _is_party_parent(parent_raw)
        is_duties_ledger = _is_duties_parent(parent_raw)
        billwise_raw = (
            str(led.get("Billwise", "") or "").strip()
            or str(led.get("IsBillwise", "") or "").strip()
            or str(led.get("ISBILLWISEON", "") or "").strip()
        )
        if billwise_raw:
            billwise_on = billwise_raw.casefold() in {"yes", "y", "true", "1", "on"}
        else:
            billwise_on = bool(is_party)

        gst_app_raw = str(led.get("GSTApplicable", "") or "").strip()
        gst_app_value = _normalize_gst_applicable(gst_app_raw, gstin=gstin_raw)
        reg_type_raw = (
            str(led.get("GSTRegistrationType", "") or "").strip()
            or str(led.get("RegistrationType", "") or "").strip()
            or str(led.get("RegType", "") or "").strip()
        )
        reg_type = _normalize_gst_registration_type(
            reg_type_raw,
            gstin=gstin_raw,
            gst_applicable=gst_app_value,
        )

        name = xml_escape(name_raw)
        parent = xml_escape(parent_raw)
        gst_app = xml_escape(gst_app_value)
        gstin = xml_escape(gstin_raw)
        state = xml_escape(state_raw)
        address1 = xml_escape(address1_raw)
        address2 = xml_escape(address2_raw)
        pincode = xml_escape(pincode_raw)
        country = xml_escape(country_raw)
        mailing_name = xml_escape(mailing_name_raw)
        tax_type_raw = str(led.get("TypeOfTaxation", "") or "").strip()
        if tax_type_raw.casefold() in {"not applicable", "na", "n/a"}:
            tax_type_raw = ""
        if is_duties_ledger and not tax_type_raw:
            tax_type_raw = _infer_tax_type_from_ledger_name(name_raw)
        is_gst_duty_ledger = _is_protected_gst_tax_ledger(name_raw, parent_raw, tax_type_raw)
        tax_type = xml_escape(tax_type_raw)
        gst_rate = str(led.get("GSTRate", "") or "").strip()
        applicable_from = _current_fy_start()
        ledger_action = "Create Alter" if alter_existing else "Create"

        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a(f'    <LEDGER NAME="{name}" RESERVEDNAME="" ACTION="{ledger_action}">')
        a(f'     <NAME>{name}</NAME>')
        a(f'     <PARENT>{parent}</PARENT>')
        a(f'     <ISBILLWISEON>{"Yes" if billwise_on else "No"}</ISBILLWISEON>')
        a('     <ISCOSTCENTRESON>No</ISCOSTCENTRESON>')
        a('     <ISINTERESTON>No</ISINTERESTON>')
        a('     <ALLOWINMOBILE>No</ALLOWINMOBILE>')
        a('     <ISUPDATINGTARGETID>No</ISUPDATINGTARGETID>')
        a('     <ASORIGINAL>Yes</ASORIGINAL>')
        a('     <AFFECTSSTOCK>No</AFFECTSSTOCK>')
        a('     <CURRENCYNAME>INR</CURRENCYNAME>')
        a(f'     <COUNTRYOFRESIDENCE>{country}</COUNTRYOFRESIDENCE>')

        if is_party and gst_app:
            a(f'     <GSTAPPLICABLE>{gst_app}</GSTAPPLICABLE>')
        if is_party and reg_type:
            a(f'     <GSTREGISTRATIONTYPE>{xml_escape(reg_type)}</GSTREGISTRATIONTYPE>')
        if is_party and gstin:
            a(f'     <PARTYGSTIN>{gstin}</PARTYGSTIN>')
        if state:
            a(f'     <PRIORSTATENAME>{state}</PRIORSTATENAME>')
            if is_party:
                a(f'     <LEDSTATENAME>{state}</LEDSTATENAME>')

        a('     <LANGUAGENAME.LIST>')
        a('      <NAME.LIST TYPE="String">')
        a(f'       <NAME>{name}</NAME>')
        a('      </NAME.LIST>')
        a('      <LANGUAGEID>1033</LANGUAGEID>')
        a('     </LANGUAGENAME.LIST>')

        if is_party and (gstin or reg_type):
            a('     <LEDGSTREGDETAILS.LIST>')
            a(f'      <APPLICABLEFROM>{applicable_from}</APPLICABLEFROM>')
            if reg_type:
                a(f'      <GSTREGISTRATIONTYPE>{xml_escape(reg_type)}</GSTREGISTRATIONTYPE>')
            if state:
                a(f'      <PLACEOFSUPPLY>{state}</PLACEOFSUPPLY>')
            if gstin:
                a(f'      <GSTIN>{gstin}</GSTIN>')
            a('      <ISOTHTERRITORYASSESSEE>No</ISOTHTERRITORYASSESSEE>')
            a('      <CONSIDERPURCHASEFOREXPORT>No</CONSIDERPURCHASEFOREXPORT>')
            a('      <ISTRANSPORTER>No</ISTRANSPORTER>')
            a('      <ISCOMMONPARTY>No</ISCOMMONPARTY>')
            a('     </LEDGSTREGDETAILS.LIST>')

        if is_party and (state or gstin or address1 or address2 or country or pincode):
            a('     <LEDMAILINGDETAILS.LIST>')
            if address1 or address2:
                a('      <ADDRESS.LIST TYPE="String">')
                if address1:
                    a(f'       <ADDRESS>{address1}</ADDRESS>')
                if address2:
                    a(f'       <ADDRESS>{address2}</ADDRESS>')
                a('      </ADDRESS.LIST>')
            a(f'      <APPLICABLEFROM>{applicable_from}</APPLICABLEFROM>')
            if pincode:
                a(f'      <PINCODE>{pincode}</PINCODE>')
            a(f'      <MAILINGNAME>{mailing_name}</MAILINGNAME>')
            if state:
                a(f'      <STATE>{state}</STATE>')
            a(f'      <COUNTRY>{country}</COUNTRY>')
            a('     </LEDMAILINGDETAILS.LIST>')

        if is_duties_ledger:
            # In Tally ledger XML, Type of Duty/Tax is controlled by TAXTYPE.
            # For GST duty ledgers (IGST/CGST/SGST/etc.), this must be GST.
            if is_gst_duty_ledger:
                a('     <TAXTYPE>GST</TAXTYPE>')
            elif tax_type:
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
#  JOURNAL ENTRY HELPERS & XML GENERATION
# ═══════════════════════════════════════════════════════════════════════════

def _ledger_or_default(value: str, fallback: str = SUSPENSE_LEDGER) -> str:
    text = str(value or "").strip()
    return text or fallback


def _state_from_gstin(g):
    M = {
        '01': 'Jammu and Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab', '04': 'Chandigarh',
        '05': 'Uttarakhand', '06': 'Haryana', '07': 'Delhi', '08': 'Rajasthan', '09': 'Uttar Pradesh',
        '10': 'Bihar', '11': 'Sikkim', '12': 'Arunachal Pradesh', '13': 'Nagaland', '14': 'Manipur',
        '15': 'Mizoram', '16': 'Tripura', '17': 'Meghalaya', '18': 'Assam', '19': 'West Bengal',
        '20': 'Jharkhand', '21': 'Odisha', '22': 'Chattisgarh', '23': 'Madhya Pradesh', '24': 'Gujarat',
        '26': 'Dadra and Nagar Haveli and Daman and Diu', '27': 'Maharashtra',
        '28': 'Andhra Pradesh', '29': 'Karnataka', '30': 'Goa', '31': 'Lakshadweep', '32': 'Kerala',
        '33': 'Tamil Nadu', '34': 'Puducherry', '35': 'Andaman and Nicobar Islands',
        '36': 'Telangana', '37': 'Andhra Pradesh (New)', '38': 'Ladakh', '97': 'Other Territory',
    }
    return M.get(str(g or '')[:2], '')


def _fetch_tally_ledgers(tally_url: str, timeout: float = 15.0, company_name: str = "") -> dict:
    selected_company = _normalize_company_name(company_name)
    static_vars = "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    if selected_company:
        static_vars += f"<SVCURRENTCOMPANY>{xml_escape(selected_company)}</SVCURRENTCOMPANY>"
    static_vars += "</STATICVARIABLES>"

    request_xml_variants = [
        (
            "collection-ledger",
            "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
            "<TYPE>Collection</TYPE><ID>Ledger Collection</ID></HEADER>"
            f"<BODY><DESC>{static_vars}<TDL><TDLMESSAGE><COLLECTION NAME='Ledger Collection'>"
            "<TYPE>Ledger</TYPE><FETCH>Name,Parent</FETCH><NATIVEMETHOD>Name</NATIVEMETHOD>"
            "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>",
        ),
        (
            "report-list-ledgers",
            "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>"
            "<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Ledgers</REPORTNAME>"
            f"{static_vars}</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>",
        ),
    ]

    def _extract_from_response(response_text: str):
        all_entries = []
        try:
            root = ET.fromstring(response_text)
            for ledger_node in root.findall(".//LEDGER"):
                name = _normalize_ledger_name(
                    ledger_node.attrib.get("NAME")
                    or ledger_node.findtext("NAME")
                )
                parent = _normalize_ledger_name(
                    ledger_node.attrib.get("PARENT")
                    or ledger_node.findtext("PARENT")
                )
                if name:
                    all_entries.append((name, parent))
        except ET.ParseError:
            pass

        for match in re.findall(r'LEDGER[^>]*NAME="([^"]+)"', response_text, flags=re.IGNORECASE):
            name = _normalize_ledger_name(match)
            if name:
                all_entries.append((name, ""))
        return all_entries

    all_ledgers_map = {}
    errors = []

    for label, payload in request_xml_variants:
        try:
            response_text = _post_tally_xml(tally_url, payload, timeout=timeout)
            entries = _extract_from_response(response_text)
            for name, parent in entries:
                key = _ledger_key(name)
                if key not in all_ledgers_map:
                    all_ledgers_map[key] = {"name": name, "parent": parent}
                elif not all_ledgers_map[key].get("parent") and parent:
                    all_ledgers_map[key]["parent"] = parent
        except HTTPError as exc:
            errors.append(f"{label}: HTTP {exc.code}")
        except URLError:
            errors.append(f"{label}: ConnectionError")
        except Exception as exc:
            errors.append(f"{label}: {exc}")

    ledgers = sorted((v["name"] for v in all_ledgers_map.values() if v.get("name")), key=lambda x: _ledger_key(x))

    party_group_keys = {"SUNDRY CREDITORS", "SUNDRY DEBTORS"}
    party_ledgers = sorted(
        (
            v["name"]
            for v in all_ledgers_map.values()
            if v.get("name") and re.sub(r"\s+", " ", str(v.get("parent") or "")).strip().upper() in party_group_keys
        ),
        key=lambda x: _ledger_key(x),
    )

    if party_ledgers or ledgers:
        return {
            "success": True,
            "ledgers": ledgers,
            "party_ledgers": party_ledgers,
        }

    err = "; ".join(errors) if errors else "No ledgers returned by Tally."
    return {
        "success": False,
        "error": err,
        "ledgers": [],
        "party_ledgers": [],
    }


def _fetch_cmp_gst_regs_for_journal(url, company='', timeout=15):
    sc = f'<SVCURRENTCOMPANY>{xml_escape(company)}</SVCURRENTCOMPANY>' if company else ''
    col = ("<COLLECTION NAME='TUL'><TYPE>TaxUnit</TYPE>"
           "<FETCH>Name,TaxType,TaxRegistration,GSTRegNumber,StateName</FETCH></COLLECTION>")
    rq = (
        f'<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
        f'<TYPE>Collection</TYPE><ID>TUL</ID></HEADER>'
        f'<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{sc}'
        f'</STATICVARIABLES><TDL><TDLMESSAGE>{col}</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'
    )
    try:
        r = _post_tally_xml(url, rq, timeout=timeout)
        regs = []
        for m in re.finditer(r'<TAXUNIT[^>]*>(.*?)</TAXUNIT>', r, re.DOTALL):
            blk = m.group(1)
            def _t(tag, b=blk):
                x = re.search(fr'<{tag}[^>]*>(.*?)</{tag}>', b, re.DOTALL)
                return x.group(1).strip() if x else ''
            gstin = (_t('GSTREGNUMBER') or _t('TAXREGISTRATION') or '').upper()
            state = _t('STATENAME') or ''
            name = _t('NAME') or ''
            if not gstin:
                continue
            if not state:
                state = _state_from_gstin(gstin)
            regs.append({'gstin': gstin, 'state': state, 'name': name})
        return {'success': bool(regs), 'registrations': regs}
    except Exception as e:
        return {'success': False, 'registrations': [], 'error': str(e)}


def _clean_tax_ledger(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if text.casefold() in {"0", "0.0", "none", "na", "n/a", "-"}:
        return ""
    return text


def _row_reference_number(row: dict, default: str = "") -> str:
    return (
        _row_text(row, "Reference")
        or _row_text(row, "RefNo")
        or _row_text(row, "Ref No")
        or _row_text(row, "BillRef")
        or _row_text(row, "BillNo")
        or _row_text(row, "Bill No")
        or _row_text(row, "InvoiceNo")
        or _row_text(row, "Invoice No")
        or _row_text(row, "SupplierInvoiceNo")
        or _row_text(row, "Supplier Invoice No")
        or default
    )


def _append_common_ledger_flags(add_line, is_party: bool, is_debit: bool = None) -> None:
    if is_debit is None:
        is_debit = not is_party
    add_line("      <GSTCLASS>Not Applicable</GSTCLASS>")
    add_line(f"      <ISDEEMEDPOSITIVE>{'Yes' if is_debit else 'No'}</ISDEEMEDPOSITIVE>")
    add_line("      <LEDGERFROMITEM>No</LEDGERFROMITEM>")
    add_line("      <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>")
    add_line(f"      <ISPARTYLEDGER>{'Yes' if is_party else 'No'}</ISPARTYLEDGER>")
    add_line("      <GSTOVERRIDDEN>No</GSTOVERRIDDEN>")
    add_line("      <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>")


def generate_journal_xml(
    rows: list,
    company: str,
    use_today_date: bool = False,
    date_mode: str = "",
    custom_tally_date: str = "",
    include_voucher_number: bool = True,
    include_bill_allocations: bool = True,
    journal_type: str = "purchase",
    company_gst_registrations: list = None,
) -> tuple:
    """
    Purchase: Expense Dr / Tax Dr / Party Cr
    Sale:     Party Dr  / Sales Cr / Tax Cr
    GST State/Country fetched by Tally from party ledger master via GSTREGISTRATIONTYPE=Regular.
    """
    is_sale = str(journal_type or "purchase").strip().lower() == "sale"
    lines = []
    a = lines.append
    company_static = _company_static_block(company)

    a('<?xml version="1.0" encoding="UTF-8"?>')
    a("<ENVELOPE>")
    a(" <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>")
    a(" <BODY><IMPORTDATA>")
    a("  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>")
    if company_static:
        a(company_static)
    a("  </REQUESTDESC>")
    a("  <REQUESTDATA>")

    resolved_mode = str(date_mode or ("current" if use_today_date else "excel")).strip().lower()
    if resolved_mode not in {"current", "excel", "custom"}:
        resolved_mode = "current" if use_today_date else "excel"
    resolved_custom_date = _normalize_manual_date_to_tally(custom_tally_date) if resolved_mode == "custom" else ""

    _cmp_regs = list(company_gst_registrations or [])
    _cmp_gstin = _cmp_state = _cmp_name = ""
    if _cmp_regs:
        _r = _cmp_regs[0]
        _cmp_gstin = xml_escape(str(_r.get("gstin", "")).strip())
        _cmp_state = xml_escape(str(_r.get("state", "")).strip())
        _cmp_name = xml_escape(str(_r.get("name", "")).strip())

    voucher_count = 0

    for idx, r in enumerate(rows):
        taxable = _row_float(r, "TaxableValue", 0.0)
        if taxable <= 0:
            continue

        if resolved_mode == "current":
            source_date = datetime.today()
        elif resolved_mode == "custom":
            source_date = resolved_custom_date
        else:
            source_date = _row_get(r, "Date", "")
        dt = tally_date(source_date)
        vno_raw = _row_voucher_number(r, "")

        party_raw = _ledger_or_default(_row_text(r, "PartyLedger"))
        particular_raw = (
            _row_text(r, "Particular")
            or _row_text(r, "Particulars")
            or _row_text(r, "ExpenseLedger")
            or _row_text(r, "PurchaseLedger")
            or "Journal Adjustment"
        )
        particular_raw = _ledger_or_default(particular_raw, "Journal Adjustment")

        cgst_ledger_raw = _clean_tax_ledger(_row_text(r, "CGSTLedger"))
        sgst_ledger_raw = _clean_tax_ledger(_row_text(r, "SGSTLedger"))
        igst_ledger_raw = _clean_tax_ledger(_row_text(r, "IGSTLedger"))

        cgst_rate = _row_float(r, "CGSTRate", 0.0)
        sgst_rate = _row_float(r, "SGSTRate", 0.0)
        igst_rate = _row_float(r, "IGSTRate", 0.0)

        cgst_amt = round(taxable * cgst_rate / 100, 2) if cgst_rate > 0 and cgst_ledger_raw else 0.0
        sgst_amt = round(taxable * sgst_rate / 100, 2) if sgst_rate > 0 and sgst_ledger_raw else 0.0
        igst_amt = round(taxable * igst_rate / 100, 2) if igst_rate > 0 and igst_ledger_raw else 0.0
        total = taxable + cgst_amt + sgst_amt + igst_amt
        has_gst = (cgst_amt + sgst_amt + igst_amt) > 0

        # TDS fields (only applied for Purchase journal)
        jnl_tds_ledger_raw = (
            _row_text(r, "TDSLedger")
            or _row_text(r, "TDS Ledger")
            or _row_text(r, "Tds Ledger")
        ) if not is_sale else ""
        jnl_tds_rate = (_row_float(r, "TDSRate", 0.0) or _row_float(r, "TDS Rate", 0.0)) if not is_sale else 0.0
        jnl_tds_amount_raw = (_row_float(r, "TDSAmount", 0.0) or _row_float(r, "TDS Amount", 0.0)) if not is_sale else 0.0
        if jnl_tds_ledger_raw and jnl_tds_amount_raw <= 0 and jnl_tds_rate > 0:
            jnl_tds_amount = round(taxable * jnl_tds_rate / 100, 2)
        else:
            jnl_tds_amount = abs(jnl_tds_amount_raw)
        jnl_tds_led = xml_escape(jnl_tds_ledger_raw)
        jnl_party_total = total - jnl_tds_amount if (jnl_tds_led and jnl_tds_amount > 0) else total

        bill_reference_raw = _row_reference_number(r, "")
        voucher_reference_raw = bill_reference_raw or (vno_raw if include_voucher_number else "")
        vno = xml_escape(vno_raw)
        reference = xml_escape(voucher_reference_raw)
        bill_reference = xml_escape(bill_reference_raw)
        party = xml_escape(party_raw)
        particular = xml_escape(particular_raw)
        narration = xml_escape(_row_text(r, "Narration"))

        party_gstin_raw = str(_row_text(r, "PartyGSTIN") or _row_text(r, "GSTIN") or "").strip().upper()
        pos_raw = str(_row_text(r, "PlaceOfSupply") or "").strip()
        party_state_raw = pos_raw or (_state_from_gstin(party_gstin_raw) if party_gstin_raw else "")
        place_of_supply = xml_escape(party_state_raw if is_sale else (pos_raw or _cmp_state))
        party_gstin = xml_escape(party_gstin_raw)
        party_state = xml_escape(party_state_raw)

        voucher_count += 1
        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a('    <VOUCHER REMOTEID="" VCHTYPE="Journal" ACTION="Create" OBJVIEW="Accounting Voucher View">')
        a(f"     <DATE>{dt}</DATE>")
        a("     <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>")
        a("     <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>")
        a("     <VCHENTRYMODE>Accounting Voucher View</VCHENTRYMODE>")
        a("     <ISINVOICE>No</ISINVOICE>")
        a(f"     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>")
        a("     <ISELIGIBLEFORITC>No</ISELIGIBLEFORITC>")

        if has_gst:
            a("     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>")
            a("     <GSTTRANSACTIONTYPE>Tax Invoice</GSTTRANSACTIONTYPE>")
            a("     <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>")
            if party_gstin:
                a(f'     <PARTYGSTIN>{party_gstin}</PARTYGSTIN>')
            if _cmp_gstin and _cmp_name:
                a(f'     <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="{_cmp_gstin}">{_cmp_name}</GSTREGISTRATION>')
                a(f'     <CMPGSTIN>{_cmp_gstin}</CMPGSTIN>')
                a('     <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>')
            if _cmp_state:
                a(f'     <CMPGSTSTATE>{_cmp_state}</CMPGSTSTATE>')
            if place_of_supply:
                a(f'     <PLACEOFSUPPLY>{place_of_supply}</PLACEOFSUPPLY>')
        if reference:
            a(f"     <REFERENCE>{reference}</REFERENCE>")
        if include_voucher_number and vno:
            a(f"     <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>")
        if narration:
            a(f"     <NARRATION>{narration}</NARRATION>")

        if is_sale:
            # ---- Sale Journal: Party Dr / Income Cr / Output Tax Cr ----
            a("     <LEDGERENTRIES.LIST>")
            a(f"      <LEDGERNAME>{party}</LEDGERNAME>")
            _append_common_ledger_flags(a, is_party=True, is_debit=True)
            if has_gst:
                a("      <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>")
                if party_gstin:
                    a(f"      <GSTIN>{party_gstin}</GSTIN>")
                if party_state:
                    a(f"      <STATENAME>{party_state}</STATENAME>")
                a("      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>")
            a(f"      <AMOUNT>-{fmt_amt(total)}</AMOUNT>")
            if include_bill_allocations and bill_reference:
                a("      <BILLALLOCATIONS.LIST>")
                a(f"       <NAME>{bill_reference}</NAME>")
                a("       <BILLTYPE>New Ref</BILLTYPE>")
                a(f"       <AMOUNT>-{fmt_amt(total)}</AMOUNT>")
                a("      </BILLALLOCATIONS.LIST>")
            a("     </LEDGERENTRIES.LIST>")

            a("     <LEDGERENTRIES.LIST>")
            a(f"      <LEDGERNAME>{particular}</LEDGERNAME>")
            _append_common_ledger_flags(a, is_party=False, is_debit=False)
            a(f"      <AMOUNT>{fmt_amt(taxable)}</AMOUNT>")
            a("     </LEDGERENTRIES.LIST>")

            for _ln, _la in [(cgst_ledger_raw, cgst_amt), (sgst_ledger_raw, sgst_amt), (igst_ledger_raw, igst_amt)]:
                if _la > 0 and _ln:
                    a("     <LEDGERENTRIES.LIST>")
                    a(f"      <LEDGERNAME>{xml_escape(_ln)}</LEDGERNAME>")
                    _append_common_ledger_flags(a, is_party=False, is_debit=False)
                    a(f"      <AMOUNT>{fmt_amt(_la)}</AMOUNT>")
                    a("     </LEDGERENTRIES.LIST>")

        else:
            # ---- Purchase Journal: Expense Dr / Input Tax Dr / Party Cr ----
            a("     <LEDGERENTRIES.LIST>")
            a(f"      <LEDGERNAME>{particular}</LEDGERNAME>")
            _append_common_ledger_flags(a, is_party=False)
            a(f"      <AMOUNT>-{fmt_amt(taxable)}</AMOUNT>")
            a("     </LEDGERENTRIES.LIST>")

            if cgst_amt > 0 and cgst_ledger_raw:
                a("     <LEDGERENTRIES.LIST>")
                a(f"      <LEDGERNAME>{xml_escape(cgst_ledger_raw)}</LEDGERNAME>")
                _append_common_ledger_flags(a, is_party=False)
                a(f"      <AMOUNT>-{fmt_amt(cgst_amt)}</AMOUNT>")
                a("     </LEDGERENTRIES.LIST>")

            if sgst_amt > 0 and sgst_ledger_raw:
                a("     <LEDGERENTRIES.LIST>")
                a(f"      <LEDGERNAME>{xml_escape(sgst_ledger_raw)}</LEDGERNAME>")
                _append_common_ledger_flags(a, is_party=False)
                a(f"      <AMOUNT>-{fmt_amt(sgst_amt)}</AMOUNT>")
                a("     </LEDGERENTRIES.LIST>")

            if igst_amt > 0 and igst_ledger_raw:
                a("     <LEDGERENTRIES.LIST>")
                a(f"      <LEDGERNAME>{xml_escape(igst_ledger_raw)}</LEDGERNAME>")
                _append_common_ledger_flags(a, is_party=False)
                a(f"      <AMOUNT>-{fmt_amt(igst_amt)}</AMOUNT>")
                a("     </LEDGERENTRIES.LIST>")

            # TDS Payable - Credit (positive amount = Credit in Journal Accounting Voucher View)
            if jnl_tds_led and jnl_tds_amount > 0:
                a("     <LEDGERENTRIES.LIST>")
                a(f"      <LEDGERNAME>{jnl_tds_led}</LEDGERNAME>")
                _append_common_ledger_flags(a, is_party=False, is_debit=False)
                a(f"      <AMOUNT>{fmt_amt(jnl_tds_amount)}</AMOUNT>")
                a("     </LEDGERENTRIES.LIST>")

            a("     <LEDGERENTRIES.LIST>")
            a(f"      <LEDGERNAME>{party}</LEDGERNAME>")
            _append_common_ledger_flags(a, is_party=True)
            if has_gst:
                a("      <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>")
                if party_gstin:
                    a(f"      <GSTIN>{party_gstin}</GSTIN>")
                if party_state:
                    a(f"      <STATENAME>{party_state}</STATENAME>")
                a("      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>")
            a(f"      <AMOUNT>{fmt_amt(jnl_party_total)}</AMOUNT>")
            if include_bill_allocations and bill_reference:
                a("      <BILLALLOCATIONS.LIST>")
                a(f"       <NAME>{bill_reference}</NAME>")
                a("       <BILLTYPE>New Ref</BILLTYPE>")
                a(f"       <AMOUNT>{fmt_amt(jnl_party_total)}</AMOUNT>")
                a("      </BILLALLOCATIONS.LIST>")
            a("     </LEDGERENTRIES.LIST>")

        a("    </VOUCHER>")
        a("   </TALLYMESSAGE>")

    a("  </REQUESTDATA>")
    a(" </IMPORTDATA></BODY>")
    a("</ENVELOPE>")
    return "\n".join(lines), voucher_count


# ═══════════════════════════════════════════════════════════════════════════
#  CREDIT / DEBIT NOTE XML GENERATION
# ═══════════════════════════════════════════════════════════════════════════

def _normalize_note_type(value: str) -> str:
    text = str(value or "").strip().casefold()
    if text in {"debit note", "debit", "debitnote"}:
        return "Debit Note"
    return "Credit Note"


def _clean_note_tax_ledger(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if text.casefold() in {"0", "0.0", "none", "na", "n/a", "-"}:
        return ""
    return text


def generate_note_xml(
    rows: list,
    company: str,
    use_today_date: bool = False,
    date_mode: str = "",
    custom_tally_date: str = "",
    voucher_type: str = "Credit Note",
    company_gst_registrations: list = None,
) -> tuple:
    """
    Credit/Debit Note accounting:
    - Credit Note: party credited (ISDEEMEDPOSITIVE=No), particular/tax debited (Yes)
    - Debit Note:  party debited  (ISDEEMEDPOSITIVE=Yes), particular/tax credited (No)
    """
    normalized_type = _normalize_note_type(voucher_type)
    is_debit_note = (normalized_type == "Debit Note")
    default_particular_ledger = f"{normalized_type} Account"

    lines_out = []
    a = lines_out.append
    company_static = _company_static_block(company)

    a('<?xml version="1.0" encoding="UTF-8"?>')
    a("<ENVELOPE>")
    a(" <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>")
    a(" <BODY><IMPORTDATA>")
    a("  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>")
    if company_static:
        a(company_static)
    a("  </REQUESTDESC>")
    a("  <REQUESTDATA>")

    voucher_count = 0
    resolved_mode = str(date_mode or ("current" if use_today_date else "excel")).strip().lower()
    if resolved_mode not in {"current", "excel", "custom"}:
        resolved_mode = "current" if use_today_date else "excel"
    resolved_custom_date = _normalize_manual_date_to_tally(custom_tally_date) if resolved_mode == "custom" else ""

    for idx, r in enumerate(rows):
        taxable = _row_float(r, "TaxableValue", 0.0)
        if taxable <= 0:
            continue

        if resolved_mode == "current":
            from datetime import datetime as _dt
            source_date = _dt.today()
        elif resolved_mode == "custom":
            source_date = resolved_custom_date
        else:
            source_date = _row_get(r, "Date", "")
        dt = tally_date(source_date)

        vno_raw = _row_voucher_number(r, str(idx + 1))
        party_raw = _ledger_or_suspense(_row_text(r, "PartyLedger") or _row_text(r, "Party Ledger"))
        particular_raw = (
            _row_text(r, "Particular") or _row_text(r, "Particulars")
            or _row_text(r, "SalesLedger") or _row_text(r, "Sales Ledger")
            or _row_text(r, "Purchase Ledger") or default_particular_ledger
        )
        particular_raw = _ledger_or_suspense(particular_raw) or default_particular_ledger

        cgst_ledger_raw = _clean_note_tax_ledger(_row_text(r, "CGSTLedger"))
        sgst_ledger_raw = _clean_note_tax_ledger(_row_text(r, "SGSTLedger"))
        igst_ledger_raw = _clean_note_tax_ledger(_row_text(r, "IGSTLedger"))

        cgst_rate = _row_float(r, "CGSTRate", 0.0)
        sgst_rate = _row_float(r, "SGSTRate", 0.0)
        igst_rate = _row_float(r, "IGSTRate", 0.0)

        cgst_amt = round(taxable * cgst_rate / 100, 2) if cgst_rate > 0 else 0.0
        sgst_amt = round(taxable * sgst_rate / 100, 2) if sgst_rate > 0 else 0.0
        igst_amt = round(taxable * igst_rate / 100, 2) if igst_rate > 0 else 0.0
        total = taxable + cgst_amt + sgst_amt + igst_amt

        vno = xml_escape(vno_raw)
        party = xml_escape(party_raw)
        particular = xml_escape(particular_raw)
        narration = xml_escape(_row_text(r, "Narration"))
        gstin_raw = _row_text(r, "GSTIN") or _row_text(r, "PartyGSTIN")
        gstin = xml_escape(gstin_raw)

        state_name_raw = _state_name_from_gstin(gstin_raw)
        state_xml = xml_escape(state_name_raw)

        # Sign convention
        party_is_deemed_positive = "Yes" if is_debit_note else "No"
        party_amount = -total if is_debit_note else total
        counter_is_deemed_positive = "No" if is_debit_note else "Yes"
        taxable_amount = taxable if is_debit_note else -taxable
        cgst_amount = cgst_amt if is_debit_note else -cgst_amt
        sgst_amount = sgst_amt if is_debit_note else -sgst_amt
        igst_amount = igst_amt if is_debit_note else -igst_amt

        voucher_count += 1
        a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        a(f'    <VOUCHER VCHTYPE="{normalized_type}" ACTION="Create" OBJVIEW="Invoice Voucher View">')
        a(f"     <DATE>{dt}</DATE>")
        a(f"     <VOUCHERTYPENAME>{normalized_type}</VOUCHERTYPENAME>")
        a(f"     <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>")
        a(f"     <PARTYLEDGERNAME>{party}</PARTYLEDGERNAME>")
        a(f"     <PARTYNAME>{party}</PARTYNAME>")

        if not is_debit_note:
            a(f"     <BASICBUYERNAME>{party}</BASICBUYERNAME>")
            if state_xml:
                a(f"     <STATENAME>{state_xml}</STATENAME>")
                a(f"     <PLACEOFSUPPLY>{state_xml}</PLACEOFSUPPLY>")
        else:
            _cmp_regs = list(company_gst_registrations or [])
            _cmp_state = _cmp_gstin = _cmp_name = ""
            if _cmp_regs:
                _cr = _cmp_regs[0]
                _cmp_gstin = xml_escape(str(_cr.get("gstin", "") or "").strip())
                _cmp_state = xml_escape(str(_cr.get("state", "") or "").strip())
                _cmp_name  = xml_escape(str(_cr.get("name", "") or "").strip())
            if state_xml:
                a(f"     <STATENAME>{state_xml}</STATENAME>")
            if _cmp_state:
                a(f"     <PLACEOFSUPPLY>{_cmp_state}</PLACEOFSUPPLY>")
            _dn_reg = "Regular" if gstin_raw else "Unregistered"
            a(f"     <GSTREGISTRATIONTYPE>{_dn_reg}</GSTREGISTRATIONTYPE>")
            if gstin_raw:
                a("     <VATDEALERTYPE>Regular</VATDEALERTYPE>")
            if _cmp_gstin and _cmp_name:
                a(f'     <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="{_cmp_gstin}">{_cmp_name}</GSTREGISTRATION>')
                a(f'     <CMPGSTIN>{_cmp_gstin}</CMPGSTIN>')
                a('     <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>')
            if _cmp_state:
                a(f'     <CMPGSTSTATE>{_cmp_state}</CMPGSTSTATE>')

        a("     <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>")
        a(f"     <EFFECTIVEDATE>{dt}</EFFECTIVEDATE>")
        a("     <ISINVOICE>Yes</ISINVOICE>")
        a("     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>")
        a("     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>")
        if is_debit_note:
            a("     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>")
            _dn_gst_txn = "Tax Invoice" if gstin_raw else "Unregistered"
            a(f"     <GSTTRANSACTIONTYPE>{_dn_gst_txn}</GSTTRANSACTIONTYPE>")
        if gstin:
            a(f"     <PARTYGSTIN>{gstin}</PARTYGSTIN>")
        if narration:
            a(f"     <NARRATION>{narration}</NARRATION>")

        # Party ledger entry
        a("     <LEDGERENTRIES.LIST>")
        a(f"      <LEDGERNAME>{party}</LEDGERNAME>")
        a(f"      <ISDEEMEDPOSITIVE>{party_is_deemed_positive}</ISDEEMEDPOSITIVE>")
        a(f"      <AMOUNT>{fmt_amt(party_amount)}</AMOUNT>")
        a("     </LEDGERENTRIES.LIST>")

        # Particular / income ledger
        a("     <LEDGERENTRIES.LIST>")
        a(f"      <LEDGERNAME>{particular}</LEDGERNAME>")
        a(f"      <ISDEEMEDPOSITIVE>{counter_is_deemed_positive}</ISDEEMEDPOSITIVE>")
        a(f"      <AMOUNT>{fmt_amt(taxable_amount)}</AMOUNT>")
        a("     </LEDGERENTRIES.LIST>")

        if cgst_amt > 0 and cgst_ledger_raw:
            a("     <LEDGERENTRIES.LIST>")
            a(f"      <LEDGERNAME>{xml_escape(cgst_ledger_raw)}</LEDGERNAME>")
            a(f"      <ISDEEMEDPOSITIVE>{counter_is_deemed_positive}</ISDEEMEDPOSITIVE>")
            a(f"      <AMOUNT>{fmt_amt(cgst_amount)}</AMOUNT>")
            a("     </LEDGERENTRIES.LIST>")

        if sgst_amt > 0 and sgst_ledger_raw:
            a("     <LEDGERENTRIES.LIST>")
            a(f"      <LEDGERNAME>{xml_escape(sgst_ledger_raw)}</LEDGERNAME>")
            a(f"      <ISDEEMEDPOSITIVE>{counter_is_deemed_positive}</ISDEEMEDPOSITIVE>")
            a(f"      <AMOUNT>{fmt_amt(sgst_amount)}</AMOUNT>")
            a("     </LEDGERENTRIES.LIST>")

        if igst_amt > 0 and igst_ledger_raw:
            a("     <LEDGERENTRIES.LIST>")
            a(f"      <LEDGERNAME>{xml_escape(igst_ledger_raw)}</LEDGERNAME>")
            a(f"      <ISDEEMEDPOSITIVE>{counter_is_deemed_positive}</ISDEEMEDPOSITIVE>")
            a(f"      <AMOUNT>{fmt_amt(igst_amount)}</AMOUNT>")
            a("     </LEDGERENTRIES.LIST>")

        a("    </VOUCHER>")
        a("   </TALLYMESSAGE>")

    a("  </REQUESTDATA>")
    a(" </IMPORTDATA></BODY>")
    a("</ENVELOPE>")
    return "\n".join(lines_out), voucher_count


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
        self.voucher_date_mode_var = ctk.StringVar(value="excel")
        self.voucher_custom_date_var = ctk.StringVar(value="")
        self.voucher_date_checks = {
            "current": ctk.BooleanVar(value=False),
            "excel": ctk.BooleanVar(value=True),
            "custom": ctk.BooleanVar(value=False),
        }
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
        self._voucher_xml_fp_vars = {}
        self._voucher_xml_browse_buttons = {}
        self._voucher_import_buttons = {}  # kept for compat, not used in UI
        self._active_preview_mode = {}     # tracks "excel" or "xml" per mode
        self._note_type_var = ctk.StringVar(value="Credit Note")
        self.company_gst_registrations = []
        self._note_loaded_rows = []
        self._note_loaded_headers = []
        self.fetched_party_ledgers = []
        self._ledger_fetch_running = False
        # Journal Entry instance variables
        self._jnl_loaded_rows = []
        self._jnl_loaded_headers = []
        self._jnl_manual_rows = []
        self._jnl_manual_form_vars = {}
        self._jnl_manual_party_ledger_combo = None
        self._jnl_manual_party_search_var = ctk.StringVar(value="")
        self._jnl_manual_party_match_label = None
        self._jnl_manual_fetch_ledger_btn = None
        self._jnl_manual_tree = None
        self._jnl_excel_tree = None
        self._jnl_manual_editing_index = None
        self._jnl_manual_update_btn = None
        self._jnl_source_tabs = None
        self._jnl_manual_info_label = None
        self._jnl_excel_info_label = None
        self._journal_type_var = ctk.StringVar(value="purchase")
        self._jnl_manual_party_search_clear_btn = None
        # Note panel manual entry vars
        self._note_manual_rows = []
        self._note_manual_form_vars = {}
        self._note_manual_party_ledger_combo = None
        self._note_manual_party_search_var = ctk.StringVar(value="")
        self._note_manual_party_match_label = None
        self._note_manual_fetch_ledger_btn = None
        self._note_manual_tree = None
        self._note_excel_tree = None
        self._note_manual_editing_index = None
        self._note_manual_update_btn = None
        self._note_source_tabs = None
        self._note_manual_info_label = None
        self._note_excel_info_label = None
        self._note_manual_party_search_clear_btn = None
        self._push_running = False
        self._push_overlay = None
        self._push_message_var = ctk.StringVar(value="")
        self.workflow_demo_url = ""
        self.demo_btn = None
        self.voucher_date_current_cb = None
        self.voucher_date_excel_cb = None
        self.voucher_date_custom_cb = None
        self.voucher_custom_date_entry = None
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

        self.demo_btn = ctk.CTkButton(
            row_1,
            text="▶ View Demo",
            width=132,
            height=32,
            font=("Segoe UI", 10, "bold"),
            fg_color="#DC2626",
            hover_color="#B91C1C",
            text_color="#FFFFFF",
            corner_radius=8,
            command=self._view_workflow_demo,
        )
        self.demo_btn.pack(side="right", padx=(0, 8))

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
        self.connection_status_label.pack(side="left", padx=(0, 15))
        self.company_status_label = ctk.CTkLabel(
            status_row,
            textvariable=self.company_status_var,
            font=("Segoe UI", 10),
            text_color=COLORS["text_muted"],
        )
        self.company_status_label.pack(side="left", padx=(0, 15))

        date_mode_row = ctk.CTkFrame(settings_card, fg_color="transparent")
        date_mode_row.pack(fill="x", padx=14, pady=(0, 10))
        ctk.CTkLabel(
            date_mode_row,
            text="Voucher Date",
            font=("Segoe UI", 10),
            text_color=COLORS["text_secondary"],
        ).pack(side="left")

        checks_wrap = ctk.CTkFrame(date_mode_row, fg_color="transparent")
        checks_wrap.pack(side="left", padx=(8, 0))

        self.voucher_date_current_cb = ctk.CTkCheckBox(
            checks_wrap,
            text="Current Date",
            variable=self.voucher_date_checks["current"],
            font=("Segoe UI", 10),
            text_color=COLORS["text_secondary"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            border_color=COLORS["border"],
            command=lambda: self._set_voucher_date_mode("current"),
        )
        self.voucher_date_current_cb.pack(side="left", padx=(0, 8))

        self.voucher_date_excel_cb = ctk.CTkCheckBox(
            checks_wrap,
            text="Excel Date",
            variable=self.voucher_date_checks["excel"],
            font=("Segoe UI", 10),
            text_color=COLORS["text_secondary"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            border_color=COLORS["border"],
            command=lambda: self._set_voucher_date_mode("excel"),
        )
        self.voucher_date_excel_cb.pack(side="left", padx=(0, 8))

        self.voucher_date_custom_cb = ctk.CTkCheckBox(
            checks_wrap,
            text="Custom Date",
            variable=self.voucher_date_checks["custom"],
            font=("Segoe UI", 10),
            text_color=COLORS["text_secondary"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            border_color=COLORS["border"],
            command=lambda: self._set_voucher_date_mode("custom"),
        )
        self.voucher_date_custom_cb.pack(side="left")

        self.voucher_custom_date_entry = ctk.CTkEntry(
            checks_wrap,
            textvariable=self.voucher_custom_date_var,
            width=170,
            height=30,
            fg_color=COLORS["bg_input"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
            placeholder_text="DD/MM/YYYY",
            font=("Segoe UI", 10),
        )
        self.voucher_custom_date_entry.pack(side="left", padx=(8, 0))
        self._set_voucher_date_mode("excel")

        # --- Shifted Select Mode into date_mode_row ---
        ctk.CTkLabel(date_mode_row, text="Select Mode", font=("Segoe UI", 10, "bold"),
                     text_color=COLORS["text_secondary"]).pack(side="left", padx=(40, 10))

        _panel_options = [
            "📋  Sales Accounting Invoice",
            "📦  Sales Item Invoice",
            "🧾  Purchase Accounting Invoice",
            "🛒  Purchase Item Invoice",
            "📝  Credit Note",
            "📒  Debit Note",
            "📓  Journal Entry",
            "🏦  Create Ledgers",
            "📁  Create Stock Items",
        ]
        self._panel_option_to_key = {
            "📋  Sales Accounting Invoice": "accounting",
            "📦  Sales Item Invoice": "item",
            "🧾  Purchase Accounting Invoice": "purchase_accounting",
            "🛒  Purchase Item Invoice": "purchase_item",
            "📝  Credit Note": "credit_note",
            "📒  Debit Note": "debit_note",
            "📓  Journal Entry": "journal",
            "🏦  Create Ledgers": "ledger",
            "📁  Create Stock Items": "stock",
        }
        self._panel_var = ctk.StringVar(value=_panel_options[0])
        self._panel_option_menu = ctk.CTkOptionMenu(
            date_mode_row,
            variable=self._panel_var,
            values=_panel_options,
            width=300,
            height=34,
            font=("Segoe UI", 10, "bold"),
            fg_color=COLORS["accent"],
            button_color=COLORS["accent_hover"],
            button_hover_color=COLORS["accent_hover"],
            text_color="#FFFFFF",
            dropdown_fg_color=COLORS["bg_card"],
            dropdown_text_color=COLORS["text_primary"],
            dropdown_hover_color=COLORS["bg_input"],
            corner_radius=8,
            command=self._switch_panel,
        )
        self._panel_option_menu.pack(side="left", padx=(2, 0))



        # ─── Content container (all panels stacked in same grid cell) ────────────
        _content_outer = ctk.CTkFrame(self, fg_color=COLORS["bg_card"], corner_radius=10,
                                       border_width=1, border_color=COLORS["border"])
        _content_outer.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        _content_outer.grid_rowconfigure(0, weight=1)
        _content_outer.grid_columnconfigure(0, weight=1)

        self._panels = {}
        for _key in ("accounting", "item", "purchase_accounting", "purchase_item", "ledger", "stock"):
            _pf = ctk.CTkFrame(_content_outer, fg_color="transparent")
            _pf.grid(row=0, column=0, sticky="nsew")
            self._panels[_key] = _pf

        # Shared Credit/Debit Note panel
        _note_pf = ctk.CTkFrame(_content_outer, fg_color="transparent")
        _note_pf.grid(row=0, column=0, sticky="nsew")
        self._panels["credit_note"] = _note_pf
        self._panels["debit_note"] = _note_pf   # same frame — note_type_var controls type
        self._panels["note"] = _note_pf

        # Journal Entry panel
        journal_pf = ctk.CTkFrame(_content_outer, fg_color="transparent")
        journal_pf.grid(row=0, column=0, sticky="nsew")
        self._panels["journal"] = journal_pf

        self.tab_acct = self._panels["accounting"]
        self.tab_item = self._panels["item"]
        self.tab_purchase_acct = self._panels["purchase_accounting"]
        self.tab_purchase_item = self._panels["purchase_item"]
        self.tab_ledger = self._panels["ledger"]
        self.tab_stock = self._panels["stock"]

        self._build_voucher_tab(self.tab_acct, mode="accounting")
        self._build_voucher_tab(self.tab_item, mode="item")
        self._build_voucher_tab(self.tab_purchase_acct, mode="purchase_accounting")
        self._build_voucher_tab(self.tab_purchase_item, mode="purchase_item")
        self._build_note_panel(self._panels["note"])
        self._build_journal_panel(self._panels["journal"])
        self._build_ledger_tab()
        self._build_stock_tab()

        # Show first panel
        self._switch_panel(_panel_options[0])

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

    def _set_voucher_date_mode(self, selected_mode: str):
        mode = str(selected_mode or "excel").strip().lower()
        if mode not in {"current", "excel", "custom"}:
            mode = "excel"

        self.voucher_date_mode_var.set(mode)
        for key, var in self.voucher_date_checks.items():
            var.set(key == mode)

        if self.voucher_custom_date_entry is not None:
            self.voucher_custom_date_entry.configure(
                state="normal" if (mode == "custom" and not self._push_running) else "disabled"
            )

    def _get_voucher_date_selection(self):
        mode = str(self.voucher_date_mode_var.get() or "excel").strip().lower()
        if mode not in {"current", "excel", "custom"}:
            mode = "excel"
            self._set_voucher_date_mode(mode)

        custom_tally_date = ""
        if mode == "custom":
            custom_raw = (self.voucher_custom_date_var.get() or "").strip()
            if not custom_raw:
                raise ValueError("Enter custom date or select Current Date / Excel Date.")
            custom_tally_date = _normalize_manual_date_to_tally(custom_raw)

        return mode, custom_tally_date

    def _view_workflow_demo(self):
        # 1. Map current panel selection to specific YouTube links
        current_selection = self._panel_var.get()
        current_key = self._panel_option_to_key.get(current_selection, "")

        demo_map = {
            "accounting": "https://youtu.be/SlVhqglVSzU",           # Sales Accounting Invoice
            "purchase_accounting": "https://youtu.be/9FSOjQoHmk8",  # Purchase Accounting Invoice
            "purchase_item": "https://youtu.be/DbXzZsqb9q8",        # Purchase Item Invoice
        }

        # 2. Determine the URL (Specific -> Instance Var -> Global Default)
        demo_url = demo_map.get(current_key)
        if not demo_url:
            demo_url = (self.workflow_demo_url or "https://www.youtube.com/watch?v=OEJ7H5bJNcM").strip()

        # 3. Open the browser
        if demo_url:
            try:
                # Using open_new_tab for consistent behavior
                webbrowser.open_new_tab(demo_url)
                return
            except Exception as exc:
                messagebox.showwarning("View Demo", f"Could not open demo link.\n\n{exc}")
                return

        messagebox.showinfo(
            "View Demo",
            "Demo link is not set yet.\n\nSet self.workflow_demo_url in code to your YouTube link later.",
        )

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
        parent.grid_rowconfigure(4, weight=1)  # row 4 = preview container

        # Row 0: Template download
        template_row = ctk.CTkFrame(parent, fg_color="transparent")
        template_row.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        template_btn = ctk.CTkButton(
            template_row, text="📥  Download Template",
            fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"], width=170,
            command=lambda: self._download_template_for_mode(mode),
        )
        template_btn.pack(side="right")
        self._voucher_template_buttons[mode] = template_btn

        # Row 1: Browse Excel row
        load_frame = ctk.CTkFrame(parent, fg_color="transparent")
        load_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 0))
        fp_var = ctk.StringVar()
        ctk.CTkEntry(load_frame, textvariable=fp_var,
                     placeholder_text="Select Excel file (.xlsx / .xlsm)...",
                     width=500, state="readonly").pack(side="left", padx=(0, 8))

        self._active_preview_mode[mode] = "excel"  # default

        def browse():
            if self._voucher_load_running.get(mode):
                self.status_var.set("Please wait, file is still loading...")
                return
            f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
            if f:
                fp_var.set(f)
                self._load_preview(f, tree, mode)
                self._show_preview_mode(excel_tree_frame, xml_preview_frame, "excel",
                                        excel_toggle_btn, xml_toggle_btn, mode_key=mode)

        browse_btn = ctk.CTkButton(load_frame, text="Browse Excel", command=browse,
                                    width=90, fg_color=ACCENT, hover_color=ACCENT_HOVER)
        browse_btn.pack(side="left", padx=(0, 8))
        self._voucher_browse_buttons[mode] = browse_btn
        self._voucher_load_running[mode] = False

        if mode in {"item", "purchase_item"}:
            ctk.CTkLabel(load_frame,
                         text="⚠ Requires: ItemName, Quantity, Rate, and Per or Unit/UOM columns",
                         font=("Segoe UI", 11), text_color=COLORS["warning"]).pack(side="left")

        # Row 2: Browse XML row
        xml_row = ctk.CTkFrame(parent, fg_color="transparent")
        xml_row.grid(row=2, column=0, sticky="ew", padx=10, pady=(6, 0))
        xml_fp_var = ctk.StringVar()
        self._voucher_xml_fp_vars[mode] = xml_fp_var
        ctk.CTkEntry(xml_row, textvariable=xml_fp_var,
                     placeholder_text="Or select existing XML file to import/preview...",
                     width=500, state="readonly").pack(side="left", padx=(0, 8))

        def browse_xml():
            f = filedialog.askopenfilename(filetypes=[("XML Files", "*.xml")])
            if f:
                xml_fp_var.set(f)
                self._load_xml_preview(f, xml_text, mode, info_lbl)
                self._show_preview_mode(excel_tree_frame, xml_preview_frame, "xml",
                                        excel_toggle_btn, xml_toggle_btn, mode_key=mode)

        xml_browse_btn = ctk.CTkButton(
            xml_row, text="Browse XML", command=browse_xml, width=90,
            fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"])
        xml_browse_btn.pack(side="left", padx=(0, 8))
        self._voucher_xml_browse_buttons[mode] = xml_browse_btn
        ctk.CTkLabel(xml_row, text="Browse an existing XML to preview & push it directly to Tally",
                     font=("Segoe UI", 10), text_color=COLORS["text_muted"]).pack(side="left")

        # Row 3: Preview toggle + info label
        toggle_row = ctk.CTkFrame(parent, fg_color="transparent")
        toggle_row.grid(row=3, column=0, sticky="ew", padx=10, pady=(8, 0))

        excel_toggle_btn = ctk.CTkButton(
            toggle_row, text="📊 Excel Data", width=160, height=30,
            font=("Segoe UI", 10, "bold"), fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"], text_color="#FFFFFF", corner_radius=6)
        excel_toggle_btn.pack(side="left", padx=(0, 4))

        xml_toggle_btn = ctk.CTkButton(
            toggle_row, text="📄 XML Preview", width=150, height=30,
            font=("Segoe UI", 10, "bold"), fg_color=COLORS["bg_input"],
            hover_color=COLORS["bg_card_hover"], text_color=COLORS["text_muted"], corner_radius=6)
        xml_toggle_btn.pack(side="left", padx=(0, 8))

        info_lbl = ctk.CTkLabel(toggle_row, text="", font=("Segoe UI", 11),
                                 text_color=TEXT_MUTED)
        info_lbl.pack(side="left", padx=4)
        self._voucher_info_labels[mode] = info_lbl

        # Row 4: Preview container — Excel treeview + XML text widget stacked
        preview_container = ctk.CTkFrame(parent, fg_color="transparent")
        preview_container.grid(row=4, column=0, sticky="nsew", padx=10, pady=(4, 4))
        preview_container.grid_rowconfigure(0, weight=1)
        preview_container.grid_columnconfigure(0, weight=1)

        # Excel treeview frame
        excel_tree_frame = ctk.CTkFrame(
            preview_container, fg_color=COLORS["bg_dark"], corner_radius=8,
            border_width=1, border_color=COLORS["border"])
        excel_tree_frame.grid(row=0, column=0, sticky="nsew")
        excel_tree_frame.grid_rowconfigure(0, weight=1)
        excel_tree_frame.grid_columnconfigure(0, weight=1)

        tree_scroll_y = ttk.Scrollbar(excel_tree_frame, orient="vertical")
        tree_scroll_x = ttk.Scrollbar(excel_tree_frame, orient="horizontal")
        tree = ttk.Treeview(excel_tree_frame, show="headings",
                            yscrollcommand=tree_scroll_y.set,
                            xscrollcommand=tree_scroll_x.set)
        tree_scroll_y.config(command=tree.yview)
        tree_scroll_x.config(command=tree.xview)
        tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")

        # XML preview frame
        xml_preview_frame = ctk.CTkFrame(
            preview_container, fg_color=COLORS["bg_dark"], corner_radius=8,
            border_width=1, border_color=COLORS["border"])
        xml_preview_frame.grid(row=0, column=0, sticky="nsew")
        xml_preview_frame.grid_rowconfigure(0, weight=1)
        xml_preview_frame.grid_columnconfigure(0, weight=1)

        _xml_bg = self._resolve_theme_color("bg_card")
        _xml_fg = self._resolve_theme_color("text_primary")
        xml_text = tk.Text(
            xml_preview_frame, wrap="none", font=("Consolas", 10),
            bg=_xml_bg, fg=_xml_fg, insertbackground=_xml_fg,
            relief="flat", borderwidth=0)
        _xsy = ttk.Scrollbar(xml_preview_frame, orient="vertical", command=xml_text.yview)
        _xsx = ttk.Scrollbar(xml_preview_frame, orient="horizontal", command=xml_text.xview)
        xml_text.configure(yscrollcommand=_xsy.set, xscrollcommand=_xsx.set)
        xml_text.grid(row=0, column=0, sticky="nsew")
        _xsy.grid(row=0, column=1, sticky="ns")
        _xsx.grid(row=1, column=0, sticky="ew")
        xml_text.insert("end", "Browse an XML file above to preview its contents here.")
        xml_text.configure(state="disabled")

        # Wire toggle buttons
        excel_toggle_btn.configure(command=lambda: self._show_preview_mode(
            excel_tree_frame, xml_preview_frame, "excel", excel_toggle_btn, xml_toggle_btn, mode_key=mode))
        xml_toggle_btn.configure(command=lambda: self._show_preview_mode(
            excel_tree_frame, xml_preview_frame, "xml", excel_toggle_btn, xml_toggle_btn, mode_key=mode))

        # Show Excel frame by default
        excel_tree_frame.tkraise()

        # Row 5: Action buttons (smart Push: uses active preview mode)
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=5, column=0, sticky="ew", padx=10, pady=(4, 10))

        save_btn = ctk.CTkButton(
            btn_frame, text="💾  Save XML File",
            fg_color=SUCCESS, hover_color="#15803D", width=155,
            command=lambda: self._generate(mode, "save", fp_var.get()))
        save_btn.pack(side="left", padx=(0, 8))
        self._voucher_save_buttons[mode] = save_btn

        push_btn = ctk.CTkButton(
            btn_frame, text="🚀  Push to Tally",
            fg_color=ACCENT, hover_color=ACCENT_HOVER, width=165,
            command=lambda: self._smart_push(mode, fp_var, xml_fp_var))
        push_btn.pack(side="left", padx=(0, 8))
        self._voucher_push_buttons[mode] = push_btn

        hint_lbl = ctk.CTkLabel(
            btn_frame, text="Push uses active preview (Excel or XML)",
            font=("Segoe UI", 10), text_color=COLORS["text_muted"])
        hint_lbl.pack(side="left", padx=8)


    def _set_voucher_loading_state(self, mode: str, is_loading: bool):
        self._voucher_load_running[mode] = is_loading
        state = "disabled" if is_loading else "normal"
        browse_text = "Loading..." if is_loading else "Browse Excel"
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
        xml_browse_btn = self._voucher_xml_browse_buttons.get(mode)
        if xml_browse_btn:
            xml_browse_btn.configure(state=state)
        import_btn = self._voucher_import_buttons.get(mode)
        if import_btn:
            import_btn.configure(state=state)

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
        for btn in self._voucher_xml_browse_buttons.values():
            btn.configure(state=state)
        if self.demo_btn is not None:
            self.demo_btn.configure(state=state)

        if self.voucher_date_current_cb is not None:
            self.voucher_date_current_cb.configure(state=state)
        if self.voucher_date_excel_cb is not None:
            self.voucher_date_excel_cb.configure(state=state)
        if self.voucher_date_custom_cb is not None:
            self.voucher_date_custom_cb.configure(state=state)

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

        self._set_voucher_date_mode(self.voucher_date_mode_var.get())

        self.update_idletasks()

    def _get_template_definition(self, mode: str) -> dict:
        templates = {
            "accounting": {
                "sheet_name": "Sheet1",
                "headers": [
                    "Date",
                    "VoucherNo",
                    "InvoiceNo",
                    "PartyLedger",
                    "PartyName",
                    "PartyMailingName",
                    "PartyAddress1",
                    "PartyAddress2",
                    "PartyPincode",
                    "PartyState",
                    "PlaceOfSupply",
                    "PartyCountry",
                    "GSTApplicable",
                    "GSTRegistrationType",
                    "GSTIN/UIN",
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
                "sample_rows": [
                    ["01/04/2024", "1", "INV-001", "ABC Traders", "ABC Traders", "ABC Traders",
                     "123 Main Street", "City Centre", "400001", "Maharashtra", "Maharashtra", "India",
                     "Applicable", "Regular", "27AABCU9603R1ZM", "Sales Account", 10000,
                     "CGST", 9, "SGST", 9, "IGST", 0, "Being goods sold to ABC Traders"],
                    ["02/04/2024", "2", "INV-002", "XYZ Enterprises", "XYZ Enterprises", "XYZ Enterprises",
                     "456 Industrial Area", "", "110001", "Delhi", "Delhi", "India",
                     "Applicable", "Regular", "07AAGFX1234A1Z5", "Sales Account", 20000,
                     "CGST", 0, "SGST", 0, "IGST", 18, "Being goods sold to XYZ Enterprises"],
                ],
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
                    "InvoiceNo",
                    "SupplierInvoiceNo",
                    "PartyLedger",
                    "PartyName",
                    "PartyMailingName",
                    "PartyAddress1",
                    "PartyAddress2",
                    "PartyPincode",
                    "PartyState",
                    "PlaceOfSupply",
                    "PartyCountry",
                    "GSTApplicable",
                    "GSTRegistrationType",
                    "GSTIN/UIN",
                    "PurchaseLedger",
                    "TaxableValue",
                    "CGSTLedger",
                    "CGSTRate",
                    "SGSTLedger",
                    "SGSTRate",
                    "IGSTLedger",
                    "IGSTRate",
                    "Narration",
                    "TDSLedger",
                    "TDSRate",
                    "TDSAmount",
                ],
                "sample_rows": [
                    ["01/04/2024", "1", "SINV-001", "SINV-001", "PQR Suppliers", "PQR Suppliers", "PQR Suppliers",
                     "789 MIDC Road", "", "411001", "Maharashtra", "Maharashtra", "India",
                     "Applicable", "Regular", "27AAAPQ1234B1Z3", "Purchase Account", 10000,
                     "CGST", 9, "SGST", 9, "IGST", 0, "Being goods purchased from PQR Suppliers",
                     "TDS Payable on Professional", 10, ""],
                    ["02/04/2024", "2", "SINV-002", "SINV-002", "LMN Industries", "LMN Industries", "LMN Industries",
                     "321 Ring Road", "", "302001", "Rajasthan", "Maharashtra", "India",
                     "Applicable", "Regular", "08AAELM9876C1Z1", "Purchase Account", 20000,
                     "CGST", 0, "SGST", 0, "IGST", 18, "Being goods purchased from LMN Industries",
                     "", "", ""],
                ],
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
        date_mode, custom_tally_date = self._get_voucher_date_selection()
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

        def build_voucher_xml(
            selected_mode: str,
            selected_date_mode: str,
            selected_custom_date: str,
            voucher_start,
            rows_data,
            company_gst_registrations=None,
            legacy_item_invoice_context: bool = False,
        ):
            if selected_mode == "accounting":
                return generate_accounting_xml(
                    rows_data,
                    company,
                    date_mode=selected_date_mode,
                    custom_tally_date=selected_custom_date,
                    start_voucher_number=voucher_start,
                    company_gst_registrations=company_gst_registrations,
                )
            if selected_mode == "item":
                return generate_item_xml(
                    rows_data,
                    company,
                    date_mode=selected_date_mode,
                    custom_tally_date=selected_custom_date,
                    start_voucher_number=voucher_start,
                    company_gst_registrations=company_gst_registrations,
                    legacy_invoice_context=legacy_item_invoice_context,
                )
            if selected_mode == "purchase_accounting":
                return generate_purchase_accounting_xml(
                    rows_data,
                    company,
                    date_mode=selected_date_mode,
                    custom_tally_date=selected_custom_date,
                    start_voucher_number=voucher_start,
                    company_gst_registrations=company_gst_registrations,
                )
            if selected_mode == "purchase_item":
                return generate_purchase_item_xml(
                    rows_data,
                    company,
                    date_mode=selected_date_mode,
                    custom_tally_date=selected_custom_date,
                    start_voucher_number=voucher_start,
                    company_gst_registrations=company_gst_registrations,
                )
            raise ValueError(f"Unsupported mode: {selected_mode}")

        try:
            if action == "save":
                xml = build_voucher_xml(mode, date_mode, custom_tally_date, None, rows_to_use)
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
                date_mode_label = {
                    "current": "current date",
                    "excel": "excel date",
                    "custom": "custom date",
                }.get(date_mode, "excel date")
                voucher_type = mode_to_voucher_type.get(mode, "Sales")
                rows_snapshot = list(rows_to_use)

                self._set_push_loading_state(True, f"Preparing vouchers for {target_company}...")
                self.status_var.set(f"Posting to Tally ({target_company}, {date_mode_label})...")

                def worker():
                    result = {
                        "ok": False,
                        "target_company": target_company,
                        "parsed": {},
                        "message": "",
                        "detail": "",
                    }
                    try:
                        effective_date_mode = date_mode
                        effective_custom_tally_date = custom_tally_date
                        company_gst_registrations = []
                        existing_ledger_names = set()

                        self.after(0, lambda: self._push_message_var.set("Fetching company GST registrations..."))
                        gst_reg_result = _fetch_company_gst_registrations(
                            tally_url,
                            company_name=company,
                            timeout=15,
                        )
                        if gst_reg_result.get("success"):
                            company_gst_registrations = gst_reg_result.get("registrations", [])

                        self.after(0, lambda: self._push_message_var.set("Checking existing ledgers in Tally..."))
                        existing_ledger_result = _fetch_existing_ledger_names(
                            tally_url,
                            company_name=company,
                            timeout=15,
                        )
                        if existing_ledger_result.get("success"):
                            existing_ledger_names = set(existing_ledger_result.get("ledgers") or set())

                        auto_ledger_defs = _collect_auto_voucher_ledgers(rows_snapshot, mode)
                        auto_ledger_defs_to_create = _filter_out_existing_ledgers(
                            auto_ledger_defs,
                            existing_ledger_names,
                        )
                        if auto_ledger_defs_to_create:
                            self.after(0, lambda: self._push_message_var.set("Creating required ledgers in Tally..."))
                            auto_ledger_xml = generate_ledger_xml(auto_ledger_defs_to_create, company)
                            auto_ledger_resp = push_to_tally(auto_ledger_xml, host, port_value)
                            auto_ledger_parsed = _parse_tally_response_details(auto_ledger_resp)
                            self._append_debug_log(
                                "auto-ledger",
                                target_company,
                                auto_ledger_xml,
                                auto_ledger_resp,
                                auto_ledger_parsed,
                                note=(
                                    f"mode={mode}, ledgers_total={len(auto_ledger_defs)}, "
                                    f"new_ledgers={len(auto_ledger_defs_to_create)}"
                                ),
                            )
                            existing_ledger_names.update(
                                str(entry.get("Name", "") or "").strip()
                                for entry in auto_ledger_defs_to_create
                                if str(entry.get("Name", "") or "").strip()
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
                                f"Posting to Tally ({target_company}, {date_mode_label}, {vno_label})..."
                            ),
                        )
                        self.after(0, lambda: self._push_message_var.set("Posting voucher data to Tally..."))

                        xml = build_voucher_xml(
                            mode,
                            effective_date_mode,
                            effective_custom_tally_date,
                            next_voucher,
                            rows_snapshot,
                            company_gst_registrations,
                        )
                        resp = push_to_tally(xml, host, port_value)
                        parsed = _parse_tally_response_details(resp)
                        self._append_debug_log(
                            mode,
                            target_company,
                            xml,
                            resp,
                            parsed,
                            note=f"voucher_type={voucher_type}, date_mode={effective_date_mode}, {voucher_note}",
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
                            if date_mode == "excel":
                                self.after(0, lambda: self._push_message_var.set("Retrying with today date..."))
                                retry_xml = build_voucher_xml(
                                    mode,
                                    "current",
                                    "",
                                    next_voucher,
                                    rows_snapshot,
                                    company_gst_registrations,
                                )
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
                                    effective_date_mode = "current"
                                    effective_custom_tally_date = ""

                            if not result.get("ok"):
                                missing_ledger_defs = _build_missing_ledger_defs(
                                    parsed.get("line_errors") or [],
                                    rows_snapshot,
                                    mode,
                                )
                                missing_ledger_defs = _filter_out_existing_ledgers(
                                    missing_ledger_defs,
                                    existing_ledger_names,
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
                                        existing_ledger_names.update(
                                            str(entry.get("Name", "") or "").strip()
                                            for entry in missing_ledger_defs
                                            if str(entry.get("Name", "") or "").strip()
                                        )
                                        self.after(0, lambda: self._push_message_var.set("Retrying voucher post after ledger creation..."))
                                        post_ledger_retry_xml = build_voucher_xml(
                                            mode,
                                            effective_date_mode,
                                            effective_custom_tally_date,
                                            next_voucher,
                                            rows_snapshot,
                                            company_gst_registrations,
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

                            if (
                                not result.get("ok")
                                and mode == "item"
                                and voucher_type == "Sales"
                                and parsed.get("exceptions", 0) > 0
                                and not (parsed.get("line_errors") or [])
                            ):
                                self.after(
                                    0,
                                    lambda: self._push_message_var.set(
                                        "Retrying sales item invoice with compatibility header..."
                                    ),
                                )
                                compatibility_xml = build_voucher_xml(
                                    mode,
                                    effective_date_mode,
                                    effective_custom_tally_date,
                                    next_voucher,
                                    rows_snapshot,
                                    company_gst_registrations,
                                    legacy_item_invoice_context=True,
                                )
                                compatibility_resp = push_to_tally(compatibility_xml, host, port_value)
                                compatibility_parsed = _parse_tally_response_details(compatibility_resp)
                                self._append_debug_log(
                                    mode,
                                    target_company,
                                    compatibility_xml,
                                    compatibility_resp,
                                    compatibility_parsed,
                                    note=(
                                        f"voucher_type={voucher_type}, compatibility_retry_legacy_item_header, "
                                        f"{voucher_note}, date_mode={effective_date_mode}"
                                    ),
                                )
                                if compatibility_parsed.get("success"):
                                    result = {
                                        "ok": True,
                                        "target_company": target_company,
                                        "parsed": compatibility_parsed,
                                        "message": (
                                            "Posted to Tally successfully after retrying the sales item invoice "
                                            "with the compatibility header.\n\n"
                                            f"Target Company: {target_company}\n"
                                            f"Created: {compatibility_parsed.get('created', 0)}\n"
                                            f"Altered: {compatibility_parsed.get('altered', 0)}\n"
                                            f"Ignored: {compatibility_parsed.get('ignored', 0)}"
                                        ),
                                    }
                                else:
                                    parsed = compatibility_parsed

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

        ctk.CTkLabel(form, text="Ledger Name *", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_name"] = ctk.CTkEntry(
            form,
            placeholder_text="e.g. ABC Traders",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        fields["led_name"].pack(fill="x", padx=12, pady=(0,2))

        ctk.CTkLabel(form, text="Parent Group *", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_parent"] = ctk.CTkComboBox(
            form,
            values=LEDGER_PARENT_OPTIONS,
            state="readonly",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            button_color=COLORS["accent"],
            button_hover_color=COLORS["accent_hover"],
            text_color=COLORS["text_primary"],
        )
        fields["led_parent"].set(LEDGER_PARENT_OPTIONS[0])
        fields["led_parent"].pack(fill="x", padx=12, pady=(0,2))

        ctk.CTkLabel(form, text="GST Applicable", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_gst_app"] = ctk.CTkComboBox(
            form,
            values=LEDGER_GST_APPLICABLE_OPTIONS,
            state="readonly",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            button_color=COLORS["accent"],
            button_hover_color=COLORS["accent_hover"],
            text_color=COLORS["text_primary"],
        )
        fields["led_gst_app"].set("Not Applicable")
        fields["led_gst_app"].pack(fill="x", padx=12, pady=(0,2))

        ctk.CTkLabel(form, text="GSTIN", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_gstin"] = ctk.CTkEntry(
            form,
            placeholder_text="e.g. 07AAACR1718Q1ZZ",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        fields["led_gstin"].pack(fill="x", padx=12, pady=(0,2))

        ctk.CTkLabel(form, text="State", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_state"] = ctk.CTkComboBox(
            form,
            values=LEDGER_STATE_OPTIONS,
            state="readonly",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            button_color=COLORS["accent"],
            button_hover_color=COLORS["accent_hover"],
            text_color=COLORS["text_primary"],
        )
        fields["led_state"].set("Not Applicable")
        fields["led_state"].pack(fill="x", padx=12, pady=(0,2))

        ctk.CTkLabel(form, text="Address Line 1", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_addr1"] = ctk.CTkEntry(
            form,
            placeholder_text="e.g. Street / Building",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        fields["led_addr1"].pack(fill="x", padx=12, pady=(0,2))

        ctk.CTkLabel(form, text="Address Line 2", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_addr2"] = ctk.CTkEntry(
            form,
            placeholder_text="e.g. Area / Locality",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        fields["led_addr2"].pack(fill="x", padx=12, pady=(0,2))

        ctk.CTkLabel(form, text="GST Rate %", font=("Segoe UI", 11), text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(6,0))
        fields["led_gst_rate"] = ctk.CTkEntry(
            form,
            placeholder_text="e.g. 9",
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        fields["led_gst_rate"].pack(fill="x", padx=12, pady=(0,2))

        def _set_field(widget, value=""):
            text = str(value or "").strip()
            if isinstance(widget, ctk.CTkComboBox):
                options = list(widget.cget("values") or [])
                if text in options:
                    widget.set(text)
                elif not text and options:
                    widget.set(options[0])
                else:
                    widget.set(text)
                return
            widget.delete(0, "end")
            if text:
                widget.insert(0, text)

        def _clear_ledger_form():
            _set_field(fields["led_name"], "")
            _set_field(fields["led_parent"], LEDGER_PARENT_OPTIONS[0])
            _set_field(fields["led_gst_app"], "Not Applicable")
            _set_field(fields["led_gstin"], "")
            _set_field(fields["led_state"], "Not Applicable")
            _set_field(fields["led_addr1"], "")
            _set_field(fields["led_addr2"], "")
            _set_field(fields["led_gst_rate"], "")

        self._ledger_list = []
        ledger_edit_index = None

        def add_ledger():
            nonlocal ledger_edit_index
            name = fields["led_name"].get().strip()
            parent_grp = fields["led_parent"].get().strip()
            if not name or not parent_grp:
                messagebox.showwarning("Required","Ledger Name and Parent Group are required.")
                return

            gst_app = fields["led_gst_app"].get().strip()
            gstin = fields["led_gstin"].get().strip().upper()
            state = _normalize_state_for_ledger(fields["led_state"].get().strip())
            address1 = fields["led_addr1"].get().strip()
            address2 = fields["led_addr2"].get().strip()
            is_party_parent = _is_party_parent(parent_grp)

            entry = {
                "Name": name, "Parent": parent_grp,
                "GSTApplicable": gst_app,
                "GSTIN": gstin,
                "StateOfSupply": state,
                "Address1": address1,
                "Address2": address2,
                "MailingName": name,
                "Country": "India",
                "Pincode": "",
                "Billwise": "Yes" if is_party_parent else "No",
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

            _clear_ledger_form()
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
            _set_field(fields["led_name"], entry.get("Name", ""))
            _set_field(fields["led_parent"], entry.get("Parent", ""))
            _set_field(fields["led_gst_app"], entry.get("GSTApplicable", ""))
            _set_field(fields["led_gstin"], entry.get("GSTIN", ""))
            _set_field(fields["led_state"], entry.get("StateOfSupply", "Not Applicable") or "Not Applicable")
            _set_field(fields["led_addr1"], entry.get("Address1", ""))
            _set_field(fields["led_addr2"], entry.get("Address2", ""))
            _set_field(fields["led_gst_rate"], entry.get("GSTRate", ""))

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
            _clear_ledger_form()

        def export_ledgers(action):
            company = self._get_selected_company()
            if not self._ledger_list:
                messagebox.showwarning("Empty","Add at least one ledger.")
                return
            if action == "save":
                xml = generate_ledger_xml(self._ledger_list, company)
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
                    ledgers_to_push = list(self._ledger_list)
                    existing_ledger_result = _fetch_existing_ledger_names(
                        tally_url,
                        company_name=company,
                        timeout=15,
                    )
                    if existing_ledger_result.get("success"):
                        ledgers_to_push = _filter_out_existing_ledgers(
                            ledgers_to_push,
                            existing_ledger_result.get("ledgers") or set(),
                        )
                    if not ledgers_to_push:
                        self.status_var.set(f"No new ledgers to create in Tally ({target_company})")
                        messagebox.showinfo(
                            "Nothing To Create",
                            "All queued ledgers already exist in Tally, so nothing was changed.",
                        )
                        return

                    xml = generate_ledger_xml(ledgers_to_push, company)
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
        ctk.CTkButton(btn_row, text="📂 Import XML", fg_color=COLORS["warning"], hover_color="#B45309",
                       text_color="#FFFFFF",
                       command=lambda: self._import_xml_direct("ledger")).grid(row=0, column=2, sticky="ew", padx=(6,0), pady=(0,6))
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
                        "Name": _row_text(r, "Name") or _row_text(r, "LedgerName") or _row_text(r, "Ledger Name"),
                        "Parent": _row_text(r, "Parent") or _row_text(r, "ParentGroup") or _row_text(r, "Parent Group") or "Sundry Debtors",
                        "GSTApplicable": _row_text(r, "GSTApplicable") or _row_text(r, "GST Applicable") or _row_text(r, "GST"),
                        "GSTIN": (_row_text(r, "GSTIN") or _row_text(r, "PartyGSTIN") or _row_text(r, "Party GSTIN")).upper(),
                        "StateOfSupply": _row_text(r, "State") or _row_text(r, "StateOfSupply") or _row_text(r, "PlaceOfSupply") or _row_text(r, "Place Of Supply"),
                        "MailingName": _row_text(r, "MailingName") or _row_text(r, "PartyMailingName") or _row_text(r, "Mailing Name"),
                        "Country": _row_text(r, "Country") or _row_text(r, "CountryOfResidence") or _row_text(r, "Country Of Residence") or "India",
                        "Pincode": _row_text(r, "Pincode") or _row_text(r, "PinCode") or _row_text(r, "PIN") or _row_text(r, "PostalCode"),
                        "Billwise": _row_text(r, "Billwise") or _row_text(r, "IsBillwise") or _row_text(r, "ISBILLWISEON"),
                        "Address1": _row_text(r, "Address1") or _row_text(r, "Address Line 1") or _row_text(r, "AddressLine1") or _row_text(r, "Address"),
                        "Address2": _row_text(r, "Address2") or _row_text(r, "Address Line 2") or _row_text(r, "AddressLine2"),
                        "TypeOfTaxation": _row_text(r, "TaxType") or _row_text(r, "TypeOfTaxation") or _row_text(r, "Tax Type"),
                        "GSTRate": _row_text(r, "GSTRate") or _row_text(r, "GST Rate") or _row_text(r, "Rate"),
                        "GSTRegistrationType": _row_text(r, "GSTRegistrationType") or _row_text(r, "GST Registration Type") or _row_text(r, "RegistrationType") or _row_text(r, "RegType"),
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
        ctk.CTkButton(btn_row, text="📂 Import XML", fg_color=COLORS["warning"], hover_color="#B45309",
                       text_color="#FFFFFF",
                       command=lambda: self._import_xml_direct("stock")).grid(row=0, column=2, sticky="ew", padx=(6,0), pady=(0,6))
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
    # ─── PANEL SWITCHING ──────────────────────────────────────────────────────

    def _switch_panel(self, selected_label: str):
        """Raise the panel matching the dropdown selection."""
        key = self._panel_option_to_key.get(selected_label, "accounting")
        # Update note type var when a note option is selected
        if key == "credit_note":
            self._note_type_var.set("Credit Note")
            if hasattr(self, "_note_type_display_label"):
                self._note_type_display_label.configure(text="Mode: Credit Note 📝",
                                                         text_color=COLORS["success"])
        elif key == "debit_note":
            self._note_type_var.set("Debit Note")
            if hasattr(self, "_note_type_display_label"):
                self._note_type_display_label.configure(text="Mode: Debit Note 📒",
                                                         text_color=COLORS["warning"])
        elif key == "journal":
            pass  # no special state to set
        panel = self._panels.get(key)
        if panel:
            panel.tkraise()

    # ─── PREVIEW MODE TOGGLE ──────────────────────────────────────────────────

    def _show_preview_mode(self, excel_frame, xml_frame, mode: str,
                           excel_btn, xml_btn, mode_key: str = ""):
        """Raise the correct preview frame, update toggle styles, track active view."""
        if mode_key:
            self._active_preview_mode[mode_key] = mode
        if mode == "xml":
            xml_frame.tkraise()
            xml_btn.configure(fg_color=COLORS["accent"], text_color="#FFFFFF")
            excel_btn.configure(fg_color=COLORS["bg_input"], text_color=COLORS["text_muted"])
        else:
            excel_frame.tkraise()
            excel_btn.configure(fg_color=COLORS["accent"], text_color="#FFFFFF")
            xml_btn.configure(fg_color=COLORS["bg_input"], text_color=COLORS["text_muted"])

    # ─── XML PREVIEW ──────────────────────────────────────────────────────────

    def _load_xml_preview(self, filepath: str, xml_text_widget, mode: str, info_label):
        """Read an XML file and show its pretty-printed content in the text widget."""
        try:
            with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                raw = f.read()

            import xml.dom.minidom as _minidom
            try:
                dom = _minidom.parseString(raw.encode("utf-8", errors="replace"))
                pretty = dom.toprettyxml(indent="  ")
                # Strip the <?xml ...?> declaration for cleaner display
                lines = pretty.splitlines()
                if lines and lines[0].startswith("<?xml"):
                    pretty = "\n".join(lines[1:])
            except Exception:
                pretty = raw  # show raw if XML is malformed

            xml_text_widget.configure(state="normal")
            xml_text_widget.delete("1.0", "end")
            xml_text_widget.insert("end", pretty)
            xml_text_widget.configure(state="disabled")

            # Count vouchers for info label
            import xml.etree.ElementTree as _ET
            count_str = ""
            try:
                root = _ET.fromstring(raw)
                vouchers = root.findall(".//VOUCHER")
                masters = root.findall(".//LEDGER") + root.findall(".//STOCKITEM")
                if vouchers:
                    count_str = f"📄 {len(vouchers)} voucher(s)"
                elif masters:
                    count_str = f"📄 {len(masters)} master(s)"
                else:
                    count_str = "📄 XML loaded"
            except Exception:
                count_str = "📄 XML loaded"

            info_text = f"{count_str} — {len(raw):,} bytes"
            if info_label:
                info_label.configure(text=info_text)
            self.status_var.set(f"XML loaded: {os.path.basename(filepath)}")
        except Exception as e:
            messagebox.showerror("XML Load Error", str(e))

    # ─── IMPORT XML (Voucher modes) ───────────────────────────────────────────

    def _import_xml(self, mode: str, xml_fp_var):
        """Push an existing XML file to Tally (used by voucher tabs)."""
        filepath = xml_fp_var.get() if hasattr(xml_fp_var, "get") else str(xml_fp_var)
        if not filepath or not os.path.isfile(filepath):
            messagebox.showwarning("No XML File",
                                   "Please browse and select an XML file first using the Browse XML button.")
            return
        if self._push_running:
            self.status_var.set("Push already in progress. Please wait...")
            return

        host = self.tally_host_var.get() or "localhost"
        port_text = self.tally_port_var.get() or "9000"
        try:
            port = int(port_text)
        except ValueError:
            messagebox.showerror("Invalid Port", f"Invalid port number: {port_text}")
            return

        try:
            with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                xml_content = f.read()
        except Exception as e:
            messagebox.showerror("File Read Error", str(e))
            return

        self._set_push_loading_state(True, f"Importing XML to Tally...")

        def worker():
            try:
                resp = push_to_tally(xml_content, host, port)
                parsed = _parse_tally_response_details(resp)
                def done():
                    self._set_push_loading_state(False)
                    if parsed.get("success"):
                        self.status_var.set("XML imported to Tally successfully")
                        messagebox.showinfo(
                            "Import Successful",
                            f"XML file imported to Tally successfully.\n\n"
                            f"File: {os.path.basename(filepath)}\n"
                            f"Created: {parsed.get('created', 0)}\n"
                            f"Altered: {parsed.get('altered', 0)}\n"
                            f"Ignored: {parsed.get('ignored', 0)}",
                        )
                    else:
                        err = parsed.get("error") or "Unknown Tally error"
                        self.status_var.set("Import failed")
                        messagebox.showerror("Import Failed",
                                             f"Tally rejected the import:\n\n{err}\n\n"
                                             f"Raw response (first 500 chars):\n{resp[:500]}")
                self.after(0, done)
            except Exception as e:
                def done_err():
                    self._set_push_loading_state(False)
                    messagebox.showerror("Error", str(e))
                self.after(0, done_err)

        threading.Thread(target=worker, daemon=True).start()

    # ─── IMPORT XML DIRECT (Ledger/Stock tabs — browse a file) ───────────────

    def _import_xml_direct(self, panel_hint: str = ""):
        """Browse for an XML file and push it directly to Tally (ledger/stock tabs)."""
        if self._push_running:
            self.status_var.set("Push already in progress. Please wait...")
            return

        f = filedialog.askopenfilename(
            title="Select XML file to import into Tally",
            filetypes=[("XML Files", "*.xml"), ("All Files", "*.*")],
        )
        if not f:
            return

        host = self.tally_host_var.get() or "localhost"
        port_text = self.tally_port_var.get() or "9000"
        try:
            port = int(port_text)
        except ValueError:
            messagebox.showerror("Invalid Port", f"Invalid port number: {port_text}")
            return

        try:
            with open(f, "r", encoding="utf-8", errors="replace") as fh:
                xml_content = fh.read()
        except Exception as e:
            messagebox.showerror("File Read Error", str(e))
            return

        self._set_push_loading_state(True, f"Importing {os.path.basename(f)} to Tally...")

        def worker():
            try:
                resp = push_to_tally(xml_content, host, port)
                parsed = _parse_tally_response_details(resp)
                def done():
                    self._set_push_loading_state(False)
                    if parsed.get("success"):
                        self.status_var.set("XML imported to Tally successfully")
                        messagebox.showinfo(
                            "Import Successful",
                            f"XML file imported to Tally successfully.\n\n"
                            f"File: {os.path.basename(f)}\n"
                            f"Created: {parsed.get('created', 0)}\n"
                            f"Altered: {parsed.get('altered', 0)}\n"
                            f"Ignored: {parsed.get('ignored', 0)}",
                        )
                    else:
                        err = parsed.get("error") or "Unknown Tally error"
                        self.status_var.set("Import failed")
                        messagebox.showerror("Import Failed",
                                             f"Tally rejected the import:\n\n{err}\n\n"
                                             f"Raw response (first 500 chars):\n{resp[:500]}")
                self.after(0, done)
            except Exception as e:
                def done_err():
                    self._set_push_loading_state(False)
                    messagebox.showerror("Error", str(e))
                self.after(0, done_err)

        threading.Thread(target=worker, daemon=True).start()
    # ─── SMART PUSH (context-aware: Excel or XML based on active toggle) ─────

    def _smart_push(self, mode: str, fp_var, xml_fp_var):
        """Push to Tally using whichever preview is currently active."""
        active = self._active_preview_mode.get(mode, "excel")
        if active == "xml":
            self._import_xml(mode, xml_fp_var)
        else:
            self._generate(mode, "push", fp_var.get() if hasattr(fp_var, "get") else str(fp_var))

    # ─── SMART PUSH FOR NOTE PANEL ────────────────────────────────────────────

    def _smart_push_note(self, note_fp_var, note_xml_fp_var):
        """Smart push for the Credit/Debit Note panel."""
        active = self._active_preview_mode.get("note", "excel")
        if active == "xml":
            self._import_xml("note", note_xml_fp_var)
        else:
            self._generate_note("push", note_fp_var.get() if hasattr(note_fp_var, "get") else "")

    # ─── NOTE GENERATION ─────────────────────────────────────────────────────

    def _generate_note(self, action: str, filepath: str):
        if self._push_running:
            self.status_var.set("Push already in progress...")
            return

        note_type = self._note_type_var.get() or "Credit Note"

        # Determine rows based on active tab
        if hasattr(self, '_note_source_tabs') and self._note_source_tabs is not None:
            active_tab = self._note_source_tabs.get()
            if active_tab == "Manual Entry":
                rows = getattr(self, '_note_manual_rows', [])
                source_label = "Manual Entry"
            else:
                rows = getattr(self, '_note_loaded_rows', [])
                source_label = "Excel Upload"
        else:
            rows = getattr(self, '_note_loaded_rows', [])
            source_label = "Excel Upload"

        if not rows:
            messagebox.showwarning("No Data", f"No rows available in {source_label}.")
            return

        company = self._get_selected_company()
        if action == "push" and not company and len(getattr(self, "fetched_companies", [])) > 1:
            messagebox.showwarning("Select Company", "Please select a target company before pushing.")
            return

        try:
            date_mode, custom_tally_date = self._get_voucher_date_selection()

            _cmp_gst_regs = list(self.company_gst_registrations or [])
            if _normalize_note_type(note_type) == "Debit Note" and not _cmp_gst_regs:
                try:
                    _tally_url = self._get_tally_url()
                    _gst_fetch = _fetch_company_gst_registrations(_tally_url, company_name=company, timeout=10)
                    if _gst_fetch.get("success"):
                        _cmp_gst_regs = _gst_fetch.get("registrations", [])
                        self.company_gst_registrations = _cmp_gst_regs
                except Exception:
                    pass

            xml_payload, voucher_count = generate_note_xml(
                rows,
                company=company,
                date_mode=date_mode,
                custom_tally_date=custom_tally_date,
                voucher_type=note_type,
                company_gst_registrations=_cmp_gst_regs,
            )
            if voucher_count <= 0:
                messagebox.showwarning("No Vouchers", "No valid rows (TaxableValue must be > 0).")
                return

            if action == "save":
                stem = note_type.replace(" ", "")
                out = filedialog.asksaveasfilename(
                    defaultextension=".xml",
                    initialfile=f"{stem}.xml",
                    filetypes=[("XML", "*.xml")],
                )
                if not out:
                    return
                with open(out, "w", encoding="utf-8") as f:
                    f.write(xml_payload)
                self.status_var.set(f"{note_type} XML saved: {os.path.basename(out)}")
                messagebox.showinfo("Saved", f"{note_type} XML saved.\n{out}")
                return

            host = (self.tally_host_var.get() or "localhost").strip()
            port = int((self.tally_port_var.get() or "9000").strip())

            self._set_push_loading_state(True, f"Pushing {voucher_count} {note_type} voucher(s)...")
            self.status_var.set("Pushing to Tally...")

            def worker():
                try:
                    resp = push_to_tally(xml_payload, host=host, port=port)
                    parsed = _parse_tally_response_details(resp)
                    result = {"ok": True, "parsed": parsed}
                except Exception as exc:
                    result = {"ok": False, "error": str(exc)}

                def done():
                    self._set_push_loading_state(False)
                    if not result.get("ok"):
                        messagebox.showerror("Push Failed", str(result.get("error", "Unknown error")))
                        return
                    p = result["parsed"]
                    created, altered = p.get("created", 0), p.get("altered", 0)
                    errors = p.get("errors", 0)
                    summary = f"Created: {created}\nAltered: {altered}\nErrors: {errors}"
                    if p.get("success"):
                        self.status_var.set(f"{note_type} pushed: {created} created, {altered} altered")
                        messagebox.showinfo("Push Successful", summary)
                    else:
                        line_errors = p.get("line_errors", [])
                        if line_errors:
                            summary += "\n\nLine Errors:\n- " + "\n- ".join(line_errors[:8])
                        self.status_var.set(f"{note_type} push completed with errors.")
                        messagebox.showwarning("Push With Errors", summary)
                self.after(0, done)

            threading.Thread(target=worker, daemon=True).start()

        except ValueError as exc:
            messagebox.showerror("Validation Error", str(exc))
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    # ─── NOTE PANEL BUILDER ───────────────────────────────────────────────────

    NOTE_TEMPLATE_HEADERS = [
        "Date", "VoucherNo", "GSTIN", "PartyLedger", "Particular",
        "TaxableValue", "CGSTLedger", "CGSTRate", "SGSTLedger", "SGSTRate",
        "IGSTLedger", "IGSTRate", "Narration",
    ]

    def _build_note_panel(self, parent):
        """Build the shared Credit / Debit Note panel (tabbed: Excel Upload, Manual Entry, Create Party)."""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        # Row 0: Note type indicator + template
        top_row = ctk.CTkFrame(parent, fg_color="transparent")
        top_row.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))

        type_badge = ctk.CTkFrame(top_row, fg_color=COLORS["bg_input"], corner_radius=8)
        type_badge.pack(side="left")
        self._note_type_display_label = ctk.CTkLabel(
            type_badge, text="Mode: Credit Note 📝",
            font=("Segoe UI", 11, "bold"), text_color=COLORS["success"], padx=10, pady=4)
        self._note_type_display_label.pack()

        ctk.CTkLabel(top_row,
                     text="Select Credit Note or Debit Note from the dropdown above",
                     font=("Segoe UI", 10), text_color=COLORS["text_muted"]).pack(side="left", padx=12)

        note_template_btn = ctk.CTkButton(
            top_row, text="📥  Download Template",
            fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"], width=170,
            command=self._download_note_template)
        note_template_btn.pack(side="right")

        # Row 1: Tabview
        self._note_source_tabs = ctk.CTkTabview(
            parent,
            fg_color="transparent",
            segmented_button_fg_color=COLORS["bg_input"],
            segmented_button_selected_color=COLORS["accent"],
            segmented_button_selected_hover_color=COLORS["accent_hover"],
            segmented_button_unselected_color=COLORS["bg_input"],
            segmented_button_unselected_hover_color=COLORS["bg_card_hover"],
        )
        self._note_source_tabs.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 4))

        excel_tab = self._note_source_tabs.add("Excel Upload")
        manual_tab = self._note_source_tabs.add("Manual Entry")
        party_tab = self._note_source_tabs.add("Create Party Ledger")
        self._note_source_tabs.set("Excel Upload")

        self._build_note_excel_tab(excel_tab)
        self._build_note_manual_tab(manual_tab)
        self._build_note_create_party_tab(party_tab)

        # Row 2: Action buttons
        btn_row = ctk.CTkFrame(parent, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))

        ctk.CTkButton(
            btn_row, text="💾  Save XML File", fg_color=SUCCESS, hover_color="#15803D", width=155,
            command=lambda: self._generate_note("save", "")
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_row, text="🚀  Push to Tally", fg_color=ACCENT, hover_color=ACCENT_HOVER, width=165,
            command=lambda: self._generate_note("push", "")
        ).pack(side="left", padx=(0, 8))

    def _build_note_excel_tab(self, parent):
        """Excel Upload sub-tab for the Note panel."""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        load_frame = ctk.CTkFrame(parent, fg_color="transparent")
        load_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        note_fp_var = ctk.StringVar()
        ctk.CTkEntry(load_frame, textvariable=note_fp_var,
                     placeholder_text="Select Credit/Debit Note Excel (.xlsx/.xlsm)...",
                     state="readonly", fg_color=COLORS["bg_input"], border_color=COLORS["border"],
                     ).pack(side="left", padx=(0, 8), fill="x", expand=True)

        def browse_note_excel():
            f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
            if not f:
                return
            note_fp_var.set(f)
            try:
                headers, rows = read_excel(f)
                self._note_loaded_headers = headers
                self._note_loaded_rows = rows
                if self._note_excel_tree is not None:
                    self._note_excel_tree.delete(*self._note_excel_tree.get_children())
                    self._note_excel_tree["columns"] = headers
                    for col in headers:
                        self._note_excel_tree.heading(col, text=col)
                        self._note_excel_tree.column(col, width=max(80, len(col) * 9), minwidth=50)
                    for row in rows[:300]:
                        self._note_excel_tree.insert("", "end", values=[row.get(h, "") for h in headers])
                if self._note_excel_info_label is not None:
                    self._note_excel_info_label.configure(text=f"📊 {len(rows)} row(s) — {len(headers)} column(s)")
                self.status_var.set(f"Loaded: {os.path.basename(f)}")
            except Exception as exc:
                messagebox.showerror("Load Error", str(exc))

        ctk.CTkButton(load_frame, text="Browse Excel", command=browse_note_excel,
                      width=120, fg_color=ACCENT, hover_color=ACCENT_HOVER).pack(side="left")

        info_frame = ctk.CTkFrame(parent, fg_color="transparent")
        info_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 4))
        self._note_excel_info_label = ctk.CTkLabel(info_frame, text="", font=("Segoe UI", 11),
                                                    text_color=TEXT_MUTED)
        self._note_excel_info_label.pack(side="left")

        tree_frame = ctk.CTkFrame(parent, fg_color=COLORS["bg_dark"],
                                   corner_radius=8, border_width=1, border_color=COLORS["border"])
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        _sy = ttk.Scrollbar(tree_frame, orient="vertical")
        _sx = ttk.Scrollbar(tree_frame, orient="horizontal")
        self._note_excel_tree = ttk.Treeview(tree_frame, show="headings",
                                              yscrollcommand=_sy.set, xscrollcommand=_sx.set)
        _sy.config(command=self._note_excel_tree.yview)
        _sx.config(command=self._note_excel_tree.xview)
        self._note_excel_tree.grid(row=0, column=0, sticky="nsew")
        _sy.grid(row=0, column=1, sticky="ns")
        _sx.grid(row=1, column=0, sticky="ew")

    def _build_note_manual_tab(self, parent):
        """Manual Entry sub-tab for the Note panel."""
        _NOTE_HDRS = self.NOTE_TEMPLATE_HEADERS
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.pack(fill="both", expand=True, padx=10, pady=8)
        wrapper.grid_columnconfigure(0, weight=1, uniform="note_manual_split")
        wrapper.grid_columnconfigure(1, weight=1, uniform="note_manual_split")
        wrapper.grid_rowconfigure(0, weight=1)

        left_panel = ctk.CTkFrame(wrapper, fg_color=COLORS["bg_card"],
                                   border_width=1, border_color=COLORS["border"], corner_radius=10)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left_panel.grid_columnconfigure(0, weight=1)
        left_panel.grid_rowconfigure(0, weight=1)

        form_scroll = ctk.CTkScrollableFrame(left_panel, fg_color="transparent", corner_radius=8)
        form_scroll.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 6))
        form_card = ctk.CTkFrame(form_scroll, fg_color="transparent")
        form_card.pack(fill="x", padx=2, pady=2)

        fields = [
            ("Date", "Date", "DD-MM-YY"),
            ("VoucherNo", "Voucher No", "Optional"),
            ("GSTIN", "GSTIN", "Optional"),
            ("PartyLedger", "Party Ledger", "Required"),
            ("Particular", "Particular", "Required"),
            ("TaxableValue", "Taxable Value", "Required Amount"),
            ("CGSTLedger", "CGST Ledger", "Optional"),
            ("CGSTRate", "CGST Rate", "0"),
            ("SGSTLedger", "SGST Ledger", "Optional"),
            ("SGSTRate", "SGST Rate", "0"),
            ("IGSTLedger", "IGST Ledger", "Optional"),
            ("IGSTRate", "IGST Rate", "0"),
            ("Narration", "Narration", "Optional"),
        ]

        cols = 2
        for i, (key, label, placeholder) in enumerate(fields):
            row_idx = i // cols
            col_idx = i % cols
            field_wrap = ctk.CTkFrame(form_card, fg_color="transparent")
            field_wrap.grid(row=row_idx, column=col_idx, sticky="ew", padx=8, pady=6)
            form_card.grid_columnconfigure(col_idx, weight=1)

            if key == "PartyLedger":
                top_line = ctk.CTkFrame(field_wrap, fg_color="transparent")
                top_line.pack(fill="x")
                ctk.CTkLabel(top_line, text=label, font=("Segoe UI", 10),
                             text_color=COLORS["text_secondary"]).pack(side="left")
                self._note_manual_fetch_ledger_btn = ctk.CTkButton(
                    top_line, text="Fetch", width=70, height=26,
                    font=("Segoe UI", 10, "bold"),
                    fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
                    text_color=COLORS["text_secondary"],
                    command=self._note_fetch_party_ledgers_thread)
                self._note_manual_fetch_ledger_btn.pack(side="right")

                search_row = ctk.CTkFrame(field_wrap, fg_color="transparent")
                search_row.pack(fill="x", pady=(4, 2))
                search_entry = ctk.CTkEntry(
                    search_row, textvariable=self._note_manual_party_search_var,
                    placeholder_text="Search party ledger...", height=34,
                    fg_color=COLORS["bg_input"], border_color=COLORS["border"], font=("Segoe UI", 10))
                search_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
                search_entry.bind("<KeyRelease>", self._on_note_party_search_change)

                self._note_manual_party_search_clear_btn = ctk.CTkButton(
                    search_row, text="Clear", width=58, height=30,
                    fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
                    text_color=COLORS["text_secondary"], font=("Segoe UI", 9, "bold"),
                    command=lambda: (self._note_manual_party_search_var.set(""), self._on_note_party_search_change()))
                self._note_manual_party_search_clear_btn.pack(side="right")

                var = ctk.StringVar(value="")
                combo = ctk.CTkComboBox(field_wrap, values=[""], variable=var, height=36,
                                        fg_color=COLORS["bg_input"], border_color=COLORS["border"],
                                        button_color=COLORS["accent"], button_hover_color=COLORS["accent_hover"],
                                        font=("Segoe UI", 10), state="readonly")
                combo.pack(fill="x", pady=(0, 2))
                self._note_manual_party_match_label = ctk.CTkLabel(
                    field_wrap, text="Type in search box after fetching ledgers",
                    font=("Segoe UI", 9), text_color=COLORS["text_muted"])
                self._note_manual_party_match_label.pack(anchor="w")
                self._note_manual_party_search_var.trace_add("write", lambda *_: self._on_note_party_search_change())
                self._note_manual_party_ledger_combo = combo
                self._note_manual_form_vars[key] = var
                continue

            ctk.CTkLabel(field_wrap, text=label, font=("Segoe UI", 10),
                         text_color=COLORS["text_secondary"]).pack(anchor="w")
            var = ctk.StringVar(value="")
            ctk.CTkEntry(field_wrap, textvariable=var, placeholder_text=placeholder,
                         height=36, fg_color=COLORS["bg_input"], border_color=COLORS["border"]).pack(fill="x")
            self._note_manual_form_vars[key] = var

        if "Date" in self._note_manual_form_vars:
            self._note_manual_form_vars["Date"].set(datetime.today().strftime("%d-%m-%y"))

        btn_row = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        for ci in range(3):
            btn_row.grid_columnconfigure(ci, weight=1)

        add_btn = ctk.CTkButton(btn_row, text="Add Entry", fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                height=34, command=self._note_add_manual_entry)
        add_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=(0, 6))

        edit_btn = ctk.CTkButton(btn_row, text="Edit Selected", fg_color="#0EA5E9", hover_color="#0284C7",
                                  text_color="#FFFFFF", height=34, command=self._note_edit_selected_manual)
        edit_btn.grid(row=0, column=1, sticky="ew", padx=3, pady=(0, 6))

        self._note_manual_update_btn = ctk.CTkButton(
            btn_row, text="Update Entry", fg_color="#10B981", hover_color="#059669",
            text_color="#FFFFFF", height=34, state="disabled", command=self._note_update_manual_entry)
        self._note_manual_update_btn.grid(row=0, column=2, sticky="ew", padx=(6, 0), pady=(0, 6))

        ctk.CTkButton(btn_row, text="Clear Form", fg_color=COLORS["bg_input"],
                      hover_color=COLORS["bg_card_hover"], text_color=COLORS["text_secondary"],
                      height=34, command=self._note_clear_manual_form).grid(row=1, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(btn_row, text="Remove Selected", fg_color=COLORS["warning"], hover_color="#B45309",
                      text_color="#FFFFFF", height=34, command=self._note_remove_selected_manual).grid(
            row=1, column=1, sticky="ew", padx=3)

        ctk.CTkButton(btn_row, text="Clear All", fg_color=COLORS["error"], hover_color="#B91C1C",
                      text_color="#FFFFFF", height=34, command=self._note_clear_all_manual).grid(
            row=1, column=2, sticky="ew", padx=(6, 0))

        # Right panel: review treeview
        right_panel = ctk.CTkFrame(wrapper, fg_color=COLORS["bg_card"],
                                    border_width=1, border_color=COLORS["border"], corner_radius=10)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(1, weight=1)

        right_header = ctk.CTkFrame(right_panel, fg_color="transparent")
        right_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        ctk.CTkLabel(right_header, text="Review (Excel Format)", font=("Segoe UI", 12, "bold"),
                     text_color=COLORS["text_primary"]).pack(side="left")
        self._note_manual_info_label = ctk.CTkLabel(right_header, text="Manual entries: 0",
                                                     font=("Segoe UI", 11), text_color=TEXT_MUTED)
        self._note_manual_info_label.pack(side="right")

        tree_frame = ctk.CTkFrame(right_panel, fg_color=COLORS["bg_dark"],
                                   corner_radius=8, border_width=1, border_color=COLORS["border"])
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        _sy = ttk.Scrollbar(tree_frame, orient="vertical")
        _sx = ttk.Scrollbar(tree_frame, orient="horizontal")
        self._note_manual_tree = ttk.Treeview(tree_frame, show="headings", selectmode="extended",
                                               yscrollcommand=_sy.set, xscrollcommand=_sx.set)
        _sy.config(command=self._note_manual_tree.yview)
        _sx.config(command=self._note_manual_tree.xview)
        _sy.pack(side="right", fill="y")
        _sx.pack(side="bottom", fill="x")
        self._note_manual_tree.pack(fill="both", expand=True)
        self._note_manual_tree.bind("<Double-1>", lambda e: self._note_edit_selected_manual())

        ctk.CTkLabel(right_panel, text="Double-click a row to edit it.",
                     font=("Segoe UI", 10), text_color=COLORS["text_muted"]).grid(
            row=2, column=0, sticky="w", padx=10, pady=(0, 8))

        self._note_populate_tree(self._note_manual_tree, _NOTE_HDRS, [])

    def _build_note_create_party_tab(self, parent):
        """Create Party Ledger sub-tab for the Note panel."""
        wrapper = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        wrapper.pack(fill="both", expand=True, padx=10, pady=8)

        card = ctk.CTkFrame(wrapper, fg_color=COLORS["bg_card"],
                             border_width=1, border_color=COLORS["border"], corner_radius=10)
        card.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(card, text="Create Party Ledger In Tally",
                     font=("Segoe UI", 13, "bold"), text_color=COLORS["text_primary"]).pack(
            anchor="w", padx=12, pady=(12, 4))
        ctk.CTkLabel(card, text="Uses current Host, Port, and selected Target Company.",
                     font=("Segoe UI", 10), text_color=COLORS["text_muted"]).pack(
            anchor="w", padx=12, pady=(0, 8))

        ctk.CTkLabel(card, text="Ledger Name", font=("Segoe UI", 10),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12)
        self._note_create_party_name_entry = ctk.CTkEntry(
            card, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], placeholder_text="Required party ledger name",
            font=("Segoe UI", 10))
        self._note_create_party_name_entry.pack(fill="x", padx=12, pady=(4, 8))

        row_parent = ctk.CTkFrame(card, fg_color="transparent")
        row_parent.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkLabel(row_parent, text="Parent Group", font=("Segoe UI", 10),
                     text_color=COLORS["text_secondary"]).pack(side="left")
        self._note_create_party_parent_cb = ctk.CTkComboBox(
            row_parent, values=["Sundry Debtors", "Sundry Creditors"], width=200, height=32,
            fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            button_color=COLORS["accent"], button_hover_color=COLORS["accent_hover"], font=("Segoe UI", 10))
        self._note_create_party_parent_cb.set("Sundry Debtors")
        self._note_create_party_parent_cb.pack(side="right")

        ctk.CTkLabel(card, text="Mailing Name", font=("Segoe UI", 10),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", padx=12)
        self._note_create_party_mailing_entry = ctk.CTkEntry(
            card, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], placeholder_text="Optional mailing name", font=("Segoe UI", 10))
        self._note_create_party_mailing_entry.pack(fill="x", padx=12, pady=(4, 8))

        row_gst = ctk.CTkFrame(card, fg_color="transparent")
        row_gst.pack(fill="x", padx=12, pady=(0, 8))
        self._note_create_party_gstin_entry = ctk.CTkEntry(
            row_gst, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], placeholder_text="GSTIN", font=("Segoe UI", 10))
        self._note_create_party_gstin_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self._note_create_party_pincode_entry = ctk.CTkEntry(
            row_gst, width=130, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], placeholder_text="Pincode", font=("Segoe UI", 10))
        self._note_create_party_pincode_entry.pack(side="left", padx=(4, 0))

        row_geo = ctk.CTkFrame(card, fg_color="transparent")
        row_geo.pack(fill="x", padx=12, pady=(0, 8))
        self._note_create_party_state_entry = ctk.CTkEntry(
            row_geo, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], placeholder_text="State", font=("Segoe UI", 10))
        self._note_create_party_state_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self._note_create_party_country_entry = ctk.CTkEntry(
            row_geo, width=130, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], font=("Segoe UI", 10))
        self._note_create_party_country_entry.insert(0, "India")
        self._note_create_party_country_entry.pack(side="left", padx=(4, 0))

        self._note_create_party_address1_entry = ctk.CTkEntry(
            card, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], placeholder_text="Address line 1", font=("Segoe UI", 10))
        self._note_create_party_address1_entry.pack(fill="x", padx=12, pady=(0, 6))
        self._note_create_party_address2_entry = ctk.CTkEntry(
            card, height=34, fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            text_color=COLORS["text_primary"], placeholder_text="Address line 2", font=("Segoe UI", 10))
        self._note_create_party_address2_entry.pack(fill="x", padx=12, pady=(0, 8))

        row_billwise = ctk.CTkFrame(card, fg_color="transparent")
        row_billwise.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkLabel(row_billwise, text="Billwise", font=("Segoe UI", 10),
                     text_color=COLORS["text_secondary"]).pack(side="left")
        self._note_create_party_billwise_cb = ctk.CTkComboBox(
            row_billwise, values=["Yes", "No"], width=120, height=32,
            fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            button_color=COLORS["accent"], button_hover_color=COLORS["accent_hover"], font=("Segoe UI", 10))
        self._note_create_party_billwise_cb.set("Yes")
        self._note_create_party_billwise_cb.pack(side="right")

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x", padx=12, pady=(0, 12))

        self._note_create_party_fetch_btn = ctk.CTkButton(
            btn_row, text="Fetch Party Ledgers", width=160, height=34,
            fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"], command=self._note_fetch_party_ledgers_thread)
        self._note_create_party_fetch_btn.pack(side="left", padx=(0, 8))

        self._note_create_party_clear_btn = ctk.CTkButton(
            btn_row, text="Clear", width=90, height=34,
            fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"], command=self._note_clear_create_party_form)
        self._note_create_party_clear_btn.pack(side="left", padx=(0, 8))

        self._note_create_party_create_btn = ctk.CTkButton(
            btn_row, text="Create Party Ledger", height=34,
            fg_color=COLORS["success"], hover_color="#047857",
            text_color="#FFFFFF", command=self._note_create_party_ledger_thread)
        self._note_create_party_create_btn.pack(side="left", fill="x", expand=True)

        self._note_create_party_status_var = ctk.StringVar(value="Ready to create party ledger")
        ctk.CTkLabel(card, textvariable=self._note_create_party_status_var,
                     font=("Segoe UI", 10), text_color=COLORS["text_muted"]).pack(
            anchor="w", padx=12, pady=(0, 12))

    # ─── NOTE MANUAL ENTRY METHODS ────────────────────────────────────────────

    def _note_populate_tree(self, tree, headers, rows, limit=500):
        tree.delete(*tree.get_children())
        tree["columns"] = headers
        for h in headers:
            tree.heading(h, text=h)
            tree.column(h, width=max(120, min(260, len(h) * 12)), minwidth=80)
        for idx, row in enumerate(rows[:limit]):
            values = [str(_row_get(row, h, "") or "") for h in headers]
            tree.insert("", "end", iid=str(idx), values=values)

    def _on_note_party_search_change(self, _event=None):
        combo = self._note_manual_party_ledger_combo
        if combo is None:
            return
        if not self.fetched_party_ledgers:
            combo.configure(values=[""])
            combo.set("")
            if self._note_manual_party_match_label is not None:
                self._note_manual_party_match_label.configure(text="No fetched party ledgers yet")
            return
        typed_text = (self._note_manual_party_search_var.get() or "").strip()
        typed = typed_text.casefold()
        current_value = (self._note_manual_form_vars.get("PartyLedger", ctk.StringVar()).get() or "").strip()
        if not typed:
            filtered = self.fetched_party_ledgers[:200]
        else:
            starts = [n for n in self.fetched_party_ledgers if n.casefold().startswith(typed)]
            contains = [n for n in self.fetched_party_ledgers if typed in n.casefold() and n not in starts]
            filtered = (starts + contains)[:200]
        if typed and not filtered:
            combo.configure(values=[""])
            combo.set("")
            if "PartyLedger" in self._note_manual_form_vars:
                self._note_manual_form_vars["PartyLedger"].set("")
            if self._note_manual_party_match_label is not None:
                self._note_manual_party_match_label.configure(text=f"Search '{typed_text}': no match")
            return
        display_values = filtered if filtered else self.fetched_party_ledgers[:200]
        combo.configure(values=display_values)
        if typed and display_values and current_value not in display_values:
            combo.set(display_values[0])
            if "PartyLedger" in self._note_manual_form_vars:
                self._note_manual_form_vars["PartyLedger"].set(display_values[0])
        elif current_value:
            combo.set(current_value)
        if self._note_manual_party_match_label is not None:
            shown = len(display_values)
            total = len(self.fetched_party_ledgers)
            if typed:
                self._note_manual_party_match_label.configure(
                    text=f"Search '{typed_text}': showing {shown} of {total}")
            else:
                self._note_manual_party_match_label.configure(text=f"Showing {shown} of {total} party ledgers")

    def _note_fetch_party_ledgers_thread(self, silent=False):
        if self._ledger_fetch_running:
            return
        try:
            tally_url = self._get_tally_url()
        except ValueError as exc:
            if not silent:
                messagebox.showerror("Invalid Settings", str(exc))
            return
        selected_company = self._get_selected_company()
        self._ledger_fetch_running = True
        if self._note_manual_fetch_ledger_btn is not None:
            self._note_manual_fetch_ledger_btn.configure(state="disabled", text="...")
        if hasattr(self, '_note_create_party_fetch_btn') and self._note_create_party_fetch_btn is not None:
            self._note_create_party_fetch_btn.configure(state="disabled", text="Fetching...")

        def worker():
            result = _fetch_tally_ledgers(tally_url, timeout=15, company_name=selected_company)

            def done():
                self._ledger_fetch_running = False
                if self._note_manual_fetch_ledger_btn is not None:
                    self._note_manual_fetch_ledger_btn.configure(state="normal", text="Fetch")
                if hasattr(self, '_note_create_party_fetch_btn') and self._note_create_party_fetch_btn is not None:
                    self._note_create_party_fetch_btn.configure(state="normal", text="Fetch Party Ledgers")
                if result.get("success"):
                    party_ledgers = result.get("party_ledgers") or result.get("ledgers") or []
                    cleaned = []
                    seen = set()
                    for name in party_ledgers:
                        norm = _normalize_ledger_name(name)
                        if norm and _ledger_key(norm) not in seen:
                            seen.add(_ledger_key(norm))
                            cleaned.append(norm)
                    self.fetched_party_ledgers = sorted(cleaned, key=lambda x: _ledger_key(x))
                    if self._note_manual_party_ledger_combo is not None:
                        self._note_manual_party_ledger_combo.configure(
                            values=self.fetched_party_ledgers[:200] if self.fetched_party_ledgers else [""])
                    self._on_note_party_search_change()
                    self.status_var.set(f"Fetched {len(self.fetched_party_ledgers)} party ledger(s) from Tally")
                else:
                    err = str(result.get("error") or "Unknown error")
                    self.status_var.set("Party ledger fetch failed")
                    if not silent:
                        messagebox.showwarning("Party Ledger Fetch Failed",
                                               f"Could not fetch ledgers from Tally.\n\n{err}")
            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _note_manual_row_from_form(self):
        row = {}
        for header in self.NOTE_TEMPLATE_HEADERS:
            row[header] = self._note_manual_form_vars.get(header, ctk.StringVar()).get().strip()
        return row

    def _note_validate_manual_row(self, row):
        party = str(row.get("PartyLedger", "")).strip()
        particular = str(row.get("Particular", "")).strip()
        if not party:
            messagebox.showwarning("Missing Field", "Party Ledger is required.")
            return None
        if not particular:
            messagebox.showwarning("Missing Field", "Particular is required.")
            return None
        try:
            taxable = float(str(row.get("TaxableValue") or "0").strip())
        except ValueError:
            messagebox.showwarning("Invalid Value", "Taxable Value must be numeric.")
            return None
        if taxable <= 0:
            messagebox.showwarning("Invalid Value", "Taxable Value must be greater than zero.")
            return None
        if not str(row.get("Date") or "").strip():
            row["Date"] = datetime.today().strftime("%d-%m-%y")
        for key in ["CGSTRate", "SGSTRate", "IGSTRate"]:
            if not str(row.get(key) or "").strip():
                row[key] = "0"
        return row

    def _note_set_manual_edit_mode(self, index=None):
        self._note_manual_editing_index = index
        if self._note_manual_update_btn is not None:
            self._note_manual_update_btn.configure(state="normal" if index is not None else "disabled")

    def _note_selected_manual_index(self):
        if self._note_manual_tree is None:
            return None
        selected = list(self._note_manual_tree.selection())
        if not selected:
            return None
        try:
            return int(selected[0])
        except ValueError:
            return None

    def _note_refresh_manual_tree(self, focus_index=None):
        if self._note_manual_tree is None:
            return
        self._note_populate_tree(self._note_manual_tree, self.NOTE_TEMPLATE_HEADERS,
                                  self._note_manual_rows, limit=500)
        if focus_index is not None:
            iid = str(focus_index)
            if iid in self._note_manual_tree.get_children():
                self._note_manual_tree.selection_set(iid)
                self._note_manual_tree.focus(iid)
                self._note_manual_tree.see(iid)
        if self._note_manual_info_label is not None:
            self._note_manual_info_label.configure(text=f"Manual entries: {len(self._note_manual_rows)}")

    def _note_add_manual_entry(self):
        row = self._note_manual_row_from_form()
        validated = self._note_validate_manual_row(row)
        if validated is None:
            return
        if not str(validated.get("VoucherNo") or "").strip():
            validated["VoucherNo"] = str(len(self._note_manual_rows) + 1)
        self._note_manual_rows.append(validated)
        self._note_refresh_manual_tree(focus_index=len(self._note_manual_rows) - 1)
        self._note_set_manual_edit_mode(None)
        self.status_var.set(f"Note manual entry added. Total: {len(self._note_manual_rows)}")

    def _note_edit_selected_manual(self):
        idx = self._note_selected_manual_index()
        if idx is None:
            messagebox.showinfo("Edit Entry", "Select one row in the table to edit.")
            return
        if idx < 0 or idx >= len(self._note_manual_rows):
            return
        row = self._note_manual_rows[idx]
        for header in self.NOTE_TEMPLATE_HEADERS:
            value = _row_get(row, header, "")
            if header in self._note_manual_form_vars:
                self._note_manual_form_vars[header].set("" if value is None else str(value))
        self._note_set_manual_edit_mode(idx)
        self.status_var.set(f"Editing note entry #{idx + 1}. Update Entry to save changes.")

    def _note_update_manual_entry(self):
        idx = self._note_manual_editing_index
        if idx is None:
            messagebox.showinfo("Update Entry", "Select and edit a row first.")
            return
        if idx < 0 or idx >= len(self._note_manual_rows):
            self._note_set_manual_edit_mode(None)
            messagebox.showwarning("Update Entry", "Selected row is no longer available.")
            return
        row = self._note_manual_row_from_form()
        validated = self._note_validate_manual_row(row)
        if validated is None:
            return
        if not str(validated.get("VoucherNo") or "").strip():
            validated["VoucherNo"] = str(idx + 1)
        self._note_manual_rows[idx] = validated
        self._note_refresh_manual_tree(focus_index=idx)
        self._note_set_manual_edit_mode(None)
        self.status_var.set(f"Note manual entry #{idx + 1} updated.")

    def _note_clear_manual_form(self):
        keep_date = datetime.today().strftime("%d-%m-%y")
        for key, var in self._note_manual_form_vars.items():
            if key == "Date":
                var.set(keep_date)
            elif key in {"CGSTRate", "SGSTRate", "IGSTRate"}:
                var.set("0")
            else:
                var.set("")
        self._note_set_manual_edit_mode(None)
        self._note_manual_party_search_var.set("")
        self._on_note_party_search_change()

    def _note_remove_selected_manual(self):
        if self._note_manual_tree is None:
            return
        selected = list(self._note_manual_tree.selection())
        if not selected:
            messagebox.showinfo("Remove Entry", "Select at least one row to remove.")
            return
        indexes = []
        for iid in selected:
            try:
                indexes.append(int(iid))
            except ValueError:
                continue
        for idx in sorted(indexes, reverse=True):
            if 0 <= idx < len(self._note_manual_rows):
                self._note_manual_rows.pop(idx)
        self._note_refresh_manual_tree()
        self._note_set_manual_edit_mode(None)
        self.status_var.set(f"Removed. Remaining: {len(self._note_manual_rows)}")

    def _note_clear_all_manual(self):
        if not self._note_manual_rows:
            return
        if not messagebox.askyesno("Clear All", "Remove all note manual entries?"):
            return
        self._note_manual_rows = []
        self._note_refresh_manual_tree()
        self._note_set_manual_edit_mode(None)
        self.status_var.set("All note manual entries cleared.")

    def _note_clear_create_party_form(self):
        for attr in ['_note_create_party_name_entry', '_note_create_party_mailing_entry',
                     '_note_create_party_gstin_entry', '_note_create_party_state_entry',
                     '_note_create_party_pincode_entry', '_note_create_party_address1_entry',
                     '_note_create_party_address2_entry']:
            w = getattr(self, attr, None)
            if w is not None:
                w.delete(0, "end")
        w = getattr(self, '_note_create_party_country_entry', None)
        if w is not None:
            w.delete(0, "end")
            w.insert(0, "India")
        if hasattr(self, '_note_create_party_parent_cb') and self._note_create_party_parent_cb:
            self._note_create_party_parent_cb.set("Sundry Debtors")
        if hasattr(self, '_note_create_party_billwise_cb') and self._note_create_party_billwise_cb:
            self._note_create_party_billwise_cb.set("Yes")
        if hasattr(self, '_note_create_party_status_var'):
            self._note_create_party_status_var.set("Ready to create party ledger")

    def _note_create_party_ledger_thread(self):
        name_entry = getattr(self, '_note_create_party_name_entry', None)
        ledger_name = _normalize_ledger_name(name_entry.get() if name_entry else "")
        if not ledger_name:
            messagebox.showwarning("Missing Field", "Ledger Name is required.")
            return
        try:
            tally_url = self._get_tally_url()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc))
            return
        selected_company = self._get_selected_company()
        parent_cb = getattr(self, '_note_create_party_parent_cb', None)
        parent_name = parent_cb.get().strip() if parent_cb else "Sundry Debtors"
        mailing_entry = getattr(self, '_note_create_party_mailing_entry', None)
        gstin_entry = getattr(self, '_note_create_party_gstin_entry', None)
        state_entry = getattr(self, '_note_create_party_state_entry', None)
        country_entry = getattr(self, '_note_create_party_country_entry', None)
        pincode_entry = getattr(self, '_note_create_party_pincode_entry', None)
        addr1_entry = getattr(self, '_note_create_party_address1_entry', None)
        addr2_entry = getattr(self, '_note_create_party_address2_entry', None)
        billwise_cb = getattr(self, '_note_create_party_billwise_cb', None)

        mailing_name = mailing_entry.get().strip() if mailing_entry else ""
        gstin = gstin_entry.get().strip() if gstin_entry else ""
        state = state_entry.get().strip() if state_entry else ""
        country = country_entry.get().strip() if country_entry else "India"
        pincode = pincode_entry.get().strip() if pincode_entry else ""
        address1 = addr1_entry.get().strip() if addr1_entry else ""
        address2 = addr2_entry.get().strip() if addr2_entry else ""
        billwise_raw = billwise_cb.get().strip().upper() if billwise_cb else "YES"
        billwise_on = billwise_raw in {"YES", "Y", "TRUE", "1"}

        status_var = getattr(self, '_note_create_party_status_var', None)
        if status_var:
            status_var.set("Creating ledger in Tally...")
        self.status_var.set("Creating party ledger...")

        def worker():
            try:
                result = _create_tally_ledger(
                    tally_url=tally_url,
                    ledger_name=ledger_name,
                    parent_name=parent_name,
                    company_name=selected_company,
                    gstin=gstin,
                    state=state,
                    country=country,
                    pincode=pincode,
                    mailing_name=mailing_name,
                    address1=address1,
                    address2=address2,
                    billwise_on=billwise_on,
                    timeout=30,
                )
            except Exception as exc:
                result = {"success": False, "error": str(exc)}

            def done():
                if result.get("success"):
                    created = int(result.get("created", 0) or 0)
                    altered = int(result.get("altered", 0) or 0)
                    if status_var:
                        status_var.set(f"Created/updated. Created={created}, Altered={altered}")
                    self.status_var.set("Party ledger created successfully")
                    if "PartyLedger" in self._note_manual_form_vars:
                        self._note_manual_form_vars["PartyLedger"].set(ledger_name)
                    self._note_fetch_party_ledgers_thread(silent=True)
                    messagebox.showinfo("Create Party Ledger",
                                        f"Ledger created/updated successfully.\n\nName: {ledger_name}\nParent: {parent_name}")
                else:
                    err = str(result.get("error") or "Ledger creation failed in Tally.")
                    if status_var:
                        status_var.set(f"Create failed: {err}")
                    self.status_var.set("Party ledger create failed")
                    messagebox.showerror("Create Party Ledger", err)
            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    # ─── NOTE TEMPLATE DOWNLOAD ───────────────────────────────────────────────

    def _download_note_template(self):
        note_type = self._note_type_var.get() or "Credit Note"
        headers = self.NOTE_TEMPLATE_HEADERS
        sample = ["16-12-25", "1", "", "Interactive Media Pvt Ltd", "Lecture Income",
                  100000, "CGST", 9, "SGST", 9, "", 0, "Test Note"]
        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Template_{note_type.replace(' ', '_')}.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not out:
            return
        try:
            import openpyxl as _xl
            wb = _xl.Workbook()
            ws = wb.active
            ws.title = note_type
            ws.append(headers)
            ws.append(sample)
            wb.save(out)
            messagebox.showinfo("Template Saved", f"{note_type} template saved:\n{out}")
        except Exception as exc:
            messagebox.showerror("Template Error", str(exc))

    # ─── JOURNAL PANEL BUILDER ────────────────────────────────────────────────

    JNL_TEMPLATE_HEADERS = [
        "Date", "VoucherNo", "PartyLedger", "Particular", "TaxableValue",
        "CGSTLedger", "CGSTRate", "SGSTLedger", "SGSTRate",
        "IGSTLedger", "IGSTRate", "Narration", "TDSLedger", "TDSRate", "TDSAmount",
    ]

    def _build_journal_panel(self, parent):
        """Build the Journal Entry panel."""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=0)
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_rowconfigure(2, weight=0)

        # Row 0: Journal Type + header
        top_row = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"],
                                border_width=1, border_color=COLORS["border"], corner_radius=8)
        top_row.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 0))
        inner = ctk.CTkFrame(top_row, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=8)

        ctk.CTkLabel(inner, text="Journal Type:", font=("Segoe UI", 10),
                     text_color=COLORS["text_secondary"]).pack(side="left")
        for _lbl, _val in (("Purchase  (Dr Expense / Cr Party)", "purchase"),
                            ("Sale  (Dr Party / Cr Income)", "sale")):
            ctk.CTkRadioButton(
                inner, text=_lbl, variable=self._journal_type_var, value=_val,
                font=("Segoe UI", 10), text_color=COLORS["text_secondary"],
                fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"]).pack(side="left", padx=(12, 0))

        ctk.CTkButton(
            inner, text="📥  Download Template",
            fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
            text_color=COLORS["text_secondary"], width=170,
            command=self._download_journal_template).pack(side="right")

        # Row 1: Tabview
        self._jnl_source_tabs = ctk.CTkTabview(
            parent,
            fg_color="transparent",
            segmented_button_fg_color=COLORS["bg_input"],
            segmented_button_selected_color=COLORS["accent"],
            segmented_button_selected_hover_color=COLORS["accent_hover"],
            segmented_button_unselected_color=COLORS["bg_input"],
            segmented_button_unselected_hover_color=COLORS["bg_card_hover"],
        )
        self._jnl_source_tabs.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 2))

        excel_tab = self._jnl_source_tabs.add("Excel Upload")
        manual_tab = self._jnl_source_tabs.add("Manual Entry")
        self._jnl_source_tabs.set("Excel Upload")

        self._build_journal_excel_tab(excel_tab)
        self._build_journal_manual_tab(manual_tab)

        # Row 2: Action buttons
        btn_row = ctk.CTkFrame(parent, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))

        ctk.CTkButton(
            btn_row, text="💾  Save XML File", fg_color=SUCCESS, hover_color="#15803D", width=155,
            command=lambda: self._generate_journal("save")
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_row, text="🚀  Push to Tally", fg_color=ACCENT, hover_color=ACCENT_HOVER, width=165,
            command=lambda: self._generate_journal("push")
        ).pack(side="left", padx=(0, 8))

    def _build_journal_excel_tab(self, parent):
        """Excel Upload sub-tab for Journal Entry panel."""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        load_frame = ctk.CTkFrame(parent, fg_color="transparent")
        load_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        jnl_fp_var = ctk.StringVar()
        ctk.CTkEntry(load_frame, textvariable=jnl_fp_var,
                     placeholder_text="Select Journal Entry Excel (.xlsx/.xlsm/.xls)...",
                     state="readonly", fg_color=COLORS["bg_input"], border_color=COLORS["border"],
                     ).pack(side="left", padx=(0, 8), fill="x", expand=True)

        def browse_jnl_excel():
            f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
            if not f:
                return
            jnl_fp_var.set(f)
            try:
                headers, rows = read_excel(f)
                self._jnl_loaded_headers = headers
                self._jnl_loaded_rows = rows
                if self._jnl_excel_tree is not None:
                    self._jnl_excel_tree.delete(*self._jnl_excel_tree.get_children())
                    self._jnl_excel_tree["columns"] = headers
                    for col in headers:
                        self._jnl_excel_tree.heading(col, text=col)
                        self._jnl_excel_tree.column(col, width=max(80, len(col) * 9), minwidth=50)
                    for row in rows[:300]:
                        self._jnl_excel_tree.insert("", "end", values=[row.get(h, "") for h in headers])
                if self._jnl_excel_info_label is not None:
                    self._jnl_excel_info_label.configure(
                        text=f"📊 {len(rows)} row(s) — {len(headers)} column(s)")
                self.status_var.set(f"Loaded: {os.path.basename(f)}")
            except Exception as exc:
                messagebox.showerror("Load Error", str(exc))

        ctk.CTkButton(load_frame, text="Browse Excel", command=browse_jnl_excel,
                      width=120, fg_color=ACCENT, hover_color=ACCENT_HOVER).pack(side="left")

        info_frame = ctk.CTkFrame(parent, fg_color="transparent")
        info_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 4))
        self._jnl_excel_info_label = ctk.CTkLabel(info_frame, text="", font=("Segoe UI", 11),
                                                    text_color=TEXT_MUTED)
        self._jnl_excel_info_label.pack(side="left")

        tree_frame = ctk.CTkFrame(parent, fg_color=COLORS["bg_dark"],
                                   corner_radius=8, border_width=1, border_color=COLORS["border"])
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        _sy = ttk.Scrollbar(tree_frame, orient="vertical")
        _sx = ttk.Scrollbar(tree_frame, orient="horizontal")
        self._jnl_excel_tree = ttk.Treeview(tree_frame, show="headings",
                                              yscrollcommand=_sy.set, xscrollcommand=_sx.set)
        _sy.config(command=self._jnl_excel_tree.yview)
        _sx.config(command=self._jnl_excel_tree.xview)
        self._jnl_excel_tree.grid(row=0, column=0, sticky="nsew")
        _sy.grid(row=0, column=1, sticky="ns")
        _sx.grid(row=1, column=0, sticky="ew")

    def _build_journal_manual_tab(self, parent):
        """Manual Entry sub-tab for Journal Entry panel."""
        _JNL_HDRS = self.JNL_TEMPLATE_HEADERS
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.pack(fill="both", expand=True, padx=10, pady=8)
        wrapper.grid_columnconfigure(0, weight=1, uniform="jnl_manual_split")
        wrapper.grid_columnconfigure(1, weight=1, uniform="jnl_manual_split")
        wrapper.grid_rowconfigure(0, weight=1)

        left_panel = ctk.CTkFrame(wrapper, fg_color=COLORS["bg_card"],
                                   border_width=1, border_color=COLORS["border"], corner_radius=10)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left_panel.grid_columnconfigure(0, weight=1)
        left_panel.grid_rowconfigure(0, weight=1)

        form_scroll = ctk.CTkScrollableFrame(left_panel, fg_color="transparent", corner_radius=8)
        form_scroll.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 6))
        form_card = ctk.CTkFrame(form_scroll, fg_color="transparent")
        form_card.pack(fill="x", padx=2, pady=2)

        fields = [
            ("Date", "Date", "DD-MM-YY"),
            ("VoucherNo", "Voucher No", "Optional"),
            ("PartyLedger", "Party Ledger", "Required"),
            ("Particular", "Particular", "Required"),
            ("TaxableValue", "Taxable Value", "Required Amount"),
            ("CGSTLedger", "CGST Ledger", "Optional"),
            ("CGSTRate", "CGST Rate", "0"),
            ("SGSTLedger", "SGST Ledger", "Optional"),
            ("SGSTRate", "SGST Rate", "0"),
            ("IGSTLedger", "IGST Ledger", "Optional"),
            ("IGSTRate", "IGST Rate", "0"),
            ("Narration", "Narration", "Optional"),
        ]

        cols = 2
        for i, (key, label, placeholder) in enumerate(fields):
            row_idx = i // cols
            col_idx = i % cols
            field_wrap = ctk.CTkFrame(form_card, fg_color="transparent")
            field_wrap.grid(row=row_idx, column=col_idx, sticky="ew", padx=8, pady=6)
            form_card.grid_columnconfigure(col_idx, weight=1)

            if key == "PartyLedger":
                top_line = ctk.CTkFrame(field_wrap, fg_color="transparent")
                top_line.pack(fill="x")
                ctk.CTkLabel(top_line, text=label, font=("Segoe UI", 10),
                             text_color=COLORS["text_secondary"]).pack(side="left")
                self._jnl_manual_fetch_ledger_btn = ctk.CTkButton(
                    top_line, text="Fetch", width=70, height=26,
                    font=("Segoe UI", 10, "bold"),
                    fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
                    text_color=COLORS["text_secondary"],
                    command=self._jnl_fetch_party_ledgers_thread)
                self._jnl_manual_fetch_ledger_btn.pack(side="right")

                search_row = ctk.CTkFrame(field_wrap, fg_color="transparent")
                search_row.pack(fill="x", pady=(4, 2))
                search_entry = ctk.CTkEntry(
                    search_row, textvariable=self._jnl_manual_party_search_var,
                    placeholder_text="Search party ledger...", height=34,
                    fg_color=COLORS["bg_input"], border_color=COLORS["border"], font=("Segoe UI", 10))
                search_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
                search_entry.bind("<KeyRelease>", self._on_jnl_party_search_change)

                self._jnl_manual_party_search_clear_btn = ctk.CTkButton(
                    search_row, text="Clear", width=58, height=30,
                    fg_color=COLORS["bg_input"], hover_color=COLORS["bg_card_hover"],
                    text_color=COLORS["text_secondary"], font=("Segoe UI", 9, "bold"),
                    command=lambda: (self._jnl_manual_party_search_var.set(""), self._on_jnl_party_search_change()))
                self._jnl_manual_party_search_clear_btn.pack(side="right")

                var = ctk.StringVar(value="")
                combo = ctk.CTkComboBox(field_wrap, values=[""], variable=var, height=36,
                                        fg_color=COLORS["bg_input"], border_color=COLORS["border"],
                                        button_color=COLORS["accent"], button_hover_color=COLORS["accent_hover"],
                                        font=("Segoe UI", 10), state="readonly")
                combo.pack(fill="x", pady=(0, 2))
                self._jnl_manual_party_match_label = ctk.CTkLabel(
                    field_wrap, text="Type in search box after fetching ledgers",
                    font=("Segoe UI", 9), text_color=COLORS["text_muted"])
                self._jnl_manual_party_match_label.pack(anchor="w")
                self._jnl_manual_party_search_var.trace_add("write", lambda *_: self._on_jnl_party_search_change())
                self._jnl_manual_party_ledger_combo = combo
                self._jnl_manual_form_vars[key] = var
                continue

            ctk.CTkLabel(field_wrap, text=label, font=("Segoe UI", 10),
                         text_color=COLORS["text_secondary"]).pack(anchor="w")
            var = ctk.StringVar(value="")
            ctk.CTkEntry(field_wrap, textvariable=var, placeholder_text=placeholder,
                         height=36, fg_color=COLORS["bg_input"], border_color=COLORS["border"]).pack(fill="x")
            self._jnl_manual_form_vars[key] = var

        if "Date" in self._jnl_manual_form_vars:
            self._jnl_manual_form_vars["Date"].set(datetime.today().strftime("%d-%m-%y"))

        btn_row = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        for ci in range(3):
            btn_row.grid_columnconfigure(ci, weight=1)

        add_btn = ctk.CTkButton(btn_row, text="Add Entry", fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                height=28, font=("Segoe UI", 10, "bold"), command=self._jnl_add_manual_entry)
        add_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=(0, 4))

        edit_btn = ctk.CTkButton(btn_row, text="Edit Selected", fg_color="#0EA5E9", hover_color="#0284C7",
                                  text_color="#FFFFFF", height=28, font=("Segoe UI", 10, "bold"), command=self._jnl_edit_selected_manual)
        edit_btn.grid(row=0, column=1, sticky="ew", padx=3, pady=(0, 4))

        self._jnl_manual_update_btn = ctk.CTkButton(
            btn_row, text="Update Entry", fg_color="#10B981", hover_color="#059669",
            text_color="#FFFFFF", height=28, font=("Segoe UI", 10, "bold"), state="disabled", command=self._jnl_update_manual_entry)
        self._jnl_manual_update_btn.grid(row=0, column=2, sticky="ew", padx=(6, 0), pady=(0, 4))

        ctk.CTkButton(btn_row, text="Clear Form", fg_color=COLORS["bg_input"],
                      hover_color=COLORS["bg_card_hover"], text_color=COLORS["text_secondary"],
                      height=28, font=("Segoe UI", 10, "bold"), command=self._jnl_clear_manual_form).grid(row=1, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(btn_row, text="Remove Selected", fg_color=COLORS["warning"], hover_color="#B45309",
                      text_color="#FFFFFF", height=28, font=("Segoe UI", 10, "bold"), command=self._jnl_remove_selected_manual).grid(
            row=1, column=1, sticky="ew", padx=3)

        ctk.CTkButton(btn_row, text="Clear All", fg_color=COLORS["error"], hover_color="#B91C1C",
                      text_color="#FFFFFF", height=28, font=("Segoe UI", 10, "bold"), command=self._jnl_clear_all_manual).grid(
            row=1, column=2, sticky="ew", padx=(6, 0))

        # Right panel: review treeview
        right_panel = ctk.CTkFrame(wrapper, fg_color=COLORS["bg_card"],
                                    border_width=1, border_color=COLORS["border"], corner_radius=10)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(1, weight=1)

        right_header = ctk.CTkFrame(right_panel, fg_color="transparent")
        right_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        ctk.CTkLabel(right_header, text="Review (Excel Format)", font=("Segoe UI", 12, "bold"),
                     text_color=COLORS["text_primary"]).pack(side="left")
        self._jnl_manual_info_label = ctk.CTkLabel(right_header, text="Manual entries: 0",
                                                     font=("Segoe UI", 11), text_color=TEXT_MUTED)
        self._jnl_manual_info_label.pack(side="right")

        tree_frame = ctk.CTkFrame(right_panel, fg_color=COLORS["bg_dark"],
                                   corner_radius=8, border_width=1, border_color=COLORS["border"])
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        _sy = ttk.Scrollbar(tree_frame, orient="vertical")
        _sx = ttk.Scrollbar(tree_frame, orient="horizontal")
        self._jnl_manual_tree = ttk.Treeview(tree_frame, show="headings", selectmode="extended",
                                               yscrollcommand=_sy.set, xscrollcommand=_sx.set)
        _sy.config(command=self._jnl_manual_tree.yview)
        _sx.config(command=self._jnl_manual_tree.xview)
        _sy.pack(side="right", fill="y")
        _sx.pack(side="bottom", fill="x")
        self._jnl_manual_tree.pack(fill="both", expand=True)
        self._jnl_manual_tree.bind("<Double-1>", lambda e: self._jnl_edit_selected_manual())

        ctk.CTkLabel(right_panel, text="Double-click a row to edit it.",
                     font=("Segoe UI", 10), text_color=COLORS["text_muted"]).grid(
            row=2, column=0, sticky="w", padx=10, pady=(0, 8))

        self._jnl_populate_tree(self._jnl_manual_tree, _JNL_HDRS, [])

    # ─── JOURNAL MANUAL ENTRY METHODS ─────────────────────────────────────────

    def _jnl_populate_tree(self, tree, headers, rows, limit=500):
        tree.delete(*tree.get_children())
        tree["columns"] = headers
        for h in headers:
            tree.heading(h, text=h)
            tree.column(h, width=max(120, min(260, len(h) * 12)), minwidth=80)
        for idx, row in enumerate(rows[:limit]):
            values = [str(_row_get(row, h, "") or "") for h in headers]
            tree.insert("", "end", iid=str(idx), values=values)

    def _on_jnl_party_search_change(self, _event=None):
        combo = self._jnl_manual_party_ledger_combo
        if combo is None:
            return
        if not self.fetched_party_ledgers:
            combo.configure(values=[""])
            combo.set("")
            if self._jnl_manual_party_match_label is not None:
                self._jnl_manual_party_match_label.configure(text="No fetched party ledgers yet")
            return
        typed_text = (self._jnl_manual_party_search_var.get() or "").strip()
        typed = typed_text.casefold()
        current_value = (self._jnl_manual_form_vars.get("PartyLedger", ctk.StringVar()).get() or "").strip()
        if not typed:
            filtered = self.fetched_party_ledgers[:200]
        else:
            starts = [n for n in self.fetched_party_ledgers if n.casefold().startswith(typed)]
            contains = [n for n in self.fetched_party_ledgers if typed in n.casefold() and n not in starts]
            filtered = (starts + contains)[:200]
        if typed and not filtered:
            combo.configure(values=[""])
            combo.set("")
            if "PartyLedger" in self._jnl_manual_form_vars:
                self._jnl_manual_form_vars["PartyLedger"].set("")
            if self._jnl_manual_party_match_label is not None:
                self._jnl_manual_party_match_label.configure(text=f"Search '{typed_text}': no match")
            return
        display_values = filtered if filtered else self.fetched_party_ledgers[:200]
        combo.configure(values=display_values)
        if typed and display_values and current_value not in display_values:
            combo.set(display_values[0])
            if "PartyLedger" in self._jnl_manual_form_vars:
                self._jnl_manual_form_vars["PartyLedger"].set(display_values[0])
        elif current_value:
            combo.set(current_value)
        if self._jnl_manual_party_match_label is not None:
            shown = len(display_values)
            total = len(self.fetched_party_ledgers)
            if typed:
                self._jnl_manual_party_match_label.configure(
                    text=f"Search '{typed_text}': showing {shown} of {total}")
            else:
                self._jnl_manual_party_match_label.configure(text=f"Showing {shown} of {total} party ledgers")

    def _jnl_fetch_party_ledgers_thread(self, silent=False):
        if self._ledger_fetch_running:
            return
        try:
            tally_url = self._get_tally_url()
        except ValueError as exc:
            if not silent:
                messagebox.showerror("Invalid Settings", str(exc))
            return
        selected_company = self._get_selected_company()
        self._ledger_fetch_running = True
        if self._jnl_manual_fetch_ledger_btn is not None:
            self._jnl_manual_fetch_ledger_btn.configure(state="disabled", text="...")

        def worker():
            result = _fetch_tally_ledgers(tally_url, timeout=15, company_name=selected_company)

            def done():
                self._ledger_fetch_running = False
                if self._jnl_manual_fetch_ledger_btn is not None:
                    self._jnl_manual_fetch_ledger_btn.configure(state="normal", text="Fetch")
                if result.get("success"):
                    party_ledgers = result.get("party_ledgers") or result.get("ledgers") or []
                    cleaned = []
                    seen = set()
                    for name in party_ledgers:
                        norm = _normalize_ledger_name(name)
                        if norm and _ledger_key(norm) not in seen:
                            seen.add(_ledger_key(norm))
                            cleaned.append(norm)
                    self.fetched_party_ledgers = sorted(cleaned, key=lambda x: _ledger_key(x))
                    if self._jnl_manual_party_ledger_combo is not None:
                        self._jnl_manual_party_ledger_combo.configure(
                            values=self.fetched_party_ledgers[:200] if self.fetched_party_ledgers else [""])
                    self._on_jnl_party_search_change()
                    self.status_var.set(f"Fetched {len(self.fetched_party_ledgers)} party ledger(s) from Tally")
                else:
                    err = str(result.get("error") or "Unknown error")
                    self.status_var.set("Party ledger fetch failed")
                    if not silent:
                        messagebox.showwarning("Party Ledger Fetch Failed",
                                               f"Could not fetch ledgers from Tally.\n\n{err}")
            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _jnl_manual_row_from_form(self):
        row = {}
        for header in self.JNL_TEMPLATE_HEADERS:
            row[header] = self._jnl_manual_form_vars.get(header, ctk.StringVar()).get().strip()
        return row

    def _jnl_validate_manual_row(self, row):
        party = str(row.get("PartyLedger", "")).strip()
        particular = str(row.get("Particular", "")).strip()
        if not party:
            messagebox.showwarning("Missing Field", "Party Ledger is required.")
            return None
        if not particular:
            messagebox.showwarning("Missing Field", "Particular is required.")
            return None
        try:
            taxable = float(str(row.get("TaxableValue") or "0").strip())
        except ValueError:
            messagebox.showwarning("Invalid Value", "Taxable Value must be numeric.")
            return None
        if taxable <= 0:
            messagebox.showwarning("Invalid Value", "Taxable Value must be greater than zero.")
            return None
        if not str(row.get("Date") or "").strip():
            row["Date"] = datetime.today().strftime("%d-%m-%y")
        for key in ["CGSTRate", "SGSTRate", "IGSTRate"]:
            if not str(row.get(key) or "").strip():
                row[key] = "0"
        return row

    def _jnl_set_manual_edit_mode(self, index=None):
        self._jnl_manual_editing_index = index
        if self._jnl_manual_update_btn is not None:
            self._jnl_manual_update_btn.configure(state="normal" if index is not None else "disabled")

    def _jnl_selected_manual_index(self):
        if self._jnl_manual_tree is None:
            return None
        selected = list(self._jnl_manual_tree.selection())
        if not selected:
            return None
        try:
            return int(selected[0])
        except ValueError:
            return None

    def _jnl_refresh_manual_tree(self, focus_index=None):
        if self._jnl_manual_tree is None:
            return
        self._jnl_populate_tree(self._jnl_manual_tree, self.JNL_TEMPLATE_HEADERS,
                                  self._jnl_manual_rows, limit=500)
        if focus_index is not None:
            iid = str(focus_index)
            if iid in self._jnl_manual_tree.get_children():
                self._jnl_manual_tree.selection_set(iid)
                self._jnl_manual_tree.focus(iid)
                self._jnl_manual_tree.see(iid)
        if self._jnl_manual_info_label is not None:
            self._jnl_manual_info_label.configure(text=f"Manual entries: {len(self._jnl_manual_rows)}")

    def _jnl_add_manual_entry(self):
        row = self._jnl_manual_row_from_form()
        validated = self._jnl_validate_manual_row(row)
        if validated is None:
            return
        if not str(validated.get("VoucherNo") or "").strip():
            validated["VoucherNo"] = str(len(self._jnl_manual_rows) + 1)
        self._jnl_manual_rows.append(validated)
        self._jnl_refresh_manual_tree(focus_index=len(self._jnl_manual_rows) - 1)
        self._jnl_set_manual_edit_mode(None)
        self.status_var.set(f"Journal manual entry added. Total: {len(self._jnl_manual_rows)}")

    def _jnl_edit_selected_manual(self):
        idx = self._jnl_selected_manual_index()
        if idx is None:
            messagebox.showinfo("Edit Entry", "Select one row in the table to edit.")
            return
        if idx < 0 or idx >= len(self._jnl_manual_rows):
            return
        row = self._jnl_manual_rows[idx]
        for header in self.JNL_TEMPLATE_HEADERS:
            value = _row_get(row, header, "")
            if header in self._jnl_manual_form_vars:
                self._jnl_manual_form_vars[header].set("" if value is None else str(value))
        self._jnl_set_manual_edit_mode(idx)
        self.status_var.set(f"Editing journal entry #{idx + 1}. Update Entry to save changes.")

    def _jnl_update_manual_entry(self):
        idx = self._jnl_manual_editing_index
        if idx is None:
            messagebox.showinfo("Update Entry", "Select and edit a row first.")
            return
        if idx < 0 or idx >= len(self._jnl_manual_rows):
            self._jnl_set_manual_edit_mode(None)
            messagebox.showwarning("Update Entry", "Selected row is no longer available.")
            return
        row = self._jnl_manual_row_from_form()
        validated = self._jnl_validate_manual_row(row)
        if validated is None:
            return
        if not str(validated.get("VoucherNo") or "").strip():
            validated["VoucherNo"] = str(idx + 1)
        self._jnl_manual_rows[idx] = validated
        self._jnl_refresh_manual_tree(focus_index=idx)
        self._jnl_set_manual_edit_mode(None)
        self.status_var.set(f"Journal manual entry #{idx + 1} updated.")

    def _jnl_clear_manual_form(self):
        keep_date = datetime.today().strftime("%d-%m-%y")
        for key, var in self._jnl_manual_form_vars.items():
            if key == "Date":
                var.set(keep_date)
            elif key in {"CGSTRate", "SGSTRate", "IGSTRate"}:
                var.set("0")
            else:
                var.set("")
        self._jnl_set_manual_edit_mode(None)
        self._jnl_manual_party_search_var.set("")
        self._on_jnl_party_search_change()

    def _jnl_remove_selected_manual(self):
        if self._jnl_manual_tree is None:
            return
        selected = list(self._jnl_manual_tree.selection())
        if not selected:
            messagebox.showinfo("Remove Entry", "Select at least one row to remove.")
            return
        indexes = []
        for iid in selected:
            try:
                indexes.append(int(iid))
            except ValueError:
                continue
        for idx in sorted(indexes, reverse=True):
            if 0 <= idx < len(self._jnl_manual_rows):
                self._jnl_manual_rows.pop(idx)
        self._jnl_refresh_manual_tree()
        self._jnl_set_manual_edit_mode(None)
        self.status_var.set(f"Removed. Remaining: {len(self._jnl_manual_rows)}")

    def _jnl_clear_all_manual(self):
        if not self._jnl_manual_rows:
            return
        if not messagebox.askyesno("Clear All", "Remove all journal manual entries?"):
            return
        self._jnl_manual_rows = []
        self._jnl_refresh_manual_tree()
        self._jnl_set_manual_edit_mode(None)
        self.status_var.set("All journal manual entries cleared.")

    def _download_journal_template(self):
        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Template_Journal Voucher.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not out:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append(self.JNL_TEMPLATE_HEADERS)
            ws.append(["16-12-25", "1", "Interactive Media Pvt Ltd", "Professional Fee Expense",
                        100000, "", 0, "", 0, "IGST", 18, "This is Testing Voucher",
                        "TDS Payable on Professional", 10, ""])
            wb.save(out)
            messagebox.showinfo("Template Saved", f"Journal template saved to:\n{out}")
        except Exception as exc:
            messagebox.showerror("Template Error", str(exc))

    # ─── JOURNAL GENERATION ───────────────────────────────────────────────────

    def _generate_journal(self, action: str):
        if self._push_running:
            self.status_var.set("Push already in progress...")
            return

        if hasattr(self, '_jnl_source_tabs') and self._jnl_source_tabs is not None:
            active_tab = self._jnl_source_tabs.get()
            if active_tab == "Manual Entry":
                rows = getattr(self, '_jnl_manual_rows', [])
                source_label = "Manual Entry"
            else:
                rows = getattr(self, '_jnl_loaded_rows', [])
                source_label = "Excel Upload"
        else:
            rows = getattr(self, '_jnl_loaded_rows', [])
            source_label = "Excel Upload"

        if not rows:
            messagebox.showwarning("No Data", f"No rows available in {source_label}.")
            return

        company = self._get_selected_company()
        if action == "push" and not company and len(getattr(self, "fetched_companies", [])) > 1:
            messagebox.showwarning("Select Company", "Please select a target company before pushing.")
            return

        try:
            date_mode, custom_tally_date = self._get_voucher_date_selection()
            journal_type = self._journal_type_var.get() if self._journal_type_var else "purchase"
            _cmp_regs = list(self.company_gst_registrations or [])
            if not _cmp_regs:
                try:
                    _r = _fetch_cmp_gst_regs_for_journal(self._get_tally_url(), company=company, timeout=10)
                    if _r.get("success"):
                        _cmp_regs = _r["registrations"]
                        self.company_gst_registrations = _cmp_regs
                except Exception:
                    pass

            xml_payload, voucher_count = generate_journal_xml(
                rows,
                company=company,
                date_mode=date_mode,
                custom_tally_date=custom_tally_date,
                journal_type=journal_type,
                company_gst_registrations=_cmp_regs,
            )
            if voucher_count <= 0:
                messagebox.showwarning("No Vouchers", "No valid rows found (TaxableValue must be greater than zero).")
                return

            if action == "save":
                out = filedialog.asksaveasfilename(
                    defaultextension=".xml",
                    initialfile="Journal.xml",
                    filetypes=[("XML", "*.xml")],
                )
                if not out:
                    return
                with open(out, "w", encoding="utf-8") as f:
                    f.write(xml_payload)
                self.status_var.set(f"Journal XML saved: {os.path.basename(out)} ({voucher_count} voucher(s))")
                messagebox.showinfo("Saved", f"Journal XML saved successfully.\n{out}")
                return

            host = (self.tally_host_var.get() or "localhost").strip()
            port_text = (self.tally_port_var.get() or "9000").strip()
            if not port_text.isdigit():
                raise ValueError("Port must be numeric.")
            port = int(port_text)

            self._set_push_loading_state(True, f"Pushing {voucher_count} Journal voucher(s) from {source_label}...")
            self.status_var.set("Pushing to Tally...")

            def worker():
                try:
                    resp = push_to_tally(xml_payload, host=host, port=port)
                    parsed = _parse_tally_response_details(resp)
                    result = {"ok": True, "parsed": parsed}
                except Exception as exc:
                    result = {"ok": False, "error": str(exc)}

                def done():
                    self._set_push_loading_state(False)
                    if not result.get("ok"):
                        err = str(result.get("error", "Unknown error"))
                        self.status_var.set(f"Push failed: {err}")
                        messagebox.showerror("Push Failed", err)
                        return
                    parsed = result["parsed"]
                    created = parsed.get("created", 0)
                    altered = parsed.get("altered", 0)
                    errors = parsed.get("errors", 0)
                    line_errors = parsed.get("line_errors", [])
                    summary = (f"Created: {created}\nAltered: {altered}\nErrors: {errors}")
                    if parsed.get("success"):
                        self.status_var.set(f"Journal push successful: Created {created}, Altered {altered}")
                        messagebox.showinfo("Push Successful", summary)
                    else:
                        if line_errors:
                            summary += "\n\nLine Errors:\n- " + "\n- ".join(line_errors[:8])
                        self.status_var.set("Journal push completed with errors.")
                        messagebox.showwarning("Push Completed With Errors", summary)
                self.after(0, done)

            threading.Thread(target=worker, daemon=True).start()

        except ValueError as exc:
            messagebox.showerror("Validation Error", str(exc))
        except Exception as exc:
            messagebox.showerror("Error", str(exc))


# ═══════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = TallySalesApp()
    app.mainloop()
