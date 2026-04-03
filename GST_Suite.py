r"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         GST & INCOME TAX AUTOMATION SUITE  —  Unified Launcher               ║
║                                                                              ║
║  Landing page → select category → tabbed tool view → ← Home to go back.      ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import importlib
import importlib.util
import traceback
import tkinter as _tk

_SPLASH = None  # no splash screen

import customtkinter as ctk
import json
import uuid
import webbrowser
import urllib.request
import urllib.error
import threading
import shutil
import tempfile
import subprocess
from datetime import datetime

# ── Version & Update Manifest ─────────────────────────────────────────────────
VERSION            = "1.0.6"
# !! REPLACE 'YOURNAME' and 'YOURREPO' with your actual GitHub username and
#    the public releases repo you created (e.g. gst-suite-releases).
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/thestudycafesagar/gst-suite-releases/main/latest.json"

_MISSING_MODULE_PACKAGES = {
    "fitz": "PyMuPDF",
    "win32com": "pywin32",
    "pythoncom": "pywin32",
    "pywintypes": "pywin32",
    "win32api": "pywin32",
    "win32con": "pywin32",
    "win32gui": "pywin32",
}

try:
    from PIL import Image as _PILImage
except ImportError:
    _PILImage = None

# ══════════════════════════════════════════════════════════════════════════════
#  LICENSING / AUTH
# ══════════════════════════════════════════════════════════════════════════════
def _get_app_data_dir():
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller EXE — use %APPDATA%\GSTSuite
        base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "GSTSuite")
    else:
        # Running as .py script — save next to the script
        base = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(base, exist_ok=True)
    return base

def _suite_debug_log(msg: str):
    """Append loader diagnostics to APPDATA log; never fail the UI path."""
    try:
        p = os.path.join(_get_app_data_dir(), "suite_debug.log")
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(p, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def _cancel_all_after_callbacks(widget):
    """Cancel all pending Tcl 'after' callbacks for a widget/app."""
    try:
        pending = widget.tk.call("after", "info")
    except Exception:
        return

    if isinstance(pending, str):
        pending = [pending] if pending else []

    for job_id in pending:
        try:
            widget.after_cancel(job_id)
        except Exception:
            pass


def _missing_package_for_module(module_name: str):
    root = (module_name or "").split(".", 1)[0].strip()
    if not root:
        return None
    return _MISSING_MODULE_PACKAGES.get(root)

_AUTH_CONFIG = os.path.join(_get_app_data_dir(), "auth_config.json")
API_BASE_URL  = "https://studycafe-tools-api-bthzgsfvfggjd6gt.centralindia-01.azurewebsites.net"
REGISTER_URL  = f"{API_BASE_URL}/register"

def _get_hardware_id():
    mac = uuid.getnode()
    return ':'.join(['{:02x}'.format((mac >> i) & 0xff) for i in range(0, 48, 8)][::-1])

def _save_auth(email, password):
    with open(_AUTH_CONFIG, 'w') as f:
        json.dump({"email": email, "password": password}, f)

def _clear_auth():
    try:
        os.remove(_AUTH_CONFIG)
    except Exception:
        pass

def _call_api(endpoint, payload):
    data = json.dumps(payload).encode()
    req  = urllib.request.Request(
        f"{API_BASE_URL}{endpoint}",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read())

# ── Appearance (must run before any tool import) ─────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

_RealCTk           = ctk.CTk
_RealSetAppearance = ctk.set_appearance_mode

# ── Base paths ────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    # Everything (assets + tool sub-folders) is bundled inside the EXE
    _ASSETS_BASE = sys._MEIPASS
    _BASE        = sys._MEIPASS
else:
    _ASSETS_BASE = os.path.dirname(os.path.abspath(__file__))
    _BASE        = _ASSETS_BASE

_GST_BASE   = os.path.join(_BASE, "GST")
_IT_BASE    = os.path.join(_BASE, "Income Tax")
_PDF_BASE   = os.path.join(_BASE, "PDF_Utilities")
_BANK_BASE  = os.path.join(_BASE, "Bank Statement To Excel")
_EMAIL_BASE = os.path.join(_BASE, "Email-Tools")
_RECO_BASE  = os.path.join(_BASE, "GST_RECO")


# ══════════════════════════════════════════════════════════════════════════════
#  _EmbeddedFrame  – swaps in for ctk.CTk during tool import
# ══════════════════════════════════════════════════════════════════════════════
class _EmbeddedFrame(ctk.CTkFrame):
    _host: ctk.CTkFrame = None

    def __init__(self, *args, **kwargs):
        kwargs.pop("fg_color", None)
        super().__init__(_EmbeddedFrame._host)
        self.grid(row=0, column=0, sticky="nsew")

    def title(self, t=None):       return ""
    def geometry(self, g=None):    return ""
    def resizable(self, *a, **k):  pass
    def mainloop(self):            pass
    def lift(self):                pass
    def iconbitmap(self, *a, **k): pass
    def wm_title(self, t=None):    return ""
    def protocol(self, *a, **k):   pass
    def attributes(self, *a, **k): pass


# ══════════════════════════════════════════════════════════════════════════════
#  _EmbeddedTkFrame  – swaps in for tk.Tk during plain-tkinter tool import
# ══════════════════════════════════════════════════════════════════════════════
class _EmbeddedTkFrame(_tk.Frame):
    """Replaces tk.Tk so a plain-tkinter app renders inside a tab frame."""
    _host: ctk.CTkFrame = None

    def __init__(self, *args, **kwargs):
        _tk.Frame.__init__(self, _EmbeddedTkFrame._host,
                           bg="#1e1e2e")
        self.pack(fill="both", expand=True)

    def title(self, t=None):           return ""
    def geometry(self, g=None):        return ""
    def resizable(self, *a, **k):      pass
    def mainloop(self):                pass
    def lift(self):                    pass
    def iconbitmap(self, *a, **k):     pass
    def wm_title(self, t=None):        return ""
    def protocol(self, *a, **k):       pass
    def attributes(self, *a, **k):     pass
    def minsize(self, *a, **k):        pass
    def maxsize(self, *a, **k):        pass
    def state(self, *a, **k):          return "normal"
    def withdraw(self):                pass
    def deiconify(self):               pass
    def option_add(self, *a, **k):     pass
    def report_callback_exception(self, *a, **k): pass
    def configure(self, **kwargs):
        safe = {k: v for k, v in kwargs.items()
                if k in ("bg", "background", "relief", "bd", "borderwidth",
                         "cursor", "height", "width")}
        if safe:
            try: _tk.Frame.configure(self, **safe)
            except Exception: pass
    config = configure


# ══════════════════════════════════════════════════════════════════════════════
#  DESIGN TOKENS
# ══════════════════════════════════════════════════════════════════════════════
_C = {
    # Page / surface
    "banner_bg":   ("#1e293b", "#060c18"),
    "status_bg":   ("#1e293b", "#060c18"),
    "surface":     ("#ffffff", "#1e293b"),
    "surface2":    ("#f8fafc", "#111827"),
    "border":      ("#e2e8f0", "#334155"),
    # Text
    "text_hi":     ("#0f172a", "#f1f5f9"),
    "text_mid":    ("#475569", "#94a3b8"),
    "text_lo":     ("#94a3b8", "#475569"),
    # Brand
    "primary":     ("#4f46e5", "#6366f1"),
    "primary_hov": ("#4338ca", "#4f46e5"),
    # GST category
    "gst_acc":     ("#4f46e5", "#818cf8"),
    "gst_bg":      ("#ede9fe", "#12103a"),
    "gst_hover":   ("#ddd6fe", "#1c1856"),
    # Income Tax category
    "it_acc":      ("#059669", "#34d399"),
    "it_bg":       ("#d1fae5", "#062818"),
    "it_hover":    ("#a7f3d0", "#0a3d22"),
    # PDF Tools category
    "pdf_acc":     ("#7c3aed", "#a78bfa"),
    "pdf_bg":      ("#ede9fe", "#12103a"),
    "pdf_hover":   ("#ddd6fe", "#1c1856"),
    # Bank Statement category
    "bank_acc":    ("#0891b2", "#22d3ee"),
    "bank_bg":     ("#cffafe", "#0c1a1f"),
    "bank_hover":  ("#a5f3fc", "#132026"),
    # Email Tools category
    "email_acc":   ("#d97706", "#fbbf24"),
    "email_bg":    ("#fef3c7", "#1c1000"),
    "email_hover": ("#fde68a", "#261800"),
    # GST Reconciliation category
    "reco_acc":    ("#0f766e", "#2dd4bf"),
    "reco_bg":     ("#ccfbf1", "#042f2e"),
    "reco_hover":  ("#99f6e4", "#064e3b"),
}

# Per-tool accent colours (light, dark)
_GST_ACCENTS = [
    ("#4f46e5", "#818cf8"),   # GSTR-2B      Indigo
    ("#7c3aed", "#a78bfa"),   # GSTR-3B      Violet
    ("#0891b2", "#22d3ee"),   # 3B→Excel     Cyan
    ("#059669", "#34d399"),   # GST Verifier Emerald
    ("#d97706", "#fbbf24"),   # Challan      Amber
    ("#0284c7", "#38bdf8"),   # R1 JSON      Sky
    ("#db2777", "#f472b6"),   # JSON→Excel   Pink
    ("#dc2626", "#f87171"),   # R1 PDF       Red
    ("#0f766e", "#2dd4bf"),   # IMS          Teal
    ("#65a30d", "#a3e635"),   # GSTR1 Cons.  Lime
]
_IT_ACCENTS = [
    ("#0891b2", "#22d3ee"),   # 26/AIS/TIS     Cyan
    ("#059669", "#34d399"),   # IT Challan     Emerald
    ("#7c3aed", "#a78bfa"),   # ITR Bot        Violet
    ("#dc2626", "#f87171"),   # Demand Checker Red
    ("#0284c7", "#38bdf8"),   # Refund Checker Sky
]
_PDF_ACCENTS = [
    ("#7c3aed", "#a78bfa"),   # PDF Utilities  Violet
]
_BANK_ACCENTS = [
    ("#0891b2", "#22d3ee"),   # Bank → Excel   Cyan
]
_EMAIL_ACCENTS = [
    ("#d97706", "#fbbf24"),   # GST Return     Amber
    ("#0284c7", "#38bdf8"),   # Invoice Sender Sky
    ("#059669", "#34d399"),   # Payment Remind Emerald
]


# ══════════════════════════════════════════════════════════════════════════════
#  TOOL REGISTRY
# ══════════════════════════════════════════════════════════════════════════════
GST_TOOLS = [
    {"key": "GSTR2B",       "tab": "📥  GSTR-2B",      "module": os.path.join(_GST_BASE, "GST 2B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-2B returns via automated browser."},
    {"key": "GSTR3B",       "tab": "📥  GSTR-3B",      "module": os.path.join(_GST_BASE, "GST 3B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-3B returns via automated browser."},
    {"key": "GSTR3B_Excel", "tab": "📊  3B → Excel",   "module": os.path.join(_GST_BASE, "GST 3B to Excel",       "main.py"),        "class": "GSTR3BConverterPro", "desc": "Convert GSTR-3B PDF files to formatted Excel sheets."},
    {"key": "GST_Verifier", "tab": "🤖  GST Verifier", "module": os.path.join(_GST_BASE, "GST Bot",               "gst_pro_app.py"), "class": "GSTApp",             "desc": "Verify bulk GSTINs and extract filing history & details."},
    {"key": "GST_Challan",  "tab": "💰  Challan",      "module": os.path.join(_GST_BASE, "GST Challan Downloader","main.py"),        "class": "App",                "desc": "Download GST Challan PDFs in bulk (Monthly / Quarterly)."},
    {"key": "R1_JSON",      "tab": "📑  R1 JSON",      "module": os.path.join(_GST_BASE, "GST R1 Downloader",     "mai.py"),         "class": "App",                "desc": "Request or download GSTR-1 JSON files for multiple users."},
    {"key": "JSON_Excel",   "tab": "📊  JSON → Excel", "module": os.path.join(_GST_BASE, "JSON to Excel",          "main.py"),        "class": "App",                "desc": "Convert GSTR-1 JSON exports to multi-sheet Excel reports."},
    {"key": "R1_PDF",       "tab": "🖨️  R1 PDF",       "module": os.path.join(_GST_BASE, "R1 PDF Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-1 PDF filed    returns from the GST portal."},
    {"key": "IMS",          "tab": "📲  IMS",          "module": os.path.join(_GST_BASE, "IMS Downloader",        "main.py"),        "class": "App",                "desc": "Download IMS (Invoice Management System) data in bulk from the GST portal."},
    {"key": "GSTR1_Cons",   "tab": "📋  GSTR1 Cons.", "module": os.path.join(_GST_BASE, "GSTR1_Consolidation",  "gst_consolidation.py"), "class": "ChallExtractorApp", "tk": True, "desc": "Consolidate multiple GSTR-1 files into a single unified Excel report."},
]

IT_TOOLS = [
    {"key": "IT_26AS",         "tab": "📄  26/AIS/TIS",          "module": os.path.join(_IT_BASE, "26 AS Downlaoder",  "main.py"),                  "class": "App",              "desc": "Download 26AS / AIS / TIS reports in bulk."},
    {"key": "IT_Challan",      "tab": "💰  Challan Downloader", "module": os.path.join(_IT_BASE, "Challan Downloader","main.py"),                  "class": "App",              "desc": "Download Income Tax Challan PDFs in bulk."},
    {"key": "ITR_Bot",         "tab": "🤖  ITR Bot",            "module": os.path.join(_IT_BASE, "ITR - Bot",         "GUI_based_app.py"),         "class": "App",              "desc": "Automate ITR filing workflows with the ITR bot."},
    {"key": "Demand_Checker",  "tab": "🔍  Demand Checker",     "module": os.path.join(_IT_BASE, "Challan Downloader","demand_checker_app.py"),    "class": "DemandCheckerApp", "desc": "Check pending worklist and outstanding demands in bulk from the Income Tax portal."},
    {"key": "Refund_Checker",  "tab": "📊  Refund Checker",     "module": os.path.join(_IT_BASE, "26 AS Downlaoder",  "refund_checker_app.py"),    "class": "RefundCheckerApp", "desc": "Extract filed return data and generate refund status reports in bulk."},
]

PDF_TOOLS = [
    {"key": "PDF_Merge",    "tab": "⊕  Merge",    "module": os.path.join(_PDF_BASE, "main.py"), "class": "MergeApp",    "tk": True, "desc": "Merge multiple PDF files into one high-quality document."},
    {"key": "PDF_Split",    "tab": "✂  Split",    "module": os.path.join(_PDF_BASE, "main.py"), "class": "SplitApp",    "tk": True, "desc": "Split PDF files into smaller parts by page range or every N pages."},
    {"key": "PDF_Extract",  "tab": "⊙  Extract",  "module": os.path.join(_PDF_BASE, "main.py"), "class": "ExtractApp",  "tk": True, "desc": "Extract specific pages from a PDF document to a new file."},
    {"key": "PDF_Compress", "tab": "⊜  Compress", "module": os.path.join(_PDF_BASE, "main.py"), "class": "CompressApp", "tk": True, "desc": "Reduce PDF file size while maintaining visual quality."},
    {"key": "PDF_Redact",   "tab": "⬛  Redact",   "module": os.path.join(_PDF_BASE, "main.py"), "class": "RedactApp",   "tk": True, "desc": "Securely black out sensitive information and text from PDF documents."},
]

BANK_TOOLS = [
    {"key": "Bank_Excel", "tab": "🏦  Bank → Excel", "module": os.path.join(_BANK_BASE, "bank_to_excel.py"), "class": "App", "tk": True, "desc": "Convert bank statement PDFs to formatted Excel sheets. Supports HDFC, ICICI, SBI, Axis, Kotak, IDFC, BOI, Yes, UCO & Equitas."},
]

EMAIL_TOOLS = [
    {"key": "Email_GST_Request", "tab": "📋  GST Return Request", "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "GSTReturnMailApp",      "tk": True, "desc": "Send bulk GST return data request emails via Outlook. Auto-fills month, return type, deadlines and contact details."},
    {"key": "Email_Invoice",     "tab": "🧾  Invoice Sender",     "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "InvoiceSenderMailApp",   "tk": True, "desc": "Dispatch personalised invoices to clients in bulk via Outlook. Supports per-row service, period, amount & PDF attachments."},
    {"key": "Email_Payment",     "tab": "💰  Payment Reminder",   "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "PaymentReminderMailApp", "tk": True, "desc": "Send outstanding payment reminder emails in bulk via Outlook. Includes interest clause, deadline and per-client amounts."},
]

RECO_TOOLS = [
    {"key": "GST_Reco", "tab": "🔄  GST Reconciliation", "module": os.path.join(_RECO_BASE, "mainpy-reco-speqtra.py"), "class": "App", "tk": False, "desc": "Reconcile GSTR-2B portal data against Tally/books. Matches invoices, highlights mismatches and exports a detailed Excel report."},
]
_RECO_ACCENTS = [
    ("#0f766e", "#2dd4bf"),
]


# ══════════════════════════════════════════════════════════════════════════════
#  TOOL LOADER
# ══════════════════════════════════════════════════════════════════════════════
def _load_tool(tab_frame: ctk.CTkFrame, module_path: str, class_name: str):
    tab_frame.grid_rowconfigure(0, weight=1)
    tab_frame.grid_columnconfigure(0, weight=1)

    _EmbeddedFrame._host = tab_frame
    ctk.CTk = _EmbeddedFrame
    _saved_mode = ctk.get_appearance_mode()
    ctk.set_appearance_mode = lambda *a, **k: None

    tool_dir = os.path.dirname(os.path.abspath(module_path))
    try:
        if tool_dir not in sys.path:
            sys.path.insert(0, tool_dir)
        uid  = f"_suite_{class_name}_{os.path.basename(tool_dir)}"
        spec = importlib.util.spec_from_file_location(uid, module_path)
        mod  = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        AppClass = getattr(mod, class_name)
        return AppClass()

    except Exception:
        err = traceback.format_exc()
        ef = ctk.CTkFrame(tab_frame, fg_color="transparent")
        ef.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)

        hdr = ctk.CTkFrame(ef, fg_color=("#fef2f2", "#1c0a0a"), corner_radius=10,
                           border_color=("#fca5a5", "#7f1d1d"), border_width=1)
        hdr.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(hdr, text="  ⚠️   Failed to load tool — traceback below",
                     font=("Segoe UI", 13, "bold"),
                     text_color=("#ef4444", "#f87171")).pack(anchor="w", padx=18, pady=12)

        tb = ctk.CTkTextbox(ef, font=("Consolas", 11),
                            text_color=("#ef4444", "#f87171"),
                            fg_color=("#1a1a2e", "#0d0d1a"),
                            corner_radius=10, border_width=1,
                            border_color=("#7f1d1d", "#450a0a"), height=320)
        tb.pack(fill="both", expand=True)
        tb.insert("1.0", err)
        tb.configure(state="disabled")
        return None

    finally:
        ctk.CTk = _RealCTk
        ctk.set_appearance_mode = _RealSetAppearance
        _RealSetAppearance(_saved_mode)


def _load_tk_tool(tab_frame: ctk.CTkFrame, module_path: str, class_name: str):
    """Like _load_tool but swaps tk.Tk → _EmbeddedTkFrame for plain-tkinter apps."""
    tab_frame.grid_rowconfigure(0, weight=1)
    tab_frame.grid_columnconfigure(0, weight=1)

    _EmbeddedTkFrame._host = tab_frame
    _real_Tk = _tk.Tk
    _tk.Tk = _EmbeddedTkFrame

    tool_dir = os.path.dirname(os.path.abspath(module_path))
    try:
        if tool_dir not in sys.path:
            sys.path.insert(0, tool_dir)
        uid  = f"_suite_{class_name}_{os.path.basename(tool_dir)}"
        spec = importlib.util.spec_from_file_location(uid, module_path)
        mod  = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        AppClass = getattr(mod, class_name)
        return AppClass()

    except Exception:
        err = traceback.format_exc()
        ef = ctk.CTkFrame(tab_frame, fg_color="transparent")
        ef.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)

        hdr = ctk.CTkFrame(ef, fg_color=("#fef2f2", "#1c0a0a"), corner_radius=10,
                           border_color=("#fca5a5", "#7f1d1d"), border_width=1)
        hdr.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(hdr, text="  ⚠️   Failed to load tool — traceback below",
                     font=("Segoe UI", 13, "bold"),
                     text_color=("#ef4444", "#f87171")).pack(anchor="w", padx=18, pady=12)

        tb = ctk.CTkTextbox(ef, font=("Consolas", 11),
                            text_color=("#ef4444", "#f87171"),
                            fg_color=("#1a1a2e", "#0d0d1a"),
                            corner_radius=10, border_width=1,
                            border_color=("#7f1d1d", "#450a0a"), height=320)
        tb.pack(fill="both", expand=True)
        tb.insert("1.0", err)
        tb.configure(state="disabled")
        return None

    finally:
        _tk.Tk = _real_Tk


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
class GSTSuite(_RealCTk):

    _WARMUP_IMPORTS = (
        "pandas",
        "numpy",
        "openpyxl",
        "fitz",
        "pdfplumber",
        "pdfminer",
        "selenium",
        "webdriver_manager",
        "PIL",
    )

    def __init__(self, user_info: dict = None):
        super().__init__()
        self.title("GST & Income Tax Automation Suite")
        self.geometry("1200x920")
        self.minsize(980, 760)
        self.resizable(True, True)

        # Set taskbar / window icon
        _ico = os.path.join(_ASSETS_BASE, "studycafelogo.ico")
        if os.path.exists(_ico):
            self.iconbitmap(_ico)

        self._user_info      = user_info or {}
        # None = all modules allowed (enterprise / old API); set = restricted plan
        _raw_allowed = (user_info or {}).get("allowed_modules")
        self._allowed = set(_raw_allowed) if _raw_allowed is not None else None
        self._current_theme  = "Dark"
        self._theme_btn      = None
        self._active_cat     = None       # "gst" | "it" | None
        self._active_poll    = None       # matches _active_cat when polling
        self._loaded         = {}         # tab_name -> bool
        self._instances      = {}         # tab_name -> tool instance
        self._module_cache   = {}         # abs_module_path -> imported module object
        self._tabviews       = {}         # cat_key  -> CTkTabview
        self._last_tabs      = {}         # cat_key  -> last active tab
        self._cat_frames     = {}         # cat_key  -> CTkFrame
        self._trial_expired  = False      # set True only if trial expires mid-session
        self._is_closing     = False
        self._after_jobs     = set()
        self._restart_to_login = False

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_header()
        self._content = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self._content.pack(fill="both", expand=True)
        self._build_statusbar()

        self._landing = self._build_landing()
        self._show_landing()

        # Warm heavy non-UI dependencies in the background so first tab load is faster.
        self._queue_after(1200, self._start_dependency_warmup)

        # Start background update check 3 s after the window is ready
        self._queue_after(3000, self._check_for_updates)

    def _queue_after(self, delay_ms: int, callback):
        """Schedule a callback only while the main window is alive."""
        if self._is_closing:
            return None

        holder = {"id": None}

        def _run():
            job_id = holder["id"]
            if job_id is not None:
                self._after_jobs.discard(job_id)
            if self._is_closing:
                return
            try:
                if not self.winfo_exists():
                    return
            except Exception:
                return
            callback()

        try:
            holder["id"] = self.after(delay_ms, _run)
        except Exception:
            return None

        if holder["id"]:
            self._after_jobs.add(holder["id"])
        return holder["id"]

    def _cancel_after_jobs(self):
        for job_id in list(self._after_jobs):
            try:
                self.after_cancel(job_id)
            except Exception:
                pass
        self._after_jobs.clear()

        # Also cancel callbacks scheduled by CTk internals/tool frames.
        _cancel_all_after_callbacks(self)

    def _on_close(self):
        if self._is_closing:
            return
        self._is_closing = True
        self._active_poll = None

        # Force mainloop exit even if hidden toplevels/dialogs still exist.
        try:
            self.quit()
        except Exception:
            pass

        self._cancel_after_jobs()

        try:
            super().destroy()
        except Exception:
            try:
                _tk.Tk.destroy(self)
            except Exception:
                pass

    def _start_dependency_warmup(self):
        def _worker():
            _suite_debug_log("dependency warmup start")
            for mod_name in self._WARMUP_IMPORTS:
                try:
                    __import__(mod_name)
                    _suite_debug_log(f"warmup ok {mod_name}")
                except Exception as e:
                    _suite_debug_log(f"warmup fail {mod_name}: {e}")
            _suite_debug_log("dependency warmup done")
        threading.Thread(target=_worker, daemon=True).start()

    def _import_tool_module(self, module_path: str, uid: str):
        spec = importlib.util.spec_from_file_location(uid, module_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Unable to create module spec for: {module_path}")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod

    def _try_auto_install_missing_dependency(self, missing_module: str, tab_name: str) -> bool:
        package = _missing_package_for_module(missing_module)
        if not package:
            return False
        if getattr(sys, "frozen", False):
            # Frozen EXEs should already ship dependencies; don't mutate environment.
            return False

        _suite_debug_log(
            f"auto_install start tab={tab_name} module={missing_module} package={package}"
        )
        try:
            proc = subprocess.run(
                [sys.executable, "-m", "pip", "install", package],
                capture_output=True,
                text=True,
                timeout=240,
                check=False,
            )
        except Exception as e:
            _suite_debug_log(f"auto_install exception tab={tab_name} package={package}: {e}")
            return False

        if proc.returncode != 0:
            out = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
            _suite_debug_log(f"auto_install failed tab={tab_name} package={package}: {out[-800:]}")
            return False

        importlib.invalidate_caches()
        _suite_debug_log(f"auto_install success tab={tab_name} package={package}")
        return True


    # ══════════════════════════════════════════════════════════════════════════
    #  AUTO-UPDATE
    # ══════════════════════════════════════════════════════════════════════════
    def _check_for_updates(self):
        """Fetch latest.json in a background thread; show popup if newer."""
        def _worker():
            try:
                req = urllib.request.Request(
                    UPDATE_MANIFEST_URL,
                    headers={"User-Agent": "GSTSuite-Updater/1.0"},
                )
                with urllib.request.urlopen(req, timeout=8) as resp:
                    data = json.loads(resp.read())

                latest  = data.get("version", "0.0.0")
                url     = data.get("url", "")
                notes   = data.get("notes", "")

                # Compare as version tuples
                def _v(s):
                    try:
                        return tuple(int(x) for x in s.strip().split("."))
                    except Exception:
                        return (0,)

                if _v(latest) > _v(VERSION) and url:
                    self._queue_after(0, lambda: self._show_update_dialog(latest, url, notes))
            except Exception:
                pass  # silent fail — no internet / server down

        threading.Thread(target=_worker, daemon=True).start()

    def _show_update_dialog(self, latest: str, url: str, notes: str):
        """Show update prompt on the main thread."""
        msg = f"A new version is available!\n\nCurrent: v{VERSION}\nLatest:  v{latest}"
        if notes:
            msg += f"\n\nWhat's new:\n{notes}"
        msg += "\n\nUpdate now?"

        answer = _tk.messagebox.askyesno(
            title="Update Available",
            message=msg,
            icon="info",
        )
        if answer:
            self._launch_updater(url)

    def _launch_updater(self, download_url: str):
        """Extract bundled updater.exe to temp dir, launch it, then close."""
        # Only works when running as a PyInstaller EXE
        if not getattr(sys, "frozen", False):
            _tk.messagebox.showinfo(
                "Update",
                "Running from source — please update manually.",
            )
            return

        updater_src = os.path.join(sys._MEIPASS, "StudyCafeSuite_Updater.exe")
        if not os.path.exists(updater_src):
            _tk.messagebox.showerror(
                "Update Error",
                "StudyCafeSuite_Updater.exe not found inside the bundle.\n"
                "Please re-download StudyCafeSuite.exe from the releases page.",
            )
            return

        tmp_dir     = tempfile.mkdtemp(prefix="studycafe_upd_")
        updater_tmp = os.path.join(tmp_dir, "StudyCafeSuite_Updater.exe")
        shutil.copy2(updater_src, updater_tmp)

        target_exe = sys.executable  # path to the running StudyCafeSuite.exe

        subprocess.Popen(
            [updater_tmp, "--url", download_url, "--target", target_exe, "--restart"],
            creationflags=subprocess.CREATE_NEW_CONSOLE,
        )
        self._on_close()
        sys.exit(0)


    # ══════════════════════════════════════════════════════════════════════════
    #  HEADER
    # ══════════════════════════════════════════════════════════════════════════
    def _build_header(self):
        bar = ctk.CTkFrame(self, fg_color=_C["banner_bg"], corner_radius=0, height=54)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        # Accent stripe
        ctk.CTkFrame(bar, height=3, corner_radius=0,
                     fg_color=_C["primary"]).pack(fill="x", side="top")

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20)

        # Left area — rebuilt on each navigation
        self._hdr_left = ctk.CTkFrame(inner, fg_color="transparent")
        self._hdr_left.pack(side="left", fill="y")

        # Right area — theme toggle + clock
        h_right = ctk.CTkFrame(inner, fg_color="transparent")
        h_right.pack(side="right", fill="y")

        def _do_logout():
            _clear_auth()
            self._restart_to_login = True
            self._on_close()

        ctk.CTkButton(
            h_right, text="⏏  Logout",
            font=("Segoe UI", 11, "bold"),
            width=100, height=32,
            fg_color=("#334155", "#1e293b"),
            hover_color=("#ef4444", "#dc2626"),
            text_color=("#f1f5f9", "#f1f5f9"),
            corner_radius=8,
            command=_do_logout,
        ).pack(side="right", padx=(10, 0))

        try:
            self._theme_btn = ctk.CTkSegmentedButton(
                h_right,
                values=["🌙  Dark", "☀️  Light"],
                command=self._set_theme,
                width=170, height=32,
                font=("Segoe UI", 11, "bold"),
                selected_color=_C["primary"],
                selected_hover_color=_C["primary_hov"],
                unselected_color=("#334155", "#1e293b"),
                unselected_hover_color=("#475569", "#2d3a4a"),
                text_color=("#f1f5f9", "#f1f5f9"),
            )
        except TypeError:
            self._theme_btn = ctk.CTkSegmentedButton(
                h_right, values=["Dark", "Light"],
                command=self._set_theme, width=160, height=32)

        self._theme_btn.set("🌙  Dark")
        self._theme_btn.pack(side="right", padx=(10, 0))

        self._clock_lbl = ctk.CTkLabel(
            h_right, text="",
            font=("Segoe UI Mono", 11),
            text_color="#94a3b8")
        self._clock_lbl.pack(side="right", padx=(0, 18))

        # ── Powered by Study Cafe logo ────────────────────────────────────────
        self._logo_ctk = None
        if _PILImage is not None:
            try:
                _logo_path = os.path.join(_ASSETS_BASE, "studycafelogo.png")
                _pil = _PILImage.open(_logo_path).convert("RGBA")
                _pil.thumbnail((135, 48), _PILImage.LANCZOS)
                self._logo_ctk = ctk.CTkImage(
                    light_image=_pil, dark_image=_pil,
                    size=(_pil.width, _pil.height))
            except Exception:
                self._logo_ctk = None

        logo_frame = ctk.CTkFrame(h_right, fg_color="transparent")
        logo_frame.pack(side="right", padx=(0, 20))
        if self._logo_ctk is not None:
            ctk.CTkLabel(logo_frame, text="Powered by",
                         font=("Segoe UI", 10),
                         text_color="#94a3b8").pack(side="left", padx=(0, 6))
            ctk.CTkLabel(logo_frame, image=self._logo_ctk, text="").pack(side="left")
            ctk.CTkLabel(logo_frame, text="StudyCafe",
                         font=("Segoe UI", 11, "bold"),
                         text_color="#818cf8").pack(side="left", padx=(6, 0))
        else:
            ctk.CTkLabel(logo_frame, text="Powered by",
                         font=("Segoe UI", 10),
                         text_color="#94a3b8").pack(side="left", padx=(0, 6))
            ctk.CTkLabel(logo_frame, text="StudyCafe",
                         font=("Segoe UI", 13, "bold"),
                         text_color="#818cf8").pack(side="left")

        self._tick_clock()

    def _tick_clock(self):
        if self._is_closing:
            return
        self._clock_lbl.configure(
            text=datetime.now().strftime("%d %b %Y   %H:%M:%S"))
        self._queue_after(1000, self._tick_clock)

    def _set_theme(self, label: str):
        mode = "Light" if "Light" in label else "Dark"
        self._current_theme = mode
        _RealSetAppearance(mode)

        # Notify active tool instances of theme change
        for inst in self._instances.values():
            if inst and hasattr(inst, "set_theme"):
                try: inst.set_theme(mode)
                except: pass

        if self._theme_btn:
            self._theme_btn.set("☀️  Light" if mode == "Light" else "🌙  Dark")

    def _refresh_header_left(self, mode: str, cat_key: str = None):
        """Wipe and rebuild the left side of the header."""
        if self._is_closing:
            return
        for w in self._hdr_left.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass

        if mode == "landing":
            ctk.CTkLabel(self._hdr_left, text="GST & Income Tax Automation Suite",
                         font=("Segoe UI", 15, "bold"),
                         text_color=("#f1f5f9", "#f1f5f9")).pack(side="left", pady=4)

        else:
            # Back / Home button
            ctk.CTkButton(
                self._hdr_left,
                text="⬅  Home",
                font=("Segoe UI", 12, "bold"),
                width=108, height=36,
                fg_color=("#334155", "#1e293b"),
                hover_color=("#475569", "#2d3a4a"),
                text_color=("#f1f5f9", "#f1f5f9"),
                corner_radius=8,
                command=self._go_home,
            ).pack(side="left", padx=(0, 20))

            if cat_key == "gst":
                acc, icon, label, tools = _C["gst_acc"],   "🏛",  "GST Tools",                   GST_TOOLS
            elif cat_key == "pdf":
                acc, icon, label, tools = _C["pdf_acc"],   "📄",  "PDF Tools",                   PDF_TOOLS
            elif cat_key == "bank":
                acc, icon, label, tools = _C["bank_acc"],  "🏦",  "Bank Statement → Excel",      BANK_TOOLS
            elif cat_key == "email":
                acc, icon, label, tools = _C["email_acc"], "📧",  "Email Tools",                 EMAIL_TOOLS
            elif cat_key == "reco":
                acc, icon, label, tools = _C["reco_acc"],  "🔄",  "GST Reconciliation",          RECO_TOOLS
            else:
                acc, icon, label, tools = _C["it_acc"],    "💼",  "Income Tax Automation Suite", IT_TOOLS

            ctk.CTkLabel(self._hdr_left, text=label,
                         font=("Segoe UI", 15, "bold"),
                         text_color=acc).pack(side="left", padx=5)


    # ══════════════════════════════════════════════════════════════════════════
    #  STATUS BAR
    # ══════════════════════════════════════════════════════════════════════════
    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, fg_color=_C["status_bg"],
                           corner_radius=0, height=28)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self._statusbar = bar
        ctk.CTkFrame(bar, height=1, corner_radius=0,
                     fg_color=_C["border"]).pack(fill="x", side="top")

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20)

        ctk.CTkLabel(inner, text="●  Ready",
                     font=("Segoe UI", 9, "bold"),
                     text_color=("#10b981", "#10b981")).pack(side="left")
        ctk.CTkLabel(inner, text=f"Automation Suite  v{VERSION}",
                     font=("Segoe UI", 9),
                     text_color="#64748b").pack(side="right")


    # ══════════════════════════════════════════════════════════════════════════
    #  LANDING PAGE
    # ══════════════════════════════════════════════════════════════════════════
    def _build_landing(self) -> ctk.CTkFrame:
        page = ctk.CTkFrame(self._content, fg_color="transparent")

        # Scrollable canvas so the page works at any window height
        scroll = ctk.CTkScrollableFrame(page, fg_color="transparent",
                                        corner_radius=0)
        scroll.pack(fill="both", expand=True)

        # Centered wrapper inside the scrollable area
        wrapper = ctk.CTkFrame(scroll, fg_color="transparent")
        wrapper.pack(expand=True, pady=(30, 30))
        # Horizontally centre the wrapper
        scroll.grid_columnconfigure(0, weight=1)

        # ── Welcome text ──────────────────────────────────────────────────────
        ctk.CTkLabel(
            wrapper,
            text="Welcome — Select a Category",
            font=("Segoe UI", 26, "bold"),
            text_color=_C["text_hi"],
        ).pack(pady=(0, 8))

        ctk.CTkLabel(
            wrapper,
            text="All automation tools, organized by category. "
                 "Click a card to get started.",
            font=("Segoe UI", 13),
            text_color=_C["text_mid"],
        ).pack(pady=(0, 16))

        # ── Row 1: GST · IT · PDF ─────────────────────────────────────────────
        cards_row = ctk.CTkFrame(wrapper, fg_color="transparent")
        cards_row.pack()

        def _sub(tools):
            if self._allowed is None:
                return f"{len(tools)} module{'s' if len(tools) != 1 else ''}"
            n = sum(1 for t in tools if t.get("key") in self._allowed)
            return f"{n} of {len(tools)} modules"

        def _disabled(tools):
            if self._allowed is None:
                return False
            return not any(t.get("key") in self._allowed for t in tools)

        if not _disabled(GST_TOOLS):
            self._make_category_card(
                parent=cards_row,
                icon="🏛",
                label="GST Tools",
                sub_text=_sub(GST_TOOLS),
                desc="Downloads, converters & verifiers\nfor GST portal automation.",
                acc=_C["gst_acc"],
                normal_fg=_C["gst_bg"],
                hover_fg=_C["gst_hover"],
                callback=lambda: self._show_category("gst"),
            ).pack(side="left", padx=14)

        if not _disabled(IT_TOOLS):
            self._make_category_card(
                parent=cards_row,
                icon="💼",
                label="Income Tax Automation Suite",
                sub_text=_sub(IT_TOOLS),
                desc="26AS, Challan & ITR filing\nautomation tools.",
                acc=_C["it_acc"],
                normal_fg=_C["it_bg"],
                hover_fg=_C["it_hover"],
                callback=lambda: self._show_category("it"),
            ).pack(side="left", padx=14)

        if not _disabled(PDF_TOOLS):
            self._make_category_card(
                parent=cards_row,
                icon="📄",
                label="PDF Tools",
                sub_text=_sub(PDF_TOOLS),
                desc="Merge, split, extract, compress\n& redact PDF files.",
                acc=_C["pdf_acc"],
                normal_fg=_C["pdf_bg"],
                hover_fg=_C["pdf_hover"],
                callback=lambda: self._show_category("pdf"),
            ).pack(side="left", padx=14)

        # ── Row 2: Bank · Email · Reco ────────────────────────────────────────
        row2 = ctk.CTkFrame(wrapper, fg_color="transparent")
        row2.pack(pady=(18, 0))

        if not _disabled(BANK_TOOLS):
            self._make_category_card(
                parent=row2,
                icon="🏦",
                label="Bank Statement → Excel",
                sub_text=_sub(BANK_TOOLS),
                desc="Convert bank statement PDFs\nto structured Excel sheets.",
                acc=_C["bank_acc"],
                normal_fg=_C["bank_bg"],
                hover_fg=_C["bank_hover"],
                callback=lambda: self._show_category("bank"),
            ).pack(side="left", padx=14)

        if not _disabled(EMAIL_TOOLS):
            self._make_category_card(
                parent=row2,
                icon="📧",
                label="Email Tools",
                sub_text=_sub(EMAIL_TOOLS),
                desc="Bulk personalised emails via Outlook.\nGST reminders, invoices & more.",
                acc=_C["email_acc"],
                normal_fg=_C["email_bg"],
                hover_fg=_C["email_hover"],
                callback=lambda: self._show_category("email"),
            ).pack(side="left", padx=14)

        if not _disabled(RECO_TOOLS):
            self._make_category_card(
                parent=row2,
                icon="🔄",
                label="GST Reconciliation",
                sub_text=_sub(RECO_TOOLS),
                desc="Reconcile GSTR-2B vs Tally books.\nMatch invoices & export reports.",
                acc=_C["reco_acc"],
                normal_fg=_C["reco_bg"],
                hover_fg=_C["reco_hover"],
                callback=lambda: self._show_category("reco"),
            ).pack(side="left", padx=14)

        return page

    def _make_category_card(self, parent, icon, label, sub_text, desc,
                             acc, normal_fg, hover_fg, callback,
                             disabled: bool = False):
        """Build a large clickable category card. disabled=True greys it out."""
        if disabled:
            acc       = ("#94a3b8", "#475569")
            normal_fg = ("#f1f5f9", "#1e293b")
            hover_fg  = normal_fg

        card = ctk.CTkFrame(
            parent,
            fg_color=normal_fg,
            corner_radius=16,
            border_width=2,
            border_color=acc,
            width=255,
            height=270,
        )
        card.pack_propagate(False)

        # Top accent stripe
        strip = ctk.CTkFrame(card, height=6, corner_radius=16, fg_color=acc)
        strip.pack(fill="x")
        strip.pack_propagate(False)

        # Icon (lock when disabled)
        ctk.CTkLabel(card, text="🔒" if disabled else icon,
                     font=("Segoe UI Emoji", 40)).pack(pady=(16, 2))

        # Category name
        ctk.CTkLabel(card, text=label,
                     font=("Segoe UI", 16, "bold"),
                     text_color=acc, wraplength=220,
                     justify="center").pack()

        # Module count badge
        badge = ctk.CTkFrame(card, fg_color=acc, corner_radius=20)
        badge.pack(pady=(6, 0))
        ctk.CTkLabel(badge, text=f"  {sub_text}  ",
                     font=("Segoe UI", 9, "bold"),
                     text_color="#ffffff").pack(padx=6, pady=3)

        # Description text
        ctk.CTkLabel(card, text=desc,
                     font=("Segoe UI", 11),
                     text_color=("#1e293b", "#e2e8f0"),
                     justify="center", wraplength=220).pack(pady=(10, 4))

        # CTA hint
        if disabled:
            ctk.CTkLabel(card, text="Trial Expired — Locked",
                         font=("Segoe UI", 11, "bold"),
                         text_color=("#94a3b8", "#475569")).pack(pady=(4, 0))
        else:
            ctk.CTkLabel(card, text="Click to explore  →",
                         font=("Segoe UI", 11, "bold"),
                         text_color=acc).pack(pady=(4, 0))
            # Bind hover + click only when active
            self._bind_card(card, callback, normal_fg, hover_fg)
        return card

    def _bind_card(self, card, callback, normal_fg, hover_fg):
        """Recursively bind hover & click to card and every descendant."""
        def enter(_=None):
            try:
                card.configure(fg_color=hover_fg)
            except Exception:
                pass

        def leave(_=None):
            # Only un-hover if pointer has actually left the card bounds
            try:
                px, py = card.winfo_pointerxy()
                cx = card.winfo_rootx()
                cy = card.winfo_rooty()
                cw = card.winfo_width()
                ch = card.winfo_height()
                if not (cx <= px <= cx + cw and cy <= py <= cy + ch):
                    card.configure(fg_color=normal_fg)
            except Exception:
                try:
                    card.configure(fg_color=normal_fg)
                except Exception:
                    pass

        def click(_=None):
            callback()

        def attach(w):
            try:
                w.configure(cursor="hand2")
            except Exception:
                pass
            w.bind("<Enter>",    lambda _: enter(), add="+")
            w.bind("<Leave>",    lambda _: leave(), add="+")
            w.bind("<Button-1>", lambda _: click(), add="+")
            for child in w.winfo_children():
                attach(child)

        attach(card)


    # ══════════════════════════════════════════════════════════════════════════
    #  CATEGORY PAGE  (built lazily on first visit)
    # ══════════════════════════════════════════════════════════════════════════
    def _get_or_build_category(self, key: str) -> ctk.CTkFrame:
        if key in self._cat_frames:
            return self._cat_frames[key]

        if key == "gst":
            tools, accents = GST_TOOLS,   _GST_ACCENTS
        elif key == "pdf":
            tools, accents = PDF_TOOLS,   _PDF_ACCENTS
        elif key == "bank":
            tools, accents = BANK_TOOLS,  _BANK_ACCENTS
        elif key == "email":
            tools, accents = EMAIL_TOOLS, _EMAIL_ACCENTS
        elif key == "reco":
            tools, accents = RECO_TOOLS,  _RECO_ACCENTS
        else:
            tools, accents = IT_TOOLS,    _IT_ACCENTS

        frame = ctk.CTkFrame(self._content, fg_color="transparent",
                             corner_radius=0)

        # Build tabview
        try:
            tv = ctk.CTkTabview(
                frame, anchor="nw",
                segmented_button_selected_color=_C["primary"],
                segmented_button_selected_hover_color=_C["primary_hov"],
                segmented_button_unselected_hover_color=("#334155", "#1e293b"),
                text_color=_C["text_hi"],
                border_width=2,
                border_color=_C["border"],
                fg_color=_C["surface"],
            )
        except (TypeError, ValueError):
            tv = ctk.CTkTabview(frame, anchor="nw")

        # Increase tab label font via the internal segmented button
        acc_color = (_C["gst_acc"]   if key == "gst"   else
                     (_C["pdf_acc"]   if key == "pdf"   else
                      (_C["bank_acc"] if key == "bank"  else
                       (_C["email_acc"] if key == "email" else
                        (_C["reco_acc"] if key == "reco" else _C["it_acc"])))))
        try:
            tv._segmented_button.configure(font=("Segoe UI", 13, "bold"))
        except Exception:
            pass

        # Accent top-stripe on the content frame to mark where tabs end
        try:
            tv._fg_frame.configure(border_width=2, border_color=acc_color)
        except Exception:
            pass


        # Overview tab
        tv.add("🏠  Overview")
        self._build_category_overview(tv.tab("🏠  Overview"), key, tools, accents, tv)

        # Tool tabs (lazy; locked tabs get placeholder immediately)
        for t in tools:
            tv.add(t["tab"])
            if self._allowed is not None and t.get("key") not in self._allowed:
                self._loaded[t["tab"]] = True   # skip lazy-loader
                self._build_locked_tab(tv.tab(t["tab"]), t["tab"])
            else:
                self._loaded[t["tab"]] = False

        tv.pack(fill="both", expand=True, padx=10, pady=(4, 4))

        self._tabviews[key]   = tv
        self._last_tabs[key]  = ""
        self._cat_frames[key] = frame
        return frame

    def _build_locked_tab(self, tab_frame, tool_name: str):
        """Show a lock placeholder for a module not in the user's plan."""
        tab_frame.grid_rowconfigure(0, weight=1)
        tab_frame.grid_columnconfigure(0, weight=1)
        outer = ctk.CTkFrame(tab_frame, fg_color="transparent")
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        box = ctk.CTkFrame(outer, fg_color=_C["surface"], corner_radius=18,
                           border_width=2, border_color=_C["border"],
                           width=340, height=210)
        box.place(relx=0.5, rely=0.5, anchor="center")
        box.pack_propagate(False)

        ctk.CTkLabel(box, text="🔒", font=("Segoe UI Emoji", 44)).pack(pady=(28, 4))
        ctk.CTkLabel(box, text=tool_name, font=("Segoe UI", 14, "bold"),
                     text_color=("#64748b", "#94a3b8"), wraplength=300,
                     justify="center").pack()
        ctk.CTkLabel(box,
                     text="This module is not included in your current plan.\nContact StudyCafe to upgrade your subscription.",
                     font=("Segoe UI", 11), text_color=("#94a3b8", "#64748b"),
                     wraplength=300, justify="center").pack(pady=(10, 0))

    def _build_category_overview(self, frame, key, tools, accents, tv=None):
        COLS  = 3 if len(tools) <= 6 else 4
        if key == "gst":
            acc, icon, label = _C["gst_acc"],   "🏛",  "GST Tools"
        elif key == "pdf":
            acc, icon, label = _C["pdf_acc"],   "📄",  "PDF Tools"
        elif key == "bank":
            acc, icon, label = _C["bank_acc"],  "🏦",  "Bank Statement → Excel"
        elif key == "email":
            acc, icon, label = _C["email_acc"], "📧",  "Email Tools"
        elif key == "reco":
            acc, icon, label = _C["reco_acc"],  "🔄",  "GST Reconciliation"
        else:
            acc, icon, label = _C["it_acc"],    "💼",  "Income Tax Automation Suite"

        scroll = ctk.CTkScrollableFrame(frame, fg_color="transparent",
                                        corner_radius=0)
        scroll.pack(fill="both", expand=True)
        for c in range(COLS):
            scroll.grid_columnconfigure(c, weight=1)

        # Hero
        hero = ctk.CTkFrame(scroll, fg_color=_C["surface"], corner_radius=14,
                            border_width=1, border_color=_C["border"])
        hero.grid(row=0, column=0, columnspan=COLS,
                  padx=16, pady=(18, 12), sticky="ew")

        ctk.CTkFrame(hero, height=6, corner_radius=14,
                     fg_color=acc).pack(fill="x")

        hb = ctk.CTkFrame(hero, fg_color="transparent")
        hb.pack(fill="x", padx=24, pady=(16, 18))

        tr = ctk.CTkFrame(hb, fg_color="transparent")
        tr.pack(fill="x")
        ctk.CTkLabel(tr, text=f"{icon}  {label}",
                     font=("Segoe UI", 21, "bold"),
                     text_color=acc).pack(side="left")
        badge = ctk.CTkFrame(tr, fg_color=acc, corner_radius=20)
        badge.pack(side="left", padx=(12, 0))
        ctk.CTkLabel(badge, text=f"  {len(tools)} Tools  ",
                     font=("Segoe UI", 10, "bold"),
                     text_color="#ffffff").pack(padx=4, pady=4)

        ctk.CTkLabel(hb,
                     text="Click any tab to load a tool — "
                          "each loads once and stays active.",
                     font=("Segoe UI", 13),
                     text_color=("#1e293b", "#e2e8f0")).pack(anchor="w", pady=(8, 0))

        # Section label
        ctk.CTkLabel(scroll, text="Available Tools",
                     font=("Segoe UI", 14, "bold"),
                     text_color=("#0f172a", "#f1f5f9")).grid(
            row=1, column=0, columnspan=COLS,
            padx=18, pady=(6, 8), sticky="w")

        # Tool cards
        for idx, tool in enumerate(tools):
            is_locked = (self._allowed is not None and tool.get("key") not in self._allowed)
            ac    = ("#94a3b8", "#475569") if is_locked else (accents[idx] if idx < len(accents) else acc)
            rn, c = divmod(idx, COLS)

            card = ctk.CTkFrame(scroll, fg_color=_C["surface2"],
                                border_color=_C["border"], border_width=1,
                                corner_radius=12)
            card.grid(row=rn + 2, column=c, padx=9, pady=8, sticky="nsew")
            scroll.grid_rowconfigure(rn + 2, weight=0)

            ctk.CTkFrame(card, height=6, corner_radius=12,
                         fg_color=ac).pack(fill="x")

            body = ctk.CTkFrame(card, fg_color="transparent")
            body.pack(fill="both", expand=True, padx=14, pady=(10, 14))

            tab_label = ("🔒  " + tool["tab"]) if is_locked else tool["tab"]
            ctk.CTkLabel(body, text=tab_label,
                         font=("Segoe UI", 14, "bold"),
                         text_color=ac).pack(anchor="w")
            ctk.CTkFrame(body, height=1, corner_radius=0,
                         fg_color=_C["border"]).pack(fill="x", pady=(5, 8))
            ctk.CTkLabel(body, text=tool["desc"],
                         font=("Segoe UI", 12),
                         text_color=("#1e293b", "#e2e8f0"),
                         wraplength=230, justify="left").pack(anchor="w")
            action_text = "🔒  Upgrade to access" if is_locked else "→  Click to open"
            ctk.CTkLabel(body, text=action_text,
                         font=("Segoe UI", 11, "bold"),
                         text_color=ac).pack(anchor="w", pady=(10, 0))

            # Make card clickable (only if not locked) — switches directly to the tool's tab
            if tv is not None and not is_locked:
                def _make_attach(tab_name=tool["tab"], _card=card):
                    def _click(_=None): tv.set(tab_name)
                    def _enter(_=None):
                        try: _card.configure(fg_color=_C["surface"])
                        except Exception: pass
                    def _leave(_=None):
                        try: _card.configure(fg_color=_C["surface2"])
                        except Exception: pass
                    def _attach(w):
                        try: w.configure(cursor="hand2")
                        except Exception: pass
                        w.bind("<Button-1>", _click, add="+")
                        w.bind("<Enter>",    _enter, add="+")
                        w.bind("<Leave>",    _leave, add="+")
                        for child in w.winfo_children():
                            _attach(child)
                    _attach(_card)
                _make_attach()


    # ══════════════════════════════════════════════════════════════════════════
    #  NAVIGATION
    # ══════════════════════════════════════════════════════════════════════════
    def _show_landing(self):
        if self._is_closing:
            return
        self._active_poll = None
        for f in self._cat_frames.values():
            f.pack_forget()
        self._landing.pack(fill="both", expand=True)
        self._refresh_header_left("landing")

    def _show_category(self, key: str):
        if self._is_closing:
            return
        if self._trial_expired:
            return   # Hard block — trial has expired
        self._landing.pack_forget()
        # Hide any other category that might be visible
        for k, f in self._cat_frames.items():
            if k != key:
                f.pack_forget()

        # Hide the statusbar when entering the Bank Statement tool
        if key == "bank":
            self._statusbar.pack_forget()
        else:
            self._statusbar.pack(fill="x", side="bottom")

        cat_frame = self._get_or_build_category(key)
        cat_frame.pack(fill="both", expand=True)
        self._active_cat  = key
        self._active_poll = key
        self._refresh_header_left("category", key)
        self._queue_after(300, lambda: self._poll_tab(key))

    def _go_home(self):
        if self._is_closing:
            return
        if self._active_cat and self._active_cat in self._cat_frames:
            self._cat_frames[self._active_cat].pack_forget()
        self._active_cat  = None
        self._active_poll = None
        self._statusbar.pack(fill="x", side="bottom")
        self._show_landing()


    # ══════════════════════════════════════════════════════════════════════════
    #  LAZY TAB LOADER
    # ══════════════════════════════════════════════════════════════════════════
    def _poll_tab(self, key: str):
        if self._is_closing:
            return
        if self._active_poll != key:
            return   # category no longer active
        tv = self._tabviews.get(key)
        if not tv:
            self._queue_after(300, lambda: self._poll_tab(key))
            return
        try:
            current = tv.get()
        except Exception:
            self._queue_after(300, lambda: self._poll_tab(key))
            return

        last = self._last_tabs.get(key, "")
        if current != last:
            self._last_tabs[key] = current
            if current != "🏠  Overview" and not self._loaded.get(current, True):
                self._activate_tab(key, current)

        self._queue_after(300, lambda: self._poll_tab(key))

    def _activate_tab(self, cat_key: str, name: str):
        if self._is_closing:
            return
        if cat_key == "gst":
            tools, accents = GST_TOOLS,   _GST_ACCENTS
        elif cat_key == "pdf":
            tools, accents = PDF_TOOLS,   _PDF_ACCENTS
        elif cat_key == "bank":
            tools, accents = BANK_TOOLS,  _BANK_ACCENTS
        elif cat_key == "email":
            tools, accents = EMAIL_TOOLS, _EMAIL_ACCENTS
        elif cat_key == "reco":
            tools, accents = RECO_TOOLS,  _RECO_ACCENTS
        else:
            tools, accents = IT_TOOLS,    _IT_ACCENTS
        tool    = next((t for t in tools if t["tab"] == name), None)
        if tool is None:
            return
        _suite_debug_log(
            f"activate_tab start cat={cat_key} tab={name} module={tool['module']} exists={os.path.exists(tool['module'])} tk={bool(tool.get('tk'))}"
        )
        idx    = next((i for i, t in enumerate(tools) if t["tab"] == name), 0)
        accent = accents[idx] if idx < len(accents) else _C["primary"]

        tv = self._tabviews.get(cat_key)
        if tv is None:
            return
        try:
            tab_frame = tv.tab(name)
        except Exception:
            return
        try:
            if not tab_frame.winfo_exists():
                return
        except Exception:
            return

        # ── Enhanced Loading overlay with better visibility and animations ──
        overlay = ctk.CTkFrame(tab_frame, fg_color=_C["surface"],
                               corner_radius=20, border_width=3,
                               border_color=accent)
        overlay.place(relx=0.5, rely=0.5, anchor="center",
                      relwidth=0.42, relheight=0.35)

        # Loading container
        load_container = ctk.CTkFrame(overlay, fg_color="transparent")
        load_container.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.9, relheight=0.9)

        # Animated spinner emojis
        spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        spinner_label = ctk.CTkLabel(load_container, text=spinner_frames[0],
                                     font=("Segoe UI Emoji", 48),
                                     text_color=accent)
        spinner_label.pack(pady=(0, 12))

        # "Loading tool…" label
        status_label = ctk.CTkLabel(load_container, text="Loading tool…",
                                    font=("Segoe UI", 15, "bold"),
                                    text_color=accent)
        status_label.pack(pady=(8, 4))

        # Tool name subtitle
        tool_name_label = ctk.CTkLabel(load_container, text=name,
                                       font=("Segoe UI", 12),
                                       text_color=_C["text_mid"])
        tool_name_label.pack(pady=(0, 12))

        # Progress indicator
        progress_frame = ctk.CTkFrame(load_container, fg_color=_C["border"],
                                      corner_radius=6, height=4)
        progress_frame.pack(fill="x", pady=(8, 8))
        progress_frame.pack_propagate(False)

        progress_bar = ctk.CTkFrame(progress_frame, fg_color=accent,
                                    corner_radius=6, height=4)
        progress_bar.pack(side="left", fill="y", padx=0)

        # Subtitle
        info_label = ctk.CTkLabel(load_container, 
                                  text="Initializing modules & dependencies…",
                                  font=("Segoe UI", 10),
                                  text_color=_C["text_lo"],
                                  wraplength=240,
                                  justify="center")
        info_label.pack(pady=(4, 0))

        # Force the overlay to render NOW
        self.update_idletasks()

        # Animate the spinner
        spinner_idx = [0]
        def _animate_spinner():
            if self._is_closing:
                return
            try:
                alive = overlay.winfo_exists()
            except Exception:
                return
            if alive:
                spinner_idx[0] = (spinner_idx[0] + 1) % len(spinner_frames)
                spinner_label.configure(text=spinner_frames[spinner_idx[0]])
                
                # Animate progress bar width
                try:
                    cur_width = progress_bar.winfo_width()
                    max_width = progress_frame.winfo_width()
                    if cur_width < max_width - 20:
                        progress_bar.configure(width=cur_width + 20)
                except:
                    pass
                self._queue_after(80, _animate_spinner)

        # Start spinner animation
        self._queue_after(80, _animate_spinner)

        # Importing tkinter/customtkinter modules from a background thread is not
        # thread-safe and can crash/destroy the Tcl app on Windows. Keep both
        # import and instantiation on the main UI thread.
        _real_Tk_saved = _tk.Tk
        _saved_mode    = ctk.get_appearance_mode()
        _tk_container  = None

        tab_frame.grid_rowconfigure(0, weight=1)
        tab_frame.grid_columnconfigure(0, weight=1)

        if tool.get("tk"):
            # For plain tkinter tools we only need to monkeypatch tk.Tk at import
            # time so class definitions bind to _EmbeddedTkFrame. The actual host
            # frame is created later during instantiation, keeping the loading
            # overlay visible while heavy imports run.
            _tk.Tk = _EmbeddedTkFrame
        else:
            _EmbeddedFrame._host        = tab_frame
            ctk.CTk                     = _EmbeddedFrame
            ctk.set_appearance_mode     = lambda *a, **k: None

        def _do_load_on_main_thread():
            nonlocal _tk_container
            try:
                if not tab_frame.winfo_exists():
                    _suite_debug_log(f"load aborted tab={name}: tab frame no longer exists")
                    return
                module_path = os.path.abspath(tool["module"])
                if not os.path.exists(module_path):
                    raise FileNotFoundError(f"Tool module not found: {module_path}")

                mod = self._module_cache.get(module_path)
                if mod is None:
                    tool_dir = os.path.dirname(module_path)
                    if tool_dir not in sys.path:
                        sys.path.insert(0, tool_dir)
                    uid  = f"_suite_mod_{len(self._module_cache)}_{os.path.basename(tool_dir).replace(' ', '_')}"
                    _suite_debug_log(f"main_import begin tab={name} uid={uid}")
                    try:
                        mod = self._import_tool_module(module_path, uid)
                    except ModuleNotFoundError as mnf:
                        missing_mod = getattr(mnf, "name", "")
                        missing_pkg = _missing_package_for_module(missing_mod)
                        if missing_pkg:
                            status_label.configure(text=f"Installing dependency: {missing_pkg}…")
                            info_label.configure(text="One-time setup in progress. Please wait…")
                            self.update_idletasks()
                            if self._try_auto_install_missing_dependency(missing_mod, name):
                                status_label.configure(text="Retrying tool load…")
                                info_label.configure(text="Dependency installed. Re-importing module…")
                                self.update_idletasks()
                                mod = self._import_tool_module(module_path, uid)
                            else:
                                raise
                        else:
                            raise
                    self._module_cache[module_path] = mod
                    _suite_debug_log(f"main_import success tab={name} class={tool['class']}")
                else:
                    _suite_debug_log(f"main_import cache_hit tab={name} module={module_path}")

                AppClass = getattr(mod, tool["class"])
                if tool.get("tk"):
                    # Build tkinter host only when the class is ready to instantiate.
                    _tk_container = _tk.Frame(tab_frame, bg="#1e1e2e")
                    _tk_container.place(x=0, y=0, relwidth=1, relheight=1)
                    _tk_container.lift()
                    _EmbeddedTkFrame._host = _tk_container
                inst = AppClass()
                self._instances[name] = inst
                self._loaded[name]    = True
                _suite_debug_log(f"instantiate success tab={name} instance={type(inst).__name__}")
                if inst and hasattr(inst, "set_theme"):
                    try: inst.set_theme(self._current_theme)
                    except: pass
            except Exception as e:
                err_msg = traceback.format_exc()
                if isinstance(e, ModuleNotFoundError):
                    missing_mod = getattr(e, "name", "") or "unknown"
                    pkg = _missing_package_for_module(missing_mod) or missing_mod
                    if getattr(sys, "frozen", False):
                        hint = (
                            f"Missing dependency '{missing_mod}'.\n"
                            f"This bundled EXE is missing package '{pkg}'.\n"
                            "Rebuild/reinstall the suite with complete dependencies."
                        )
                    else:
                        hint = (
                            f"Missing dependency '{missing_mod}'.\n"
                            f"Install package '{pkg}' and reopen this tab:\n"
                            f"  {sys.executable} -m pip install {pkg}"
                        )
                    err_msg = f"{hint}\n\n{err_msg}"
                print(f"[ERROR] Failed to load tool '{name}': {err_msg}")
                _suite_debug_log(f"load error tab={name}: {err_msg.splitlines()[-1] if err_msg else 'unknown'}")
                try:
                    # If tk host overlay exists, remove it so the CTk traceback panel is visible.
                    if tool.get("tk") and _tk_container is not None:
                        try:
                            _tk_container.destroy()
                        except Exception:
                            pass
                    ef = ctk.CTkFrame(tab_frame, fg_color="transparent")
                    ef.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)
                    hdr = ctk.CTkFrame(ef, fg_color=("#fef2f2", "#1c0a0a"),
                                       corner_radius=10,
                                       border_color=("#fca5a5", "#7f1d1d"),
                                       border_width=1)
                    hdr.pack(fill="x", pady=(0, 14))
                    ctk.CTkLabel(hdr, text=f"  ⚠️  Failed to load  {name}",
                                 font=("Segoe UI", 13, "bold"),
                                 text_color=("#ef4444", "#f87171")).pack(anchor="w", padx=18, pady=12)
                    tb_box = ctk.CTkTextbox(ef, font=("Consolas", 11),
                                            text_color=("#ef4444", "#f87171"),
                                            fg_color=("#1a1a2e", "#0d0d1a"),
                                            corner_radius=10, border_width=1,
                                            border_color=("#7f1d1d", "#450a0a"), height=320)
                    tb_box.pack(fill="both", expand=True)
                    tb_box.insert("1.0", err_msg)
                    tb_box.configure(state="disabled")
                except Exception:
                    pass
            finally:
                # Always restore ctk.CTk / tk.Tk regardless of success or error
                if tool.get("tk"):
                    _tk.Tk = _real_Tk_saved
                else:
                    ctk.CTk                 = _RealCTk
                    ctk.set_appearance_mode = _RealSetAppearance
                    _RealSetAppearance(_saved_mode)
                try: overlay.destroy()
                except: pass

        self._queue_after(10, _do_load_on_main_thread)


# ══════════════════════════════════════════════════════════════════════════════
#  DEVICE MANAGER DIALOG
# ══════════════════════════════════════════════════════════════════════════════
class DeviceManagerDialog(ctk.CTkToplevel):
    """Modal popup shown when device limit is reached.
    Lets the user remove a registered device so they can log in."""

    def __init__(self, parent, devices: list, email: str, password: str, on_success):
        super().__init__(parent)
        self.title("Device Limit Reached")
        self.geometry("480x380")
        self.resizable(False, False)
        self.grab_set()
        self.lift()
        self.attributes("-topmost", True)
        self._email      = email
        self._password   = password
        self._on_success = on_success

        # Centre over parent
        parent.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - 480) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 380) // 2
        self.geometry(f"480x380+{px}+{py}")

        # Accent stripe
        ctk.CTkFrame(self, height=4, corner_radius=0,
                     fg_color=("#f59e0b", "#d97706")).pack(fill="x")

        ctk.CTkLabel(self, text="⚠  Device Limit Reached",
                     font=("Segoe UI", 16, "bold"),
                     text_color=("#d97706", "#fbbf24")).pack(pady=(18, 4))
        ctk.CTkLabel(self,
                     text="Your account has reached its device limit.\nRemove a device below to log in on this machine.",
                     font=("Segoe UI", 12),
                     text_color=("#475569", "#94a3b8"),
                     justify="center").pack(pady=(0, 16))

        ctk.CTkFrame(self, height=1, corner_radius=0,
                     fg_color=("#e2e8f0", "#334155")).pack(fill="x", padx=24)

        # Scrollable device list
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent", height=160)
        scroll.pack(fill="x", padx=24, pady=12)

        self._err_lbl = ctk.CTkLabel(self, text="",
                                     font=("Segoe UI", 11),
                                     text_color=("#ef4444", "#f87171"))
        self._err_lbl.pack()

        for dev in devices:
            row = ctk.CTkFrame(scroll, fg_color=("#f1f5f9", "#1e293b"),
                               corner_radius=8, height=46)
            row.pack(fill="x", pady=4)
            row.pack_propagate(False)

            added = str(dev.get("added_at", ""))[:16]
            hw    = str(dev.get("hardware_id", "Unknown"))
            did   = dev.get("device_id")

            ctk.CTkLabel(row, text=f"🖥  {hw}",
                         font=("Segoe UI", 12, "bold"),
                         text_color=("#1e293b", "#f1f5f9")).pack(side="left", padx=12)
            ctk.CTkLabel(row, text=added,
                         font=("Segoe UI", 10),
                         text_color=("#64748b", "#64748b")).pack(side="left")
            ctk.CTkButton(row, text="Remove",
                          font=("Segoe UI", 11, "bold"),
                          width=80, height=30,
                          fg_color=("#ef4444", "#dc2626"),
                          hover_color=("#b91c1c", "#ef4444"),
                          corner_radius=6,
                          command=lambda d=did: self._remove(d)).pack(side="right", padx=10)

        ctk.CTkButton(self, text="Cancel",
                      font=("Segoe UI", 12),
                      width=120, height=36,
                      fg_color=("#334155", "#1e293b"),
                      hover_color=("#475569", "#2d3a4a"),
                      corner_radius=8,
                      command=self.destroy).pack(pady=(4, 16))

    def _remove(self, device_id):
        self._err_lbl.configure(text="Removing…")
        self.update()
        try:
            resp = _call_api("/remove_device", {
                "email":     self._email,
                "password":  self._password,
                "device_id": device_id,
            })
            if resp.get("status") == "SUCCESS":
                self.destroy()
                try:
                    self.after(0, self._on_success)
                except Exception:
                    self._on_success()
            else:
                self._err_lbl.configure(text=f"Error: {resp.get('status')}")
        except Exception as e:
            self._err_lbl.configure(text=f"Network error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
#  LOGIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
class LoginWindow(_RealCTk):
    """Standalone login window shown before the main app."""

    def __init__(self):
        super().__init__()
        self.title("GST Suite — Login")
        self.geometry("420x520")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._close_window)
        self._auth_result = None

        _ico = os.path.join(_ASSETS_BASE, "studycafelogo.ico")
        if os.path.exists(_ico):
            self.iconbitmap(_ico)

        # Centre on screen
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"420x520+{(sw-420)//2}+{(sh-520)//2}")

        self._hw = _get_hardware_id()
        self._build()

    def _close_window(self):
        try:
            self.quit()
        except Exception:
            pass

        _cancel_all_after_callbacks(self)

        try:
            self.destroy()
        except Exception:
            try:
                _tk.Tk.destroy(self)
            except Exception:
                pass

    def get_auth_result(self):
        return self._auth_result

    def _build(self):
        # Accent stripe
        ctk.CTkFrame(self, height=4, corner_radius=0,
                     fg_color=("#6366f1", "#818cf8")).pack(fill="x")

        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.pack(expand=True, fill="both", padx=40, pady=30)

        # Logo / title
        _logo_ctk = None
        if _PILImage is not None:
            try:
                _pil = _PILImage.open(os.path.join(_ASSETS_BASE, "studycafelogo.png")).convert("RGBA")
                _pil.thumbnail((120, 44), _PILImage.LANCZOS)
                _logo_ctk = ctk.CTkImage(light_image=_pil, dark_image=_pil,
                                         size=(_pil.width, _pil.height))
            except Exception:
                pass

        if _logo_ctk:
            ctk.CTkLabel(wrap, image=_logo_ctk, text="").pack(pady=(0, 6))

        ctk.CTkLabel(wrap, text="GST & Income Tax Suite",
                     font=("Segoe UI", 18, "bold"),
                     text_color=("#1e293b", "#f1f5f9")).pack()
        ctk.CTkLabel(wrap, text="Sign in to continue",
                     font=("Segoe UI", 12),
                     text_color=("#64748b", "#94a3b8")).pack(pady=(2, 24))

        # Email
        ctk.CTkLabel(wrap, text="Email",
                     font=("Segoe UI", 11, "bold"),
                     text_color=("#475569", "#94a3b8"),
                     anchor="w").pack(fill="x")
        self._user_entry = ctk.CTkEntry(wrap, placeholder_text="Enter your email",
                                        height=40, corner_radius=8,
                                        font=("Segoe UI", 13))
        self._user_entry.pack(fill="x", pady=(4, 14))

        # Password
        ctk.CTkLabel(wrap, text="Password",
                     font=("Segoe UI", 11, "bold"),
                     text_color=("#475569", "#94a3b8"),
                     anchor="w").pack(fill="x")
        self._pass_entry = ctk.CTkEntry(wrap, placeholder_text="Enter your password",
                                        show="●", height=40, corner_radius=8,
                                        font=("Segoe UI", 13))
        self._pass_entry.pack(fill="x", pady=(4, 6))
        self._pass_entry.bind("<Return>", lambda e: self._do_login())

        # Status / error label
        self._status_lbl = ctk.CTkLabel(wrap, text="",
                                         font=("Segoe UI", 11),
                                         text_color=("#ef4444", "#f87171"),
                                         wraplength=320)
        self._status_lbl.pack(pady=(6, 0))

        # Login button
        self._login_btn = ctk.CTkButton(wrap, text="Login",
                                         font=("Segoe UI", 13, "bold"),
                                         height=42, corner_radius=8,
                                         fg_color=("#6366f1", "#6366f1"),
                                         hover_color=("#4f46e5", "#4f46e5"),
                                         command=self._do_login)
        self._login_btn.pack(fill="x", pady=(14, 8))

        # Free trial button — opens registration page in browser
        ctk.CTkButton(wrap, text="Try 7-Day Free Trial",
                      font=("Segoe UI", 12),
                      height=38, corner_radius=8,
                      fg_color="transparent",
                      border_width=1,
                      border_color=("#6366f1", "#818cf8"),
                      text_color=("#6366f1", "#818cf8"),
                      hover_color=("#ede9fe", "#1e1b4b"),
                      command=lambda: webbrowser.open(REGISTER_URL)).pack(fill="x")

    def _set_status(self, text, color="#ef4444"):
        self._status_lbl.configure(text=text, text_color=(color, color))
        self.update()

    def _do_login(self):
        email = self._user_entry.get().strip()
        password = self._pass_entry.get()
        if not email or not password:
            self._set_status("Please enter both email and password.")
            return

        self._login_btn.configure(state="disabled", text="Authenticating…")
        self._set_status("")
        self.update()

        try:
            resp = _call_api("/authenticate", {
                "email":       email,
                "password":    password,
                "hardware_id": self._hw,
            })
        except Exception as e:
            self._set_status(f"Network error: {e}")
            self._login_btn.configure(state="normal", text="Login")
            return

        status = resp.get("status")

        if status == "SUCCESS":
            _save_auth(email, password)
            self._auth_result = resp
            self._close_window()
            return

        elif status == "INVALID_CREDENTIALS":
            self._set_status("Invalid email or password. Please try again.")
            self._login_btn.configure(state="normal", text="Login")

        elif status == "TRIAL_EXPIRED":
            self._set_status(
                "Your 7-day trial has expired.\nPlease contact StudyCafe to upgrade your account.",
                color="#f59e0b"
            )
            self._login_btn.configure(state="normal", text="Login")

        elif status == "LIMIT_REACHED":
            self._login_btn.configure(state="normal", text="Login")
            devices = [
                {"device_id": d["device_id"], "hardware_id": d["hardware_id"], "added_at": d["added_at"]}
                for d in (resp.get("registered_devices") or [])
            ]
            DeviceManagerDialog(
                parent=self,
                devices=devices,
                email=email,
                password=password,
                on_success=self._do_login,
            )

        else:
            self._set_status(f"Unexpected response: {status}")
            self._login_btn.configure(state="normal", text="Login")


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def _relaunch_self():
    """Restart the current app process for a clean Login→Suite cycle."""
    try:
        if getattr(sys, "frozen", False):
            os.execl(sys.executable, sys.executable)
        else:
            os.execl(sys.executable, sys.executable, os.path.abspath(__file__))
    except Exception as e:
        _suite_debug_log(f"relaunch failed: {e}")


def run_app_lifecycle():
    """Single top-level event-loop orchestration to avoid nested mainloops."""
    login = LoginWindow()
    try:
        login.mainloop()
    except KeyboardInterrupt:
        try:
            login._close_window()
        except Exception:
            pass
        return

    user_info = login.get_auth_result()
    if not user_info:
        return

    suite = GSTSuite(user_info=user_info)
    try:
        suite.mainloop()
    except KeyboardInterrupt:
        try:
            suite._on_close()
        except Exception:
            pass
        return

    if getattr(suite, "_restart_to_login", False):
        _relaunch_self()


if __name__ == "__main__":
    try:
        if _SPLASH is not None:
            _SPLASH.destroy()
    except Exception:
        pass
    run_app_lifecycle()
