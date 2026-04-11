"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         GST & INCOME TAX AUTOMATION SUITE  —  Unified Launcher               ║
║                                                                              ║
║  Landing page → select category → tabbed tool view → ← Home to go back.      ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import importlib.util
import traceback
import tkinter as _tk
import customtkinter as ctk
from datetime import datetime
try:
    from PIL import Image as _PILImage
except ImportError:
    _PILImage = None

# ── Appearance (must run before any tool import) ─────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

_RealCTk           = ctk.CTk
_RealSetAppearance = ctk.set_appearance_mode

# ── Base paths ────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    _BASE = sys._MEIPASS
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))

_GST_BASE   = os.path.join(_BASE, "GST")
_IT_BASE    = os.path.join(_BASE, "Income Tax")
_PDF_BASE   = os.path.join(_BASE, "PDF_Utilities")
_BANK_BASE  = os.path.join(_BASE, "Bank Statement To Excel")
_EMAIL_BASE = os.path.join(_BASE, "Outlook Email Tools")
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
        self.grid(row=0, column=0, sticky="nsew")

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
    {"tab": "📥  GSTR-2B",      "module": os.path.join(_GST_BASE, "GST 2B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-2B returns via automated browser."},
    {"tab": "📥  GSTR-3B",      "module": os.path.join(_GST_BASE, "GST 3B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-3B returns via automated browser."},
    {"tab": "📊  3B → Excel",   "module": os.path.join(_GST_BASE, "GST 3B to Excel",       "main.py"),        "class": "GSTR3BConverterPro", "desc": "Convert GSTR-3B PDF files to formatted Excel sheets."},
    {"tab": "🤖  GST Verifier", "module": os.path.join(_GST_BASE, "GST Bot",               "gst_pro_app.py"), "class": "GSTApp",             "desc": "Verify bulk GSTINs and extract filing history & details."},
    {"tab": "💰  Challan",      "module": os.path.join(_GST_BASE, "GST Challan Downloader","main.py"),        "class": "App",                "desc": "Download GST Challan PDFs in bulk (Monthly / Quarterly)."},
    {"tab": "📑  R1 JSON",      "module": os.path.join(_GST_BASE, "GST R1 Downloader",     "mai.py"),         "class": "App",                "desc": "Request or download GSTR-1 JSON files for multiple users."},
    {"tab": "📊  JSON → Excel", "module": os.path.join(_GST_BASE, "JSON to Excel",          "main.py"),        "class": "App",                "desc": "Convert GSTR-1 JSON exports to multi-sheet Excel reports."},
    {"tab": "🖨️  R1 PDF",       "module": os.path.join(_GST_BASE, "R1 PDF Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-1 PDF filed    returns from the GST portal."},
    {"tab": "📲  IMS",          "module": os.path.join(_GST_BASE, "IMS Downloader",        "main.py"),        "class": "App",                "desc": "Download IMS (Invoice Management System) data in bulk from the GST portal."},
    {"tab": "📋  GSTR1 Cons.", "module": os.path.join(_GST_BASE, "GSTR1_Consolidation",  "gst_consolidation.py"), "class": "ChallExtractorApp", "tk": True, "desc": "Consolidate multiple GSTR-1 files into a single unified Excel report."},
]

IT_TOOLS = [
    {"tab": "📄  26/AIS/TIS",          "module": os.path.join(_IT_BASE, "26 AS Downlaoder",  "main.py"),                  "class": "App",              "desc": "Download 26AS / AIS / TIS reports in bulk."},
    {"tab": "💰  Challan Downloader", "module": os.path.join(_IT_BASE, "Challan Downloader","main.py"),                  "class": "App",              "desc": "Download Income Tax Challan PDFs in bulk."},
    {"tab": "🤖  ITR Bot",            "module": os.path.join(_IT_BASE, "ITR - Bot",         "GUI_based_app.py"),         "class": "App",              "desc": "Automate ITR filing workflows with the ITR bot."},
    {"tab": "🔍  Demand Checker",     "module": os.path.join(_IT_BASE, "Challan Downloader","demand_checker_app.py"),    "class": "DemandCheckerApp", "desc": "Check pending worklist and outstanding demands in bulk from the Income Tax portal."},
    {"tab": "📊  Refund Checker",     "module": os.path.join(_IT_BASE, "26 AS Downlaoder",  "refund_checker_app.py"),    "class": "RefundCheckerApp", "desc": "Extract filed return data and generate refund status reports in bulk."},
]

PDF_TOOLS = [
    {"tab": "⊕  Merge",    "module": os.path.join(_PDF_BASE, "main.py"), "class": "MergeApp",    "tk": True, "desc": "Merge multiple PDF files into one high-quality document."},
    {"tab": "✂  Split",    "module": os.path.join(_PDF_BASE, "main.py"), "class": "SplitApp",    "tk": True, "desc": "Split PDF files into smaller parts by page range or every N pages."},
    {"tab": "⊙  Extract",  "module": os.path.join(_PDF_BASE, "main.py"), "class": "ExtractApp",  "tk": True, "desc": "Extract specific pages from a PDF document to a new file."},
    {"tab": "⊜  Compress", "module": os.path.join(_PDF_BASE, "main.py"), "class": "CompressApp", "tk": True, "desc": "Reduce PDF file size while maintaining visual quality."},
    {"tab": "⬛  Redact",   "module": os.path.join(_PDF_BASE, "main.py"), "class": "RedactApp",   "tk": True, "desc": "Securely black out sensitive information and text from PDF documents."},
]

BANK_TOOLS = [
    {"tab": "🏦  Bank → Excel", "module": os.path.join(_BANK_BASE, "bank_to_excel.py"), "class": "App", "tk": True, "desc": "Convert bank statement PDFs to formatted Excel sheets. Supports HDFC, ICICI, SBI, Axis, Kotak, IDFC, BOI, Yes, UCO & Equitas."},
]

EMAIL_TOOLS = [
    {"tab": "📋  GST Return Request", "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "GSTReturnMailApp",      "tk": True, "desc": "Send bulk GST return data request emails via Outlook. Auto-fills month, return type, deadlines and contact details."},
    {"tab": "🧾  Invoice Sender",     "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "InvoiceSenderMailApp",   "tk": True, "desc": "Dispatch personalised invoices to clients in bulk via Outlook. Supports per-row service, period, amount & PDF attachments."},
    {"tab": "💰  Payment Reminder",   "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "PaymentReminderMailApp", "tk": True, "desc": "Send outstanding payment reminder emails in bulk via Outlook. Includes interest clause, deadline and per-client amounts."},
]

RECO_TOOLS = [
    {"tab": "🔄  GST Reconciliation", "module": os.path.join(_RECO_BASE, "mainpy-reco-speqtra.py"), "class": "App", "tk": False, "desc": "Reconcile GSTR-2B portal data against Tally/books. Matches invoices, highlights mismatches and exports a detailed Excel report."},
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

    def __init__(self):
        super().__init__()
        self.title("GST & Income Tax Automation Suite")
        self.geometry("1200x920")
        self.minsize(980, 760)
        self.resizable(True, True)
        self.attributes("-topmost", True)

        # Set taskbar / window icon
        _ico = os.path.join(_BASE, "studycafelogo.ico")
        if os.path.exists(_ico):
            self.iconbitmap(_ico)

        self._current_theme  = "Dark"
        self._theme_btn      = None
        self._active_cat     = None       # "gst" | "it" | None
        self._active_poll    = None       # matches _active_cat when polling
        self._loaded         = {}         # tab_name -> bool
        self._instances      = {}         # tab_name -> tool instance
        self._tabviews       = {}         # cat_key  -> CTkTabview
        self._last_tabs      = {}         # cat_key  -> last active tab
        self._cat_frames     = {}         # cat_key  -> CTkFrame

        self._build_header()
        self._content = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self._content.pack(fill="both", expand=True)
        self._build_statusbar()

        self._landing = self._build_landing()
        self._show_landing()


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
                _logo_path = os.path.join(_BASE, "studycafelogo.png")
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
        self._clock_lbl.configure(
            text=datetime.now().strftime("%d %b %Y   %H:%M:%S"))
        self.after(1000, self._tick_clock)

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
        for w in self._hdr_left.winfo_children():
            w.destroy()

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
        ctk.CTkLabel(inner, text="Automation Suite  v2.0",
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
        ).pack(pady=(0, 28))

        # ── Row 1: GST · IT · PDF ─────────────────────────────────────────────
        # ── Dynamic 3x3 Grid Layout ───────────────────────────────────────────
        def _sub(tools):
            return f"{len(tools)} module{'s' if len(tools) != 1 else ''}"

        categories = [
            (GST_TOOLS, "🏛", "GST Tools", "Downloads, converters & verifiers\nfor GST portal automation.", _C["gst_acc"], _C["gst_bg"], _C["gst_hover"], lambda: self._show_category("gst")),
            (IT_TOOLS, "💼", "Income Tax Automation Suite", "26AS, Challan & ITR filing\nautomation tools.", _C["it_acc"], _C["it_bg"], _C["it_hover"], lambda: self._show_category("it")),
            (PDF_TOOLS, "📄", "PDF Tools", "Merge, split, extract, compress\n& redact PDF files.", _C["pdf_acc"], _C["pdf_bg"], _C["pdf_hover"], lambda: self._show_category("pdf")),
            (BANK_TOOLS, "🏦", "Bank Statement → Excel", "Convert bank statement PDFs\nto structured Excel sheets.", _C["bank_acc"], _C["bank_bg"], _C["bank_hover"], lambda: self._show_category("bank")),
            (EMAIL_TOOLS, "📧", "Outlook Email Tools", "Bulk personalised emails via Outlook.\nGST reminders, invoices & more.", _C["email_acc"], _C["email_bg"], _C["email_hover"], lambda: self._show_category("email")),
            (RECO_TOOLS, "🔄", "GST Reconciliation", "Reconcile GSTR-2B vs Tally books.\nMatch invoices & export reports.", _C["reco_acc"], _C["reco_bg"], _C["reco_hover"], lambda: self._show_category("reco")),
            (GMAIL_TOOLS, "✉", "Gmail Tools", "Bulk personalised emails via Gmail.\nGST reminders, invoices & more.", _C["gmail_acc"], _C["gmail_bg"], _C["gmail_hover"], lambda: self._show_category("gmail")),
            (TALLY_TOOLS, "🧾", "Tally Automation Tools", "Convert GSTR-2B/Tally data\nto Tally-ready Excel and XML.", _C["tally_acc"], _C["tally_bg"], _C["tally_hover"], lambda: self._show_category("tally")),
        ]

        current_row_frame = None
        for i, (tools, icon, label, desc, acc, normal_fg, hover_fg, callback) in enumerate(categories):
            if i % 3 == 0:
                current_row_frame = ctk.CTkFrame(wrapper, fg_color="transparent")
                current_row_frame.pack(pady=(0 if i == 0 else 18, 0))

            card = self._make_category_card(
                parent=current_row_frame,
                icon=icon,
                label=label,
                sub_text=_sub(tools),
                desc=desc,
                acc=acc,
                normal_fg=normal_fg,
                hover_fg=hover_fg,
                callback=callback,
            )
            card.pack(side="left", padx=14)

        return page

    def _make_category_card(self, parent, icon, label, sub_text, desc,
                             acc, normal_fg, hover_fg, callback):
        """Build a large clickable category card."""
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

        # Icon
        ctk.CTkLabel(card, text=icon,
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
        ctk.CTkLabel(card, text="Click to explore  →",
                     font=("Segoe UI", 11, "bold"),
                     text_color=acc).pack(pady=(4, 0))

        # Bind hover + click to the entire card (including all children)
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

        # Tool tabs (lazy)
        for t in tools:
            tv.add(t["tab"])
            self._loaded[t["tab"]] = False

        tv.pack(fill="both", expand=True, padx=10, pady=(4, 4))

        self._tabviews[key]   = tv
        self._last_tabs[key]  = ""
        self._cat_frames[key] = frame
        return frame

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
            ac    = accents[idx] if idx < len(accents) else acc
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

            ctk.CTkLabel(body, text=tool["tab"],
                         font=("Segoe UI", 14, "bold"),
                         text_color=ac).pack(anchor="w")
            ctk.CTkFrame(body, height=1, corner_radius=0,
                         fg_color=_C["border"]).pack(fill="x", pady=(5, 8))
            ctk.CTkLabel(body, text=tool["desc"],
                         font=("Segoe UI", 12),
                         text_color=("#1e293b", "#e2e8f0"),
                         wraplength=230, justify="left").pack(anchor="w")
            ctk.CTkLabel(body, text="→  Click to open",
                         font=("Segoe UI", 11, "bold"),
                         text_color=ac).pack(anchor="w", pady=(10, 0))

            # Make card clickable — switches directly to the tool's tab
            if tv is not None:
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
        self._active_poll = None
        for f in self._cat_frames.values():
            f.pack_forget()
        self._landing.pack(fill="both", expand=True)
        self._refresh_header_left("landing")

    def _show_category(self, key: str):
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
        self.after(300, lambda: self._poll_tab(key))

    def _go_home(self):
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
        if self._active_poll != key:
            return   # category no longer active
        tv = self._tabviews.get(key)
        if not tv:
            self.after(300, lambda: self._poll_tab(key))
            return
        try:
            current = tv.get()
        except Exception:
            self.after(300, lambda: self._poll_tab(key))
            return

        last = self._last_tabs.get(key, "")
        if current != last:
            self._last_tabs[key] = current
            if current != "🏠  Overview" and not self._loaded.get(current, True):
                self._activate_tab(key, current)

        self.after(300, lambda: self._poll_tab(key))

    def _activate_tab(self, cat_key: str, name: str):
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
        idx    = next((i for i, t in enumerate(tools) if t["tab"] == name), 0)
        accent = accents[idx] if idx < len(accents) else _C["primary"]

        tab_frame = self._tabviews[cat_key].tab(name)

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
            if overlay.winfo_exists():
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
                self.after(80, _animate_spinner)

        # Start spinner animation
        self.after(80, _animate_spinner)

        # Defer actual loading so the overlay is visible
        def _do_load():
            try:
                if tool.get("tk"):
                    inst = _load_tk_tool(tab_frame, tool["module"], tool["class"])
                else:
                    inst = _load_tool(tab_frame, tool["module"], tool["class"])
                self._instances[name] = inst
                self._loaded[name]    = True

                # Sync theme immediately after load
                if inst and hasattr(inst, "set_theme"):
                    try: inst.set_theme(self._current_theme)
                    except: pass
            except Exception as e:
                import traceback as _tb
                err_msg = _tb.format_exc()
                print(f"[ERROR] Failed to load tool '{name}': {err_msg}")
                # Show error in the tab so it's visible (not a silent blank)
                try:
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
                # Remove overlay after loading completes (or fails)
                try: overlay.destroy()
                except: pass

        self.after(100, _do_load)


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = GSTSuite()
    app.mainloop()
