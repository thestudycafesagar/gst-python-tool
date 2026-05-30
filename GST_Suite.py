"""
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
import tkinter.ttk as _ttk

_SPLASH = None  # no splash screen

_BOOT_SPLASH = None
_BOOT_STATUS_VAR = None
_BOOT_PROGRESS = None

try:
    import pyi_splash as _pyi_splash
except Exception:
    _pyi_splash = None


def _boot_assets_base():
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _update_native_bootloader_splash(message: str):
    if _pyi_splash is None:
        return
    try:
        _pyi_splash.update_text(message)
    except Exception:
        pass


def _close_native_bootloader_splash():
    if _pyi_splash is None:
        return
    try:
        _pyi_splash.close()
    except Exception:
        pass


def _show_boot_splash():
    global _BOOT_SPLASH, _BOOT_STATUS_VAR, _BOOT_PROGRESS
    if _BOOT_SPLASH is not None:
        return
    try:
        splash = _tk.Tk()
        splash.overrideredirect(True)
        splash.configure(bg="#0b1220")
        splash.attributes("-topmost", True)

        width, height = 500, 280
        sw = splash.winfo_screenwidth()
        sh = splash.winfo_screenheight()
        x = max((sw - width) // 2, 0)
        y = max((sh - height) // 2, 0)
        splash.geometry(f"{width}x{height}+{x}+{y}")

        panel = _tk.Frame(
            splash,
            bg="#111827",
            bd=0,
            highlightthickness=1,
            highlightbackground="#1f2937",
        )
        panel.pack(fill="both", expand=True, padx=10, pady=10)

        content = _tk.Frame(panel, bg="#111827")
        content.pack(fill="both", expand=True)

        logo_path = os.path.join(_boot_assets_base(), "studycafelogo.png")
        if os.path.exists(logo_path):
            try:
                logo_img = None
                try:
                    from PIL import Image as _SplashImage, ImageTk as _SplashImageTk

                    pil = _SplashImage.open(logo_path).convert("RGBA")
                    resample = getattr(_SplashImage, "LANCZOS", _SplashImage.BICUBIC)
                    pil.thumbnail((250, 110), resample)
                    logo_img = _SplashImageTk.PhotoImage(pil)
                except Exception:
                    fallback = _tk.PhotoImage(file=logo_path)
                    scale_x = max(fallback.width() // 250, 1)
                    scale_y = max(fallback.height() // 110, 1)
                    logo_img = fallback.subsample(scale_x, scale_y)

                content._logo_img = logo_img
                _tk.Label(content, image=logo_img, text="", bg="#111827").pack(pady=(18, 8))
            except Exception:
                pass

        _tk.Label(
            content,
            text="AutomationCafe Suite",
            font=("Segoe UI", 16, "bold"),
            fg="#e2e8f0",
            bg="#111827",
        ).pack()

        _tk.Label(
            content,
            text="GST & Income Tax Automation Suite",
            font=("Segoe UI", 10),
            fg="#94a3b8",
            bg="#111827",
        ).pack(pady=(2, 10))

        _tk.Label(
            content,
            text="Loading, please wait...",
            font=("Segoe UI", 10, "bold"),
            fg="#cbd5e1",
            bg="#111827",
        ).pack(pady=(0, 2))

        _BOOT_STATUS_VAR = _tk.StringVar(value="Starting...")
        _tk.Label(
            content,
            textvariable=_BOOT_STATUS_VAR,
            font=("Segoe UI", 10),
            fg="#93c5fd",
            bg="#111827",
            wraplength=440,
            justify="center",
        ).pack(pady=(0, 14))

        _BOOT_PROGRESS = _ttk.Progressbar(content, mode="indeterminate", length=330)
        _BOOT_PROGRESS.pack(pady=(0, 14))
        _BOOT_PROGRESS.start(10)

        splash.update_idletasks()
        splash.update()
        _BOOT_SPLASH = splash
        _close_native_bootloader_splash()
    except Exception:
        _BOOT_SPLASH = None
        _BOOT_STATUS_VAR = None
        _BOOT_PROGRESS = None


def _update_boot_splash(message: str):
    _update_native_bootloader_splash(message)
    if _BOOT_SPLASH is None or _BOOT_STATUS_VAR is None:
        return
    try:
        _BOOT_STATUS_VAR.set(message)
        _BOOT_SPLASH.update_idletasks()
        _BOOT_SPLASH.update()
    except Exception:
        pass


def _close_boot_splash():
    global _BOOT_SPLASH, _BOOT_STATUS_VAR, _BOOT_PROGRESS
    _close_native_bootloader_splash()
    if _BOOT_SPLASH is None:
        return
    splash_ref = _BOOT_SPLASH
    if _BOOT_PROGRESS is not None:
        try:
            _BOOT_PROGRESS.stop()
        except Exception:
            pass
    try:
        _BOOT_SPLASH.destroy()
    except Exception:
        pass

    # Avoid image handle issues by clearing tkinter's default root reference.
    try:
        if getattr(_tk, "_default_root", None) is splash_ref:
            _tk._default_root = None
    except Exception:
        pass

    _BOOT_SPLASH = None
    _BOOT_STATUS_VAR = None
    _BOOT_PROGRESS = None


_show_boot_splash()
_update_boot_splash("Loading UI engine...")

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
VERSION            = "1.0.13"
# !! REPLACE 'YOURNAME' and 'YOURREPO' with your actual GitHub username and
#    the public releases repo you created (e.g. gst-suite-releases).
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/Mr-RohitNooB/gst-suite-releases/main/latest.json"

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

_update_boot_splash("Preparing login screen...")

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
REGISTER_URL  = "https://register.automationcafe.in/Home/Register"

def _get_hardware_id():
    mac = uuid.getnode()
    return ':'.join(['{:02x}'.format((mac >> i) & 0xff) for i in range(0, 48, 8)][::-1])

def _save_auth(email, password):
    with open(_AUTH_CONFIG, 'w') as f:
        json.dump({"email": email, "password": password}, f)

def _load_auth():
    """Return saved (email, password) tuple or (None, None) if not found."""
    try:
        with open(_AUTH_CONFIG, 'r') as f:
            data = json.load(f)
        return data.get("email"), data.get("password")
    except Exception:
        return None, None

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
    # ── When running as installed EXE, redirect the working directory to
    # the user's Documents folder so all downloaded files (GST Downloaded,
    # Income Tax Downloaded, etc.) are saved there instead of the hidden
    # AppData / temp extraction folder.
    _docs = os.path.join(os.path.expanduser("~"), "Documents", "AutomationCafe Downloads")
    os.makedirs(_docs, exist_ok=True)
    os.chdir(_docs)
else:
    _ASSETS_BASE = os.path.dirname(os.path.abspath(__file__))
    _BASE        = _ASSETS_BASE
    
    # ── Redirect in development mode as well for consistent download locations
    _docs = os.path.join(os.path.expanduser("~"), "Documents", "AutomationCafe Downloads")
    os.makedirs(_docs, exist_ok=True)
    os.chdir(_docs)

_GST_BASE   = os.path.join(_BASE, "GST")
_IT_BASE    = os.path.join(_BASE, "Income Tax")
_PDF_BASE   = os.path.join(_BASE, "PDF_Utilities")
_BANK_BASE  = os.path.join(_BASE, "Bank Statement To Excel")
_GMAIL_BASE = os.path.join(_BASE, "Gmail-Tools")
_EMAIL_BASE = os.path.join(_BASE, "Outlook Email Tools")
_RECO_BASE  = os.path.join(_BASE, "GST_RECO")
_TALLY_BASE = os.path.join(_BASE, "tally tool")


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
    def minsize(self, *a, **k):    pass
    def maxsize(self, *a, **k):    pass
    def state(self, *a, **k):      return "normal"
    def withdraw(self):            pass
    def deiconify(self):           pass
    def option_add(self, *a, **k): pass
    def report_callback_exception(self, *a, **k): pass
    def _set_appearance_mode(self, *a, **k): pass


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
#  _ScrollableTabview  – horizontally scrollable ribbon replaces CTkTabview
# ══════════════════════════════════════════════════════════════════════════════
class _ScrollableTabview:
    """Drop-in for ctk.CTkTabview with a canvas-based scrollable tab ribbon.

    API surface used by the suite:
        .add(name)  .tab(name)  .get()  .set(name)  .pack(**kw)
        ._tab_dict  ._segmented_button (proxy)  ._fg_frame (proxy)
    """

    class _Proxy:
        def configure(self, **kw): pass
        def bind(self, *a, **kw):  pass
        def winfo_children(self):  return []

    def __init__(self, parent, accent_color=None, **kwargs):
        fg_color     = kwargs.get("fg_color",     ("#ffffff", "#1e293b"))
        border_color = kwargs.get("border_color",  ("#e2e8f0", "#334155"))
        border_width = int(kwargs.get("border_width", 2))

        self._outer = ctk.CTkFrame(parent, fg_color=fg_color,
                                   border_color=border_color,
                                   border_width=border_width,
                                   corner_radius=10)
        self._acc = accent_color
        acc_c = accent_color if accent_color else ("#4f46e5", "#6366f1")

        # Top accent stripe (3 px colour bar)
        ctk.CTkFrame(self._outer, height=3, fg_color=acc_c,
                     corner_radius=0).pack(fill="x", side="top")

        # ── Ribbon row (fixed 46 px, slightly darker surface) ─────────────────
        _mode = ctk.get_appearance_mode()
        _ribbon_bg  = ("#eef2f7", "#141c2e") if _mode == "Dark" else ("#eef2f7", "#141c2e")
        _canvas_bg  = "#141c2e" if _mode == "Dark" else "#eef2f7"

        ribbon_row = ctk.CTkFrame(self._outer,
                                  fg_color=_ribbon_bg,
                                  corner_radius=0, height=46)
        ribbon_row.pack(fill="x", side="top")
        ribbon_row.pack_propagate(False)

        # ── Left / Right arrow buttons (no native scrollbar) ─────────────────
        _arrow_kw = dict(
            width=30, height=46,
            fg_color="transparent",
            hover_color=("#d1d9e6", "#1e293b"),
            font=("Segoe UI", 18, "bold"),
            corner_radius=0,
        )
        self._btn_left = ctk.CTkButton(
            ribbon_row, text="‹",
            text_color=("#94a3b8", "#475569"),
            command=self._scroll_left, **_arrow_kw)
        self._btn_left.pack(side="left")

        self._btn_right = ctk.CTkButton(
            ribbon_row, text="›",
            text_color=("#94a3b8", "#475569"),
            command=self._scroll_right, **_arrow_kw)
        self._btn_right.pack(side="right")

        # ── Canvas (no scrollbar widget — arrows only) ────────────────────────
        self._ribbon_canvas = _tk.Canvas(ribbon_row, height=46,
                                          bg=_canvas_bg,
                                          highlightthickness=0, bd=0)
        self._ribbon_canvas.pack(side="left", fill="both", expand=True)

        self._ribbon_inner = ctk.CTkFrame(self._ribbon_canvas,
                                           fg_color="transparent")
        self._ribbon_canvas.create_window(0, 0, anchor="nw",
                                           window=self._ribbon_inner)

        self._overflow = False   # True only when tabs exceed canvas width

        def _update_sr(_=None):
            bb = self._ribbon_canvas.bbox("all")
            if not bb:
                return
            self._ribbon_canvas.configure(scrollregion=bb)
            total   = bb[2] - bb[0]
            visible = self._ribbon_canvas.winfo_width()
            self._overflow = total > visible + 4
            _bright = ("#334155", "#94a3b8")
            _dim    = ("#94a3b8", "#475569")
            self._btn_left.configure(text_color=_bright if self._overflow else _dim)
            self._btn_right.configure(text_color=_bright if self._overflow else _dim)

        self._ribbon_inner.bind("<Configure>", _update_sr)
        self._ribbon_canvas.bind("<Configure>", _update_sr)
        def _mw(ev):
            if self._overflow:
                self._ribbon_canvas.xview_scroll(-1 if ev.delta > 0 else 1, "units")

        self._ribbon_canvas.bind("<MouseWheel>", _mw)
        self._ribbon_inner.bind("<MouseWheel>", _mw)
        self._mw_handler = _mw

        # ── Thin separator ────────────────────────────────────────────────────
        ctk.CTkFrame(self._outer, height=1, fg_color=border_color,
                     corner_radius=0).pack(fill="x")

        # ── Content host ──────────────────────────────────────────────────────
        self._content_host = ctk.CTkFrame(self._outer, fg_color=fg_color,
                                           corner_radius=0)
        self._content_host.pack(fill="both", expand=True)
        self._content_host.grid_rowconfigure(0, weight=1)
        self._content_host.grid_columnconfigure(0, weight=1)

        self._tab_dict = {}
        self._btn_dict = {}
        self._current  = None
        self._segmented_button = self._Proxy()
        self._fg_frame         = self._Proxy()

    # ── Scroll helpers ────────────────────────────────────────────────────────

    def _scroll_left(self):
        if self._overflow:
            self._ribbon_canvas.xview_scroll(-3, "units")

    def _scroll_right(self):
        if self._overflow:
            self._ribbon_canvas.xview_scroll(3, "units")

    # ── Public tab API ────────────────────────────────────────────────────────

    def add(self, name: str):
        frame = ctk.CTkFrame(self._content_host, fg_color="transparent",
                              corner_radius=0)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_remove()
        self._tab_dict[name] = frame

        is_first = not self._btn_dict
        acc      = self._acc or ("#4f46e5", "#6366f1")
        btn_fg   = acc if is_first else "transparent"
        btn_tc   = "#ffffff" if is_first else ("#475569", "#94a3b8")

        btn = ctk.CTkButton(
            self._ribbon_inner, text=name, height=34,
            font=("Segoe UI", 12, "bold"),
            fg_color=btn_fg, text_color=btn_tc,
            hover_color=("#dde3ec", "#1e293b"),
            corner_radius=7,
            command=lambda n=name: self.set(n),
        )
        btn.pack(side="left", padx=(4, 0), pady=6)
        try:
            btn.bind("<MouseWheel>", self._mw_handler)
        except Exception:
            pass
        self._btn_dict[name] = btn

        if is_first:
            self._current = name
            frame.grid()

        return frame

    def tab(self, name: str):
        return self._tab_dict.get(name)

    def get(self) -> str:
        return self._current or ""

    def set(self, name: str):
        if name == self._current:
            return
        if self._current:
            if self._current in self._btn_dict:
                self._btn_dict[self._current].configure(
                    fg_color="transparent",
                    text_color=("#475569", "#94a3b8"))
            if self._current in self._tab_dict:
                self._tab_dict[self._current].grid_remove()

        self._current = name
        acc = self._acc or ("#4f46e5", "#6366f1")
        if name in self._btn_dict:
            self._btn_dict[name].configure(fg_color=acc, text_color="#ffffff")
        if name in self._tab_dict:
            self._tab_dict[name].grid()

        # Auto-scroll ribbon so active button stays visible
        try:
            btn = self._btn_dict.get(name)
            if btn:
                btn.update_idletasks()
                bx  = btn.winfo_x()
                bw  = btn.winfo_width()
                cw  = self._ribbon_canvas.winfo_width()
                bb  = self._ribbon_canvas.bbox("all")
                if bb and bb[2] > 0:
                    frac = max(0.0, min(1.0, (bx + bw / 2 - cw / 2) / bb[2]))
                    self._ribbon_canvas.xview_moveto(frac)
        except Exception:
            pass

    # ── Geometry delegated to outer frame ─────────────────────────────────────

    def pack(self, **kw):   self._outer.pack(**kw)
    def grid(self, **kw):   self._outer.grid(**kw)
    def place(self, **kw):  self._outer.place(**kw)

    def winfo_exists(self):
        try:    return bool(self._outer.winfo_exists())
        except: return False


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
    # Gmail Tools category
    "gmail_acc":   ("#ea4335", "#f28b82"),
    "gmail_bg":    ("#fce8e6", "#2f1b1a"),
    "gmail_hover": ("#fad2cf", "#472725"),
    # Combined Email Tools category
    "mail_acc":    ("#7c3aed", "#a78bfa"),
    "mail_bg":     ("#ede9fe", "#1a0f2e"),
    "mail_hover":  ("#ddd6fe", "#240f40"),
    # GST Reconciliation category
    "reco_acc":    ("#0f766e", "#2dd4bf"),
    "reco_bg":     ("#ccfbf1", "#042f2e"),
    "reco_hover":  ("#99f6e4", "#064e3b"),
    # Tally Automation category
    "tally_acc":   ("#7c2d12", "#fb923c"),
    "tally_bg":    ("#ffedd5", "#2a1608"),
    "tally_hover": ("#fed7aa", "#3a1d0a"),
}

# Per-tool accent colours (light, dark)
_GST_ACCENTS = [
    ("#059669", "#34d399"),   # GST Verifier      Emerald
    ("#0284c7", "#38bdf8"),   # R1 JSON           Sky
    ("#db2777", "#f472b6"),   # JSON→Excel        Pink
    ("#dc2626", "#f87171"),   # R1 PDF            Red
    ("#0f766e", "#2dd4bf"),   # IMS               Teal
    ("#4f46e5", "#818cf8"),   # GSTR-2B           Indigo
    ("#7c3aed", "#a78bfa"),   # GSTR-3B           Violet
    ("#0891b2", "#22d3ee"),   # 3B→Excel          Cyan
    ("#d97706", "#fbbf24"),   # Challan              Amber
    ("#8b5cf6", "#c4b5fd"),   # GST Reports          Purple
    ("#0f766e", "#2dd4bf"),   # GST Reco             Teal
    ("#6366f1", "#818cf8"),   # GST Reconciliation   Indigo
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
_GMAIL_ACCENTS = [
    ("#ea4335", "#f28b82"),   # GST Return     Red
    ("#4285f4", "#8ab4f8"),   # Invoice Sender Blue
    ("#34a853", "#81c995"),   # Payment Remind Green
]
_COMBINED_EMAIL_ACCENTS = [
    ("#d97706", "#fbbf24"),   # Outlook GST Return     Amber
    ("#0284c7", "#38bdf8"),   # Outlook Invoice Sender Sky
    ("#059669", "#34d399"),   # Outlook Payment Remind Emerald
    ("#ea4335", "#f28b82"),   # Gmail GST Return       Red
    ("#4285f4", "#8ab4f8"),   # Gmail Invoice Sender   Blue
    ("#34a853", "#81c995"),   # Gmail Payment Remind   Green
]
_MAIL_GROUP_ACCENTS = [
    ("#d97706", "#fbbf24"),   # Outlook suite
    ("#ea4335", "#f28b82"),   # Gmail suite
]
_TALLY_ACCENTS = [
    ("#7c2d12", "#fb923c"),   # Tally Automation Orange
]


# ══════════════════════════════════════════════════════════════════════════════
#  TOOL REGISTRY
# ══════════════════════════════════════════════════════════════════════════════
GST_TOOLS = [
    {"key": "GST_Verifier", "tab": "🤖  GST Verifier", "module": os.path.join(_GST_BASE, "GST Bot",               "gst_pro_app.py"), "class": "GSTApp",             "desc": "Verify bulk GSTINs and extract filing history & details."},
    {"key": "R1_JSON",      "tab": "📑  GSTR1 JSON",      "module": os.path.join(_GST_BASE, "GST R1 Downloader",     "mai.py"),         "class": "App",                "desc": "Request or download GSTR-1 JSON files for multiple users."},
    {"key": "JSON_Excel",   "tab": "📊  GSTR1 Json to Excel", "module": os.path.join(_GST_BASE, "JSON to Excel",          "main.py"),        "class": "App",                "desc": "Convert GSTR-1 JSON exports to multi-sheet Excel reports."},
    {"key": "R1_PDF",       "tab": "📄  GSTR1 PDF",       "module": os.path.join(_GST_BASE, "R1 PDF Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-1 PDF filed    returns from the GST portal."},
    {"key": "IMS",          "tab": "📲  IMS",          "module": os.path.join(_GST_BASE, "IMS Downloader",        "main.py"),        "class": "App",                "desc": "Download IMS (Invoice Management System) data in bulk from the GST portal."},
    {"key": "GSTR2B",       "tab": "📥  GSTR 2B",      "module": os.path.join(_GST_BASE, "GST 2B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-2B returns via automated browser."},
    {"key": "GSTR3B",       "tab": "📥  GSTR 3B",      "module": os.path.join(_GST_BASE, "GST 3B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-3B returns via automated browser."},
    {"key": "GSTR3B_Excel", "tab": "📊  GSTR 3B to Excel",   "module": os.path.join(_GST_BASE, "GST 3B to Excel",       "main.py"),        "class": "GSTR3BConverterPro", "desc": "Convert GSTR-3B PDF files to formatted Excel sheets."},
    {"key": "GST_Challan",  "tab": "💰  GST Challan",      "module": os.path.join(_GST_BASE, "GST Challan Downloader","main.py"),        "class": "App",                "desc": "Download GST Challan PDFs in bulk (Monthly / Quarterly)."},
    {"key": "GST_Reports",  "tab": "📊  GST Reports",      "module": os.path.join(_GST_BASE, "GST_Reports",           "gst_portal_gui.py"), "class": "GstPortalApp",       "desc": "Download GSTR-1, GSTR-2A, GSTR-2B and GSTR-3B returns directly from the GST portal with login, OTP, Excel export and consolidated yearly downloads."},
    {"key": "GST_Reco",     "tab": "🔄  GST Reco",          "module": os.path.join(_RECO_BASE, "mainpy-reco-speqtra.py"), "class": "App", "tk": False, "desc": "Reconcile GSTR-2B portal data against Tally/books. Matches invoices, highlights mismatches and exports a detailed Excel report."},
    {"key": "GST_Reco_Cards", "tab": "⚖  GST Reconciliation", "builtin_ui": "reco_landing", "module": os.path.join(_RECO_BASE, "mainpy-reco-speqtra.py"), "class": "App", "tk": False, "desc": "Reconcile GST portal data against books — Sale vs GSTR-1, Purchase vs GSTR-2A/2B, 2A vs 2B and GSTR-9 reconciliation."},
]

IT_TOOLS = [
    {"key": "IT_26AS",         "tab": "📄  26AS/AIS & TIS",       "module": os.path.join(_IT_BASE, "26 AS Downlaoder",  "main.py"),                  "class": "App",              "desc": "Download 26AS / AIS / TIS reports in bulk."},
    {"key": "IT_Challan",      "tab": "💰  Challan Downloader", "module": os.path.join(_IT_BASE, "Challan Downloader","main.py"),                  "class": "App",              "desc": "Download Income Tax Challan PDFs in bulk."},
    {"key": "ITR_Bot",         "tab": "🤖  ITR Bot",            "module": os.path.join(_IT_BASE, "ITR - Bot",         "GUI_based_app.py"),         "class": "App",              "desc": "Automate ITR filing workflows with the ITR bot."},
    {"key": "Demand_Checker",  "tab": "🔍  Demand Checker",     "module": os.path.join(_IT_BASE, "Challan Downloader","demand_checker_app.py"),    "class": "DemandCheckerApp", "desc": "Check pending worklist and outstanding demands in bulk from the Income Tax portal."},
    {"key": "Refund_Checker",  "tab": "📊  Refund Checker",     "module": os.path.join(_IT_BASE, "26 AS Downlaoder",  "refund_checker_app.py"),    "class": "RefundCheckerApp", "desc": "Extract filed return data and generate refund status reports in bulk."},
]

PDF_TOOLS = [
    {"key": "PDF_Merge",    "tab": "⊕  Merge",    "module": os.path.join(_PDF_BASE, "main.py"), "class": "MergeApp",    "tk": True, "desc": "Merge multiple PDF files into one high-quality document."},
    {"key": "PDF_Split",    "tab": "✂  Split",    "module": os.path.join(_PDF_BASE, "main.py"), "class": "SplitApp",    "tk": True, "desc": "Split PDF files into smaller parts by page range or every N pages."},
    {"key": "PDF_Extract",  "tab": "⊙  Extract",  "module": os.path.join(_PDF_BASE, "main.py"), "class": "ExtractApp",  "tk": True, "desc": "Extract specific pages from a PDF document to a new file."},
    {"key": "PDF_Redact",   "tab": "⬛  Redact",   "module": os.path.join(_PDF_BASE, "main.py"), "class": "RedactApp",   "tk": True, "desc": "Securely black out sensitive information and text from PDF documents."},
]

BANK_TOOLS = [
    {"key": "Bank_Excel", "tab": "🏦  Bank → Excel", "module": os.path.join(_BANK_BASE, "bank_to_excel.py"), "class": "App", "tk": False, "desc": "Convert bank statement PDFs to formatted Excel sheets. Supports HDFC, ICICI, SBI, Axis, Kotak, IDFC, BOI, Yes, UCO & Equitas."},
]

EMAIL_TOOLS = [
    {"key": "Email_GST_Request", "tab": "📋  GST Return Request", "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "GSTReturnMailApp",   "tk": False, "desc": "Send bulk GST return data request emails via Outlook. Auto-fills month, return type, deadlines and contact details."},
    {"key": "Email_Invoice",     "tab": "🧾  Invoice Sender",     "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "InvoiceSenderMailApp",   "tk": False, "desc": "Dispatch personalised invoices to clients in bulk via Outlook. Supports per-row service, period, amount & PDF attachments."},
    {"key": "Email_Payment",     "tab": "💰  Payment Reminder",   "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "PaymentReminderMailApp", "tk": False, "desc": "Send outstanding payment reminder emails in bulk via Outlook. Includes interest clause, deadline and per-client amounts."},
    # {"key": "Email_Custom",      "tab": "✏  Custom Email",        "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "CustomMailApp",          "tk": False, "desc": "Build and send fully custom bulk emails via Outlook. Define your own subject, body with {placeholders}, save multiple templates, and generate dynamic Excel recipient sheets."},
]

GMAIL_TOOLS = [
    {"key": "Gmail_GST_Request", "tab": "📋  GST Return Request", "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "GSTReturnMailApp",   "tk": False, "desc": "Send bulk GST return data request emails via Gmail. Auto-fills month, return type, deadlines and contact details."},
    {"key": "Gmail_Invoice",     "tab": "🧾  Invoice Sender",     "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "InvoiceSenderMailApp",   "tk": False, "desc": "Dispatch personalised invoices to clients in bulk via Gmail. Supports per-row service, period, amount & PDF attachments."},
    {"key": "Gmail_Payment",     "tab": "💰  Payment Reminder",   "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "PaymentReminderMailApp", "tk": False, "desc": "Send outstanding payment reminder emails in bulk via Gmail. Includes interest clause, deadline and per-client amounts."},
    # {"key": "Gmail_Custom",      "tab": "✏  Custom Email",        "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "CustomMailApp",          "tk": False, "desc": "Build and send fully custom bulk emails via Gmail. Define your own subject, body with {placeholders}, save multiple templates, and generate dynamic Excel recipient sheets."},
]

COMBINED_EMAIL_TOOLS = [
    {"key": "Email_GST_Request", "tab": "📧  Outlook | GST Request",    "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "GSTReturnMailApp",   "tk": False, "desc": "Send bulk GST return data request emails via Outlook. Auto-fills month, return type, deadlines and contact details."},
    {"key": "Email_Invoice",     "tab": "📧  Outlook | Invoice Sender",  "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "InvoiceSenderMailApp",   "tk": False, "desc": "Dispatch personalised invoices to clients in bulk via Outlook. Supports per-row service, period, amount & PDF attachments."},
    {"key": "Email_Payment",     "tab": "📧  Outlook | Payment Reminder","module": os.path.join(_EMAIL_BASE, "main.py"), "class": "PaymentReminderMailApp", "tk": False, "desc": "Send outstanding payment reminder emails in bulk via Outlook. Includes interest clause, deadline and per-client amounts."},
    # {"key": "Email_Custom",      "tab": "📧  Outlook | Custom Email",    "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "CustomMailApp",          "tk": False, "desc": "Build and send fully custom bulk emails via Outlook with {placeholder} templates and dynamic Excel sheets."},
    {"key": "Gmail_GST_Request", "tab": "✉  Gmail | GST Request",       "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "GSTReturnMailApp",   "tk": False, "desc": "Send bulk GST return data request emails via Gmail. Auto-fills month, return type, deadlines and contact details."},
    {"key": "Gmail_Invoice",     "tab": "✉  Gmail | Invoice Sender",    "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "InvoiceSenderMailApp",   "tk": False, "desc": "Dispatch personalised invoices to clients in bulk via Gmail. Supports per-row service, period, amount & PDF attachments."},
    {"key": "Gmail_Payment",     "tab": "✉  Gmail | Payment Reminder",  "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "PaymentReminderMailApp", "tk": False, "desc": "Send outstanding payment reminder emails in bulk via Gmail. Includes interest clause, deadline and per-client amounts."},
    # {"key": "Gmail_Custom",      "tab": "✉  Gmail | Custom Email",      "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "CustomMailApp",          "tk": False, "desc": "Build and send fully custom bulk emails via Gmail with {placeholder} templates and dynamic Excel sheets."},
]

MAIL_GROUP_TOOLS = [
    {"key": "Email_Suite", "tab": "📧  Outlook Email Tools", "module": os.path.join(_EMAIL_BASE, "main.py"), "class": "BulkMailApp", "tk": False, "desc": "Outlook suite with 3 built-in templates: GST Return Request, Invoice Sender, Payment Reminder.", "is_card_only": True, "action_cat": "email"},
    {"key": "Gmail_Suite", "tab": "✉  Gmail Email Tools",    "module": os.path.join(_GMAIL_BASE, "main.py"), "class": "BulkMailApp", "tk": False, "desc": "Gmail suite with 3 built-in templates: GST Return Request, Invoice Sender, Payment Reminder.", "is_card_only": True, "action_cat": "gmail"},
]

TALLY_TOOLS = [                 
    {"key": "Tally_Automation", "tab": "🧾  GSTR-2B → Tally", "module": os.path.join(_TALLY_BASE, "main.py"), "class": "GSTR2BTallyApp", "desc": "Convert GSTR-2B and Tally sheets into Tally-ready outputs with XML generation, mapping and automation helpers."},
    {"key": "Tally_Bank", "tab": "🏦  Bank Statement → Tally", "module": os.path.join(_TALLY_BASE, "Bank_Statment_to_Tally.py"), "class": "TallyBankApp", "desc": "Convert bank statement Excel to Tally Payment/Receipt vouchers and push XML directly to Tally."},
    {"key": "Tally_Sales", "tab": "🛒  Tally Entry", "module": os.path.join(_TALLY_BASE, "sale_purchase_entry.py"), "class": "TallySalesApp", "desc": "Automate sales and purchase entries in Tally."},
    # {"key": "Tally_Credit_Debit_Note", "tab": "📝  Credit/Debit Note", "module": os.path.join(_TALLY_BASE, "credit-debit-note.py"), "class": "TallyNoteEntryApp", "desc": "Create Credit/Debit Note vouchers from Excel or manual entry, then export or push XML directly to TallyPrime."},
    # {"key": "Tally_Journal", "tab": "📒  Journal Entry", "module": os.path.join(_TALLY_BASE, "journal_entry.py"), "class": "TallyJournalApp", "desc": "Create Journal vouchers from Excel upload or manual entry, with XML export and direct push to TallyPrime."},
]
RECO_TOOLS = []   # placeholder — reconciliation lives inside GST_Reco tab now
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
        "requests",
        "PIL",
    )

    def __init__(self, user_info: dict = None):
        super().__init__()
        self.title("GST & Income Tax Automation Suite")
        self.geometry("1200x920")
        self.minsize(980, 760)
        self.resizable(True, True)
        self.after(10, lambda: self.state("zoomed"))

        # Set taskbar / window icon
        _ico = os.path.join(_ASSETS_BASE, "studycafelogo.ico")
        if os.path.exists(_ico):
            self.iconbitmap(_ico)

        self._user_info      = user_info or {}
        _raw_allowed = (user_info or {}).get("allowed_modules") or (user_info or {}).get("CustomAllowedModules")
        if isinstance(_raw_allowed, str):
            import json, re
            try:
                parsed = json.loads(_raw_allowed)
                if isinstance(parsed, list):
                    _raw_allowed = parsed
                else:
                    raise ValueError("Not a list")
            except Exception:
                # Ultra-robust fallback: if the API sends malformed JSON or a single string, strip and split by comma
                clean_str = re.sub(r'[\[\]"\'\s]', '', _raw_allowed)
                _raw_allowed = [x for x in clean_str.split(',') if x]
        
        self._allowed = set(_raw_allowed) if _raw_allowed is not None else None
        self._allowed_norm = (
            {str(m).strip().upper() for m in self._allowed}
            if self._allowed is not None else None
        )
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
        msg = f"New update found (v{latest}).\nInstalling the update..."

        _tk.messagebox.showinfo(
            title="Update Installing",
            message=msg,
            icon="info",
        )
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
                "Please re-download AutomationCafeSuite.exe from the releases page.",
            )
            return

        tmp_dir     = tempfile.mkdtemp(prefix="automationcafe_upd_")
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

        # ── Powered by AutomationCafe Suite logo ────────────────────────────────────────
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
            ctk.CTkLabel(logo_frame, text="AutomationCafe Suite",
                         font=("Segoe UI", 11, "bold"),
                         text_color="#818cf8").pack(side="left", padx=(6, 0))
        else:
            ctk.CTkLabel(logo_frame, text="Powered by",
                         font=("Segoe UI", 10),
                         text_color="#94a3b8").pack(side="left", padx=(0, 6))
            ctk.CTkLabel(logo_frame, text="AutomationCafe Suite",
                         font=("Segoe UI", 13, "bold"),
                         text_color="#818cf8").pack(side="left")

        self._tick_clock()

    def _tick_clock(self):
        if self._is_closing:
            return
        self._clock_lbl.configure(
            text=datetime.now().strftime("%d %b   %H:%M:%S"))
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

    # Tool keys that are always accessible regardless of plan
    _ALWAYS_FREE = {"Email_Custom", "Gmail_Custom", "GST_Reco", "GST_Reco_Cards"}

    def _is_tool_allowed(self, tool_key: str) -> bool:
        if self._allowed is None:
            return True

        key = str(tool_key or "").strip()
        if not key:
            return False

        if key in self._ALWAYS_FREE:
            return True

        if key in self._allowed:
            return True

        norm = self._allowed_norm or set()
        key_norm = key.upper()

        if key_norm in norm:
            return True

        # Backward-compatible aliases from older auth payloads.
        aliases = {
            "Tally_Automation": {"TALLY", "TALLY_TOOLS", "TALLY_AUTOMATION", "TALLY TOOL", "TALLY TOOLS"},
            "Email_Suite": {
                "EMAIL", "EMAIL_TOOLS", "OUTLOOK", "OUTLOOK TOOLS", "OUTLOOK EMAIL TOOLS",
                "EMAIL_GST_REQUEST", "EMAIL_INVOICE", "EMAIL_PAYMENT",
            },
            "Gmail_Suite": {
                "GMAIL", "GMAIL_TOOLS", "GMAIL TOOLS",
                "GMAIL_GST_REQUEST", "GMAIL_INVOICE", "GMAIL_PAYMENT",
            },
            "Email_Custom": {
                "EMAIL", "EMAIL_TOOLS", "OUTLOOK", "OUTLOOK TOOLS", "OUTLOOK EMAIL TOOLS",
                "EMAIL_CUSTOM", "CUSTOM_EMAIL", "EMAIL_GST_REQUEST", "EMAIL_INVOICE", "EMAIL_PAYMENT",
            },
            "Gmail_Custom": {
                "GMAIL", "GMAIL_TOOLS", "GMAIL TOOLS",
                "GMAIL_CUSTOM", "CUSTOM_EMAIL", "GMAIL_GST_REQUEST", "GMAIL_INVOICE", "GMAIL_PAYMENT",
            },
            "GST_Reco": {
                "GST_RECO", "RECO", "RECONCILIATION", "GST RECO", "GST_RECONCILIATION",
            },
        }
        if key in aliases and any(a in norm for a in aliases[key]):
            return True

        if any(token in norm for token in {"ALL", "*", "UNLIMITED", "FULL_ACCESS", "FULL"}):
            return True

        return False

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
                acc, icon, label, tools = _C["bank_acc"],  "🏦",  "Bank Statement → Excel (Beta)",      BANK_TOOLS
            elif cat_key == "mail":
                acc, icon, label, tools = _C["mail_acc"],  "📨",  "Email Tools",                 MAIL_GROUP_TOOLS
            elif cat_key == "email":
                acc, icon, label, tools = _C["email_acc"], "📧",  "Email Tools",                 EMAIL_TOOLS
            elif cat_key == "gmail":
                acc, icon, label, tools = _C["gmail_acc"], "✉",  "Gmail Tools",                 GMAIL_TOOLS
            elif cat_key == "reco":
                acc, icon, label, tools = _C["reco_acc"],  "🔄",  "GST Reconciliation",          RECO_TOOLS
            elif cat_key == "tally":
                acc, icon, label, tools = _C["tally_acc"], "🧾",  "Tally Automation Tools",      TALLY_TOOLS
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

        def _sub(tools):
            visible_tools = [t for t in tools if not t.get("hide_tab", False)]
            if getattr(self, '_allowed', None) is None:
                return f"{len(visible_tools)} module{'s' if len(visible_tools) != 1 else ''}"   
            n = sum(1 for t in visible_tools if self._is_tool_allowed(t.get("key")))    
            return f"{n} of {len(visible_tools)} modules"

        def _disabled(tools):
            visible_tools = [t for t in tools if not t.get("hide_tab", False)]
            if getattr(self, '_allowed', None) is None:
                return False
            return not any(self._is_tool_allowed(t.get("key")) for t in visible_tools)  
        # ── Dynamic 3x3 Grid Layout   ───────────────────────────────────────────
        categories = [
            (TALLY_TOOLS, "🧾", "Tally Automation Tools", "Convert GSTR-2B/Tally data\nto Tally-ready Excel and XML.", _C["tally_acc"], _C["tally_bg"], _C["tally_hover"], lambda: self._show_category("tally")),
            (GST_TOOLS, "🏛", "GST Tools", "Downloads, converters & verifiers\nfor GST portal automation.", _C["gst_acc"], _C["gst_bg"], _C["gst_hover"], lambda: self._show_category("gst")),
            (IT_TOOLS, "💼", "Income Tax Automation Suite", "26AS, Challan & ITR filing\nautomation tools.", _C["it_acc"], _C["it_bg"], _C["it_hover"], lambda: self._show_category("it")),
            (PDF_TOOLS, "📄", "PDF Tools", "Merge, split, extract, compress\n& redact PDF files.", _C["pdf_acc"], _C["pdf_bg"], _C["pdf_hover"], lambda: self._show_category("pdf")),
            # (BANK_TOOLS, "🏦", "Bank Statement → Excel (Beta)", "Convert bank statement PDFs\nto structured Excel sheets.", _C["bank_acc"], _C["bank_bg"], _C["bank_hover"], lambda: self._show_category("bank")),
            (COMBINED_EMAIL_TOOLS, "📨", "Email Tools", "Bulk personalised emails via Outlook & Gmail.\nGST reminders, invoices & more.", _C["mail_acc"], _C["mail_bg"], _C["mail_hover"], lambda: self._show_category("mail")),
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
                disabled=_disabled(tools)
            )
            card.pack(side="left", padx=14)

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
            ctk.CTkLabel(card, text="Module locked in current plan",
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
        elif key == "mail":
            tools, accents = MAIL_GROUP_TOOLS, _MAIL_GROUP_ACCENTS
        elif key == "email":
            tools, accents = EMAIL_TOOLS, _EMAIL_ACCENTS
        elif key == "gmail":
            tools, accents = GMAIL_TOOLS, _GMAIL_ACCENTS
        elif key == "reco":
            tools, accents = RECO_TOOLS,  _RECO_ACCENTS
        elif key == "tally":
            tools, accents = TALLY_TOOLS, _TALLY_ACCENTS
        else:
            tools, accents = IT_TOOLS,    _IT_ACCENTS

        frame = ctk.CTkFrame(self._content, fg_color="transparent",
                             corner_radius=0)

        # Determine category accent colour
        if key == "gst":
            acc_color = _C["gst_acc"]
        elif key == "pdf":
            acc_color = _C["pdf_acc"]
        elif key == "bank":
            acc_color = _C["bank_acc"]
        elif key == "mail":
            acc_color = _C["mail_acc"]
        elif key == "email":
            acc_color = _C["email_acc"]
        elif key == "gmail":
            acc_color = _C["gmail_acc"]
        elif key == "reco":
            acc_color = _C["reco_acc"]
        elif key == "tally":
            acc_color = _C["tally_acc"]
        else:
            acc_color = _C["it_acc"]

        # Build scrollable tabview — replaces CTkTabview; adds horizontal
        # ribbon scroll so tabs never collapse on smaller screens.
        tv = _ScrollableTabview(
            frame, accent_color=acc_color,
            fg_color=_C["surface"],
            border_color=_C["border"],
            border_width=2,
        )

        # Overview tab
        tv.add("🏠  Overview")
        self._build_category_overview(tv.tab("🏠  Overview"), key, tools, accents, tv)

        # Tool tabs (lazy; locked tabs get placeholder immediately)
        for t in tools:
            if t.get("is_card_only") or t.get("hide_tab", False):
                continue
            tv.add(t["tab"])
            if not self._is_tool_allowed(t.get("key")):
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
                     text="This module is not included in your current plan.\nContact AutomationCafe Suite to upgrade your subscription.",
                     font=("Segoe UI", 11), text_color=("#94a3b8", "#64748b"),
                     wraplength=300, justify="center").pack(pady=(10, 0))

    # ── Built-in UI dispatcher ────────────────────────────────────────────────

    def _build_builtin_ui(self, ui_key: str, tab_frame, accent):
        if ui_key == "reco_landing":
            self._build_reco_landing(tab_frame, accent)

    def _build_reco_landing(self, tab_frame, accent):
        """Arrow-card landing page — 2×2 grid, no header."""
        tab_frame.grid_rowconfigure(0, weight=1)
        tab_frame.grid_columnconfigure(0, weight=1)

        dark    = ctk.get_appearance_mode().lower() == "dark"
        area_bg = "#111827" if dark else "#f8fafc"

        outer = ctk.CTkFrame(tab_frame, fg_color=area_bg, corner_radius=0)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        holder = ctk.CTkFrame(outer, fg_color="transparent")
        holder.place(relx=0.5, rely=0.5, anchor="center")

        CARDS = [
            ("Sale Vs GSTR-1\nReconciliation",      "#f59e0b"),
            ("Purchase Vs. GSTR-2A\nReconciliation", "#10b981"),
            ("Purchase Vs. GSTR-2B\nReconciliation", "#e11d48"),
            ("2B Vs. 2A\nReconciliation",            "#7c3aed"),
        ]

        # 2 × 2 grid
        for i, (lbl, color) in enumerate(CARDS):
            r, c = divmod(i, 2)
            cv = self._make_reco_card(holder, lbl, color, area_bg, dark)
            cv.grid(row=r, column=c, padx=18, pady=18)

    def _make_reco_card(self, parent, label, color, area_bg, dark):
        """Draw a single arrow-shaped reconciliation card on a tk.Canvas."""
        W, H  = 490, 165
        arr   = 40    # right arrow depth
        cut   = 22    # left-corner diagonal cut
        bw    = 3     # border width
        m     = bw + 2

        card_bg  = "#1a2535" if dark else "#ffffff"
        text_col = "#e2e8f0" if dark else "#1e293b"
        sep_col  = "#334155" if dark else "#e5e7eb"
        dot_col  = "#4b5563" if dark else "#cbd5e1"

        cv = _tk.Canvas(parent, width=W, height=H,
                        bg=area_bg, highlightthickness=0, bd=0, cursor="hand2")

        # Outer polygon — border color fill
        cv.create_polygon(
            bw+cut, bw,   W-arr-bw, bw,   W-bw, H//2,
            W-arr-bw, H-bw,   bw+cut, H-bw,   bw, H-cut-bw,   bw, cut+bw,
            fill=color, outline="", smooth=False)

        # Inner polygon — card background
        cv.create_polygon(
            m+cut, m,   W-arr-m, m,   W-m-bw, H//2,
            W-arr-m, H-m,   m+cut, H-m,   m, H-cut-m,   m, cut+m,
            fill=card_bg, outline="", smooth=False)

        # Dashed circle (icon container)
        cx, cy, r = 88, H//2, 46
        segs = 14
        for i in range(segs):
            start = (360 / segs) * i + 90
            cv.create_arc(cx-r, cy-r, cx+r, cy+r,
                          start=start, extent=(360/segs)*0.55,
                          style="arc", outline=color, width=2)

        # Scales of justice icon
        cv.create_text(cx, cy, text="⚖", font=("Segoe UI Emoji", 26), fill=color)

        # Vertical separator
        sx = cx + r + 20
        cv.create_line(sx, 18, sx, H-18, fill=sep_col, width=1)

        # Title — two lines
        tx = sx + 20
        parts = label.split("\n")
        cv.create_text(tx, H//2 - (15 if len(parts) > 1 else 0),
                       text=parts[0], font=("Segoe UI", 14, "bold"),
                       fill=text_col, anchor="w")
        if len(parts) > 1:
            cv.create_text(tx, H//2 + 15, text=parts[1],
                           font=("Segoe UI", 14, "bold"),
                           fill=text_col, anchor="w")

        # Decorative footer: • • • ——
        dy, dx = H - 22, sx + 14
        for _ in range(3):
            cv.create_oval(dx-4, dy-4, dx+4, dy+4, fill=dot_col, outline="")
            dx += 14
        cv.create_rectangle(dx+5, dy-4, dx+28, dy+4, fill=color, outline="")

        # Click handler
        def _click(_, t=label.replace("\n", " ")):
            import tkinter.messagebox as _mb
            _mb.showinfo("Coming Soon",
                         f"'{t}'\n\nThis reconciliation module will be available soon.")
        cv.bind("<Button-1>", _click)

        return cv

    def _build_category_overview(self, frame, key, tools, accents, tv=None):
        COLS  = 3 if len(tools) <= 6 else 4
        if key == "gst":
            acc, icon, label = _C["gst_acc"],   "🏛",  "GST Tools"
        elif key == "pdf":
            acc, icon, label = _C["pdf_acc"],   "📄",  "PDF Tools"
        elif key == "bank":
            acc, icon, label = _C["bank_acc"],  "🏦",  "Bank Statement → Excel (Beta)"
        elif key == "mail":
            acc, icon, label = _C["mail_acc"],  "📨",  "Email Tools"
        elif key == "email":
            acc, icon, label = _C["email_acc"], "📧",  "Email Tools"
        elif key == "gmail":
            acc, icon, label = _C["gmail_acc"], "✉",  "Gmail Tools"
        elif key == "reco":
            acc, icon, label = _C["reco_acc"],  "🔄",  "GST Reconciliation"
        elif key == "tally":
            acc, icon, label = _C["tally_acc"], "🧾",  "Tally Automation Tools"
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
        visible_tools = [t for t in tools if not t.get("hide_card", False) and not t.get("hide_tab", False)]
        ctk.CTkLabel(badge, text=f"  {len(visible_tools)} Tools  ",
                     font=("Segoe UI", 10, "bold"),
                     text_color="#ffffff").pack(padx=4, pady=4)

        if key in ["gst", "it"]:
            acc_color = _C["gst_acc"] if key == "gst" else _C["it_acc"]
            ctk.CTkButton(tr, text="Manage ID/Pass",
                          command=lambda: _show_cred(),
                          fg_color=acc_color, hover_color=("#3730a3", "#6366f1"),
                          height=30, font=("Segoe UI", 11, "bold"),
                          corner_radius=8).pack(side="right", padx=(12, 0))

        hero_help = "Click any tab to load a tool — each loads once and stays active."
        if key == "mail":
            hero_help = "Choose Outlook or Gmail, then switch between 3 built-in templates and 1 custom template inside that suite."

        ctk.CTkLabel(hb,
                     text=hero_help,
                     font=("Segoe UI", 13),
                     text_color=("#1e293b", "#e2e8f0")).pack(anchor="w", pady=(8, 0))

        # Section label
        section_label = "Available Suites" if key == "mail" else "Available Tools"
        ctk.CTkLabel(scroll, text=section_label,
                     font=("Segoe UI", 14, "bold"),
                     text_color=("#0f172a", "#f1f5f9")).grid(
            row=1, column=0, columnspan=COLS,
            padx=18, pady=(6, 8), sticky="w")

        # Tool cards
        for idx, tool in enumerate(visible_tools):
            is_locked = (not self._is_tool_allowed(tool.get("key")))
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
            if not is_locked:
                action_cat = tool.get("action_cat")
                def _make_attach(tab_name=tool["tab"], _card=card, _action_cat=action_cat):
                    def _click(_=None):
                        if _action_cat:
                            self._show_category(_action_cat)
                        elif tv is not None:
                            tv.set(tab_name)
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

        if key in ["gst", "it"]:
            is_it = (key == "it")
            table_name = "it_profiles" if is_it else "gst_profiles"
            cat_label = "Income Tax" if is_it else "GST"
            cat_acc = _C["it_acc"] if is_it else _C["gst_acc"]
            
            cred_view = ctk.CTkFrame(frame, fg_color=("white", "#0f172a"), corner_radius=0)

            def _show_cred():
                scroll.pack_forget()
                cred_view.pack(fill="both", expand=True)
                _ov_refresh()

            def _show_cards():
                cred_view.pack_forget()
                scroll.pack(fill="both", expand=True)

            import sqlite3 as _sq_ov, os as _os_ov

            def _get_ov_db():
                p = _os_ov.path.join(_os_ov.environ.get("APPDATA", _os_ov.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
                _os_ov.makedirs(_os_ov.path.dirname(p), exist_ok=True)
                conn = _sq_ov.connect(p)
                conn.execute(f"CREATE TABLE IF NOT EXISTS {table_name} (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT)")
                try: conn.execute(f"ALTER TABLE {table_name} ADD COLUMN client_name TEXT")
                except: pass
                if table_name == "it_profiles":
                    try: conn.execute(f"ALTER TABLE {table_name} ADD COLUMN dob TEXT")
                    except: pass
                if table_name == "gst_profiles":
                    try: conn.execute(f"ALTER TABLE {table_name} ADD COLUMN filing_frequency TEXT")
                    except: pass
                    try: conn.execute(f"ALTER TABLE {table_name} ADD COLUMN gstin TEXT")
                    except: pass
                conn.commit()
                return conn

            cv_header = ctk.CTkFrame(cred_view, fg_color=_C["surface"], corner_radius=0)
            cv_header.pack(fill="x")
            ctk.CTkFrame(cv_header, height=4, fg_color=cat_acc, corner_radius=0).pack(fill="x")
            cv_hb = ctk.CTkFrame(cv_header, fg_color="transparent")
            cv_hb.pack(fill="x", padx=20, pady=10)
            ctk.CTkButton(cv_hb, text="<- Back to Tools", command=_show_cards,
                          fg_color="transparent", hover_color=("#e2e8f0", "#334155"),
                          text_color=cat_acc, font=("Segoe UI", 13, "bold"),
                          width=150, height=32).pack(side="left")
            ctk.CTkLabel(cv_hb, text=f"{cat_label} Profile Manager",
                         font=("Segoe UI", 16, "bold"),
                         text_color=cat_acc).pack(side="left", padx=(20, 0))

            cv_tabs = ctk.CTkTabview(cred_view, fg_color=_C["surface2"],
                                     segmented_button_fg_color=_C["surface"],
                                     segmented_button_selected_color=cat_acc,
                                     segmented_button_selected_hover_color=("#3730a3", "#6366f1"),
                                     segmented_button_unselected_color=_C["surface"],
                                     segmented_button_unselected_hover_color=("#e2e8f0", "#334155"))
            cv_tabs.pack(fill="both", expand=True, padx=20, pady=(8, 16))
            cv_tabs.add(f"  {cat_label} ID  ")

            # ══════════════════ PROFILE ID TAB ══════════════════
            gst_tab = cv_tabs.tab(f"  {cat_label} ID  ")
            gst_tab.grid_columnconfigure(0, weight=1)
            gst_tab.grid_columnconfigure(1, weight=1)
            gst_tab.grid_rowconfigure(0, weight=1)

            cv_left = ctk.CTkFrame(gst_tab, fg_color=_C["surface"], corner_radius=12,
                                   border_width=1, border_color=_C["border"])
            cv_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=4)
            cv_left.grid_columnconfigure(0, weight=1)
            cv_left.grid_rowconfigure(2, weight=1)
            ctk.CTkFrame(cv_left, height=4, fg_color=cat_acc, corner_radius=0).grid(row=0, column=0, sticky="ew")
            ctk.CTkLabel(cv_left, text="Saved Profiles",
                         font=("Segoe UI", 14, "bold"),
                         text_color=cat_acc).grid(row=1, column=0, sticky="w", padx=16, pady=(12, 6))
            list_box = ctk.CTkScrollableFrame(cv_left, fg_color=("#f8fafc", "#1a2535"), corner_radius=8)
            list_box.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
            list_box.grid_columnconfigure(0, weight=1)

            def _ov_refresh():
                for w in list_box.winfo_children():
                    w.destroy()
                try:
                    conn = _get_ov_db()
                    cur = conn.cursor()
                    cur.execute(f"SELECT * FROM {table_name} ORDER BY username")
                    cols = [d[0] for d in cur.description]
                    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
                    conn.close()
                except Exception:
                    rows = []
                if not rows:
                    ctk.CTkLabel(list_box, text="No profiles saved yet",
                                 font=("Segoe UI", 12),
                                 text_color=("#94a3b8", "#475569")).pack(pady=30)
                    return
                for rdata in rows:
                    rid = rdata.get("id")
                    uname = rdata.get("username", "")
                    cname = rdata.get("client_name") or ""
                    pwd = rdata.get("password", "")
                    dob = rdata.get("dob", "")
                    gstin_val = rdata.get("gstin", "")
                    row_f = ctk.CTkFrame(list_box, fg_color=("#ffffff", "#273549"),
                                          corner_radius=8, border_width=1,
                                          border_color=("#e2e8f0", "#334155"))
                    row_f.pack(fill="x", padx=6, pady=4)
                    row_f.grid_columnconfigure(0, weight=1)
                    freq = rdata.get("filing_frequency") or "Monthly"
                    
                    if not is_it:
                        disp_name = cname if cname else "Unnamed Client"
                        disp_gstin = gstin_val if gstin_val else uname
                        disp_text = f"  {disp_name} ({disp_gstin}) [{freq}]"
                    else:
                        disp_text = f"  {cname} ({uname})" if cname else f"  {uname}"
                        
                    ctk.CTkLabel(row_f, text=disp_text,
                                 font=("Segoe UI", 13, "bold"),
                                 anchor="w").grid(row=0, column=0, sticky="w", padx=12, pady=10)

                    def _edit(u=uname, p=pwd, c=cname, d=dob, f=freq, g=gstin_val):
                        ent_u.delete(0, "end"); ent_u.insert(0, u)
                        ent_p.delete(0, "end"); ent_p.insert(0, p)
                        ent_client.delete(0, "end"); ent_client.insert(0, c)
                        if ent_dob: ent_dob.delete(0, "end"); ent_dob.insert(0, d)
                        if ent_gstin: ent_gstin.delete(0, "end"); ent_gstin.insert(0, g)
                        if cb_freq: cb_freq.set(f)

                    ctk.CTkButton(row_f, text="Edit", width=60, height=30, fg_color="#2563EB", hover_color="#1D4ED8",
                                  font=("Segoe UI", 11, "bold"), command=_edit).grid(row=0, column=2, padx=(0, 5))
                    def _del(r=rid, u=uname):
                        from tkinter import messagebox as _mb2
                        if _mb2.askyesno("Delete", f"Delete profile for '{u}'?"):
                            try:
                                c = _get_ov_db(); c.execute(f"DELETE FROM {table_name} WHERE id=?", (r,)); c.commit(); c.close()
                            except Exception: pass
                            _ov_refresh()
                    ctk.CTkButton(row_f, text="Delete", width=70, height=30,
                                  fg_color="#7C3AED", hover_color="#6D28D9",
                                  font=("Segoe UI", 11, "bold"), command=_del
                                  ).grid(row=0, column=3, padx=(0, 10))

            cv_right = ctk.CTkFrame(gst_tab, fg_color=_C["surface"], corner_radius=12,
                                    border_width=1, border_color=_C["border"])
            cv_right.grid(row=0, column=1, sticky="nsew", pady=4)
            cv_right.grid_columnconfigure(0, weight=1)
            ctk.CTkFrame(cv_right, height=4, fg_color=cat_acc, corner_radius=0).pack(fill="x")
            rf = ctk.CTkFrame(cv_right, fg_color="transparent")
            rf.pack(fill="both", expand=True, padx=20, pady=(6, 12))
            ctk.CTkLabel(rf, text="Add New Profile",
                         font=("Segoe UI", 14, "bold"),
                         text_color=cat_acc).pack(anchor="w", pady=(0, 6))
            
            row1 = ctk.CTkFrame(rf, fg_color="transparent")
            row1.pack(fill="x", pady=(0, 6))
            row1.grid_columnconfigure((0, 1), weight=1)
            
            f_client = ctk.CTkFrame(row1, fg_color="transparent")
            f_client.grid(row=0, column=0, sticky="ew", padx=(0, 5))
            ctk.CTkLabel(f_client, text="Client Name", font=("Segoe UI", 12)).pack(anchor="w")
            ent_client = ctk.CTkEntry(f_client, placeholder_text="Enter Client Name", height=36)
            ent_client.pack(fill="x", pady=(2, 0))

            ent_dob = None
            cb_freq = None
            ent_gstin = None
            f_row1_col2 = ctk.CTkFrame(row1, fg_color="transparent")
            f_row1_col2.grid(row=0, column=1, sticky="ew", padx=(5, 0))
            if is_it:
                ctk.CTkLabel(f_row1_col2, text="Date of Birth", font=("Segoe UI", 12)).pack(anchor="w")
                ent_dob = ctk.CTkEntry(f_row1_col2, placeholder_text="DD/MM/YYYY", height=36)
                ent_dob.pack(fill="x", pady=(2, 0))
            else:
                ctk.CTkLabel(f_row1_col2, text="Filing Frequency", font=("Segoe UI", 12)).pack(anchor="w")
                cb_freq = ctk.CTkComboBox(f_row1_col2, values=["Monthly", "Quarterly"], height=36, state="readonly")
                cb_freq.pack(fill="x", pady=(2, 0))
                cb_freq.set("Monthly")
                def _on_freq_key(event):
                    if hasattr(event, "char") and event.char:
                        c = event.char.lower()
                        if c == 'm':
                            cb_freq.set("Monthly")
                        elif c == 'q':
                            cb_freq.set("Quarterly")
                cb_freq.bind("<Key>", _on_freq_key)

            row2 = ctk.CTkFrame(rf, fg_color="transparent")
            row2.pack(fill="x", pady=(0, 6))
            row2.grid_columnconfigure((0, 1), weight=1)

            user_field_label = "PAN / User ID" if is_it else "GST Username"
            user_placeholder = "Enter PAN/User ID" if is_it else "Enter GST username"
            
            f_u = ctk.CTkFrame(row2, fg_color="transparent")
            f_u.grid(row=0, column=0, sticky="ew", padx=(0, 5))
            ctk.CTkLabel(f_u, text=user_field_label, font=("Segoe UI", 12)).pack(anchor="w")
            ent_u = ctk.CTkEntry(f_u, placeholder_text=user_placeholder, height=36)
            ent_u.pack(fill="x", pady=(2, 0))

            f_p = ctk.CTkFrame(row2, fg_color="transparent")
            f_p.grid(row=0, column=1, sticky="ew", padx=(5, 0))
            ctk.CTkLabel(f_p, text="Password", font=("Segoe UI", 12)).pack(anchor="w")
            pr = ctk.CTkFrame(f_p, fg_color="transparent")
            pr.pack(fill="x", pady=(2, 0))
            ent_p = ctk.CTkEntry(pr, placeholder_text="Enter password", show="*", height=36)
            ent_p.pack(side="left", expand=True, fill="x")
            def _tog(e=ent_p):
                e.configure(show="" if e.cget("show") == "*" else "*")
            ctk.CTkButton(pr, text="Show", width=40, height=36,
                          fg_color="transparent", text_color=("#475569", "#94a3b8"),
                          hover_color=("#e2e8f0", "#334155"), command=_tog).pack(side="right", padx=(2, 0))

            if is_it:
                ctk.CTkLabel(rf, text="* DOB is only required for AIS and TIS tools.", font=("Segoe UI", 10), text_color=("#64748b", "#94a3b8")).pack(anchor="w", pady=(0, 5))
            else:
                row3 = ctk.CTkFrame(rf, fg_color="transparent")
                row3.pack(fill="x", pady=(0, 6))
                row3.grid_columnconfigure((0, 1), weight=1)
                
                f_gstin = ctk.CTkFrame(row3, fg_color="transparent")
                f_gstin.grid(row=0, column=0, sticky="ew", padx=(0, 5))
                ctk.CTkLabel(f_gstin, text="GSTIN (Optional)", font=("Segoe UI", 12)).pack(anchor="w")
                ent_gstin = ctk.CTkEntry(f_gstin, placeholder_text="Enter GSTIN", height=36)
                ent_gstin.pack(fill="x", pady=(2, 0))

            def _ov_save():
                from tkinter import messagebox as _mb3
                c = ent_client.get().strip()
                u = ent_u.get().strip()
                p = ent_p.get().strip()
                d = ent_dob.get().strip() if ent_dob else ""
                f = cb_freq.get() if cb_freq else "Monthly"
                g = ent_gstin.get().strip().upper() if ent_gstin else ""
                
                if not c or not u or not p:
                    _mb3.showerror("Missing", "Client Name, Username, and Password are mandatory.")
                    return
                
                try:
                    conn = _get_ov_db()
                    existing = conn.execute(f"SELECT id FROM {table_name} WHERE username=?", (u,)).fetchone()
                    if existing:
                        if is_it:
                            conn.execute(f"UPDATE {table_name} SET password=?, client_name=?, dob=? WHERE username=?", (p, c, d, u))
                        else:
                            conn.execute(f"UPDATE {table_name} SET password=?, client_name=?, filing_frequency=?, gstin=? WHERE username=?", (p, c, f, g, u))
                    else:
                        if is_it:
                            conn.execute(f"INSERT INTO {table_name} (username, password, client_name, dob) VALUES (?,?,?,?)", (u, p, c, d))
                        else:
                            conn.execute(f"INSERT INTO {table_name} (username, password, client_name, filing_frequency, gstin) VALUES (?,?,?,?,?)", (u, p, c, f, g))
                    conn.commit()
                    conn.close()
                except Exception as e:
                    _mb3.showerror("Error", str(e))
                    return
                ent_client.delete(0, "end")
                ent_u.delete(0, "end")
                ent_p.delete(0, "end")
                if ent_dob: ent_dob.delete(0, "end")
                if ent_gstin: ent_gstin.delete(0, "end")
                if cb_freq: cb_freq.set("Monthly")
                _ov_refresh()
            ctk.CTkButton(rf, text="✅  Save Profile", command=_ov_save,
                          fg_color="#059669", hover_color="#047857",
                          height=40, font=("Segoe UI", 14, "bold")).pack(fill="x", pady=(10, 0))

            # -- Import / Export Feature --
            import_frame = ctk.CTkFrame(rf, fg_color="transparent")
            import_frame.pack(fill="x", pady=(10, 0))
            import_frame.grid_columnconfigure((0,1), weight=1)

            def _dl_sample():
                from tkinter import filedialog as fd
                import pandas as pd
                from tkinter import messagebox as mb
                path = fd.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Sample_{cat_label}_Profiles.xlsx", filetypes=[("Excel", "*.xlsx")])
                if not path: return
                if is_it:
                    df = pd.DataFrame([{"Client Name": "John Doe", "Username (PAN)": "ABCDE1234F", "Password": "Pass123", "Date of Birth (DD/MM/YYYY)": "01/01/1990"}])
                else:
                    df = pd.DataFrame([{"Client Name": "Studycafe", "Filing Frequency": "Monthly", "Username": "", "Password": "Pass123", "GSTIN": "07AAAAA0000A1Z5"}])
                try:
                    df.to_excel(path, index=False)
                    mb.showinfo("Success", f"Sample downloaded successfully to:\n{path}")
                except Exception as e:
                    mb.showerror("Error", str(e))

            def _import_excel():
                from tkinter import filedialog as fd
                import pandas as pd
                from tkinter import messagebox as mb
                path = fd.askopenfilename(title="Import Excel File", filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls")])
                if not path: return
                try:
                    df = pd.read_excel(path)
                    conn = _get_ov_db()
                    count = 0
                    for _, row in df.iterrows():
                        row = row.fillna("")
                        c = str(row.get("Client Name", "")).strip()
                        
                        if is_it:
                            u = str(row.iloc[1] if "Username" not in str(df.columns[1]) else row.get(df.columns[1], "")).strip() 
                            if not u: u = str(row.get("Username (PAN)", "")).strip()
                            p = str(row.get("Password", "")).strip()
                            if not u or not p: continue
                            
                            d = str(row.get("Date of Birth (DD/MM/YYYY)", "")).strip()
                            existing = conn.execute(f"SELECT id FROM {table_name} WHERE username=?", (u,)).fetchone()
                            if existing:
                                conn.execute(f"UPDATE {table_name} SET password=?, client_name=?, dob=? WHERE username=?", (p, c, d, u))
                            else:
                                conn.execute(f"INSERT INTO {table_name} (username, password, client_name, dob) VALUES (?,?,?,?)", (u, p, c, d))
                        else:
                            g = str(row.get("GSTIN", "")).strip().upper()
                            u = str(row.get("Username", str(row.get("GST Username", "")))).strip()
                            p = str(row.get("Password", "")).strip()
                            f = str(row.get("Filing Frequency", "Monthly")).strip()
                            if not f: f = "Monthly"
                            
                            if not u:
                                u = g
                                
                            if not g and not u: continue
                            
                            existing = conn.execute(f"SELECT id FROM {table_name} WHERE username=?", (u,)).fetchone()
                            if existing:
                                conn.execute(f"UPDATE {table_name} SET password=?, client_name=?, filing_frequency=?, gstin=? WHERE username=?", (p, c, f, g, u))
                            else:
                                conn.execute(f"INSERT INTO {table_name} (username, password, client_name, filing_frequency, gstin) VALUES (?,?,?,?,?)", (u, p, c, f, g))
                        count += 1
                    conn.commit()
                    conn.close()
                    _ov_refresh()
                    mb.showinfo("Success", f"Imported {count} profiles successfully!")
                except Exception as e:
                    mb.showerror("Error", f"Failed to import:\n{e}")

            btn_dl = ctk.CTkButton(import_frame, text="Download Sample", command=_dl_sample, 
                                   fg_color="#334155", hover_color="#475569", height=38, font=("Segoe UI", 12, "bold"))
            btn_dl.grid(row=0, column=0, sticky="ew", padx=(0,5))
            btn_imp = ctk.CTkButton(import_frame, text="Import Excel", command=_import_excel, 
                                    fg_color="#2563EB", hover_color="#1D4ED8", height=38, font=("Segoe UI", 12, "bold"))
            btn_imp.grid(row=0, column=1, sticky="ew", padx=(5,0))

    # ══════════════════════════════════════════════════════════════════════════
    def _open_gst_profiles_manager(self):
        import sqlite3 as _sq, os as _os

        def _get_db():
            p = _os.path.join(_os.environ.get("APPDATA", _os.path.expanduser("~")), "GSTSuite", "suite_profiles.db")
            if not _os.path.exists(_os.path.dirname(p)):
                _os.makedirs(_os.path.dirname(p), exist_ok=True)
            conn = _sq.connect(p)
            conn.execute("CREATE TABLE IF NOT EXISTS gst_profiles (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT)")
            try: conn.execute("ALTER TABLE gst_profiles ADD COLUMN client_name TEXT")
            except: pass
            try: conn.execute("ALTER TABLE gst_profiles ADD COLUMN filing_frequency TEXT")
            except: pass
            conn.commit()
            return conn

        win = ctk.CTkToplevel(self)
        win.title("Manage GST ID / Password")
        win.geometry("480x620")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()
        win.attributes("-topmost", True)

        ctk.CTkFrame(win, height=5, fg_color=_C["gst_acc"]).pack(fill="x")

        ctk.CTkLabel(win, text="👤  GST Login Profiles",
                     font=("Segoe UI", 15, "bold"),
                     text_color=_C["gst_acc"]).pack(anchor="w", padx=20, pady=(14, 4))
        ctk.CTkLabel(win, text="Saved credentials are shared across all GST tools.",
                     font=("Segoe UI", 11),
                     text_color=("#64748b", "#94a3b8")).pack(anchor="w", padx=20, pady=(0, 10))

        list_frame = ctk.CTkScrollableFrame(win, height=220, fg_color=("#f8fafc", "#1e293b"),
                                            corner_radius=8)
        list_frame.pack(fill="x", padx=16, pady=(0, 4))   
        list_frame.grid_columnconfigure(0, weight=1)

        empty_lbl = ctk.CTkLabel(list_frame, text="No profiles saved yet.",
                                  font=("Segoe UI", 12),
                                  text_color=("#94a3b8", "#475569"))

        def _refresh_list():
            for w in list_frame.winfo_children(): w.destroy()
            try:
                conn = _get_db()
                cur = conn.cursor()
                cur.execute("SELECT * FROM gst_profiles ORDER BY username")
                cols = [d[0] for d in cur.description]
                rows = [dict(zip(cols, r)) for r in cur.fetchall()]
                conn.close()
            except Exception:
                rows = []
            if not rows:
                ctk.CTkLabel(list_frame, text="No profiles saved yet.", font=("Segoe UI", 12), text_color=("#94a3b8", "#475569")).pack(pady=20)
                return
            for rdata in rows:
                rid = rdata.get("id")
                uname = rdata.get("username", "")
                pwd = rdata.get("password", "")
                cname = rdata.get("client_name") or ""
                freq = rdata.get("filing_frequency") or "Monthly"
                row = ctk.CTkFrame(list_frame, fg_color=("#ffffff", "#273549"), corner_radius=8, border_width=1, border_color=("#e2e8f0", "#334155"))
                row.pack(fill="x", padx=4, pady=3)
                row.grid_columnconfigure(0, weight=1)
                disp_text = f"  {cname} ({uname}) [{freq}]" if cname else f"  {uname} [{freq}]"
                ctk.CTkLabel(row, text=disp_text, font=("Segoe UI", 12, "bold"), anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=6)
                ctk.CTkLabel(row, text="••••••••", font=("Segoe UI", 11), text_color=("#94a3b8", "#64748b")).grid(row=0, column=1, padx=8)
                def _edit(u=uname, p=pwd, c=cname, f=freq):
                    ent_user.delete(0, "end"); ent_user.insert(0, u)
                    ent_pass.delete(0, "end"); ent_pass.insert(0, p)
                    ent_client.delete(0, "end"); ent_client.insert(0, c)
                    cb_freq.set(f)
                ctk.CTkButton(row, text="Edit", width=50, height=28, fg_color="#2563EB", hover_color="#1D4ED8", font=("Segoe UI", 11, "bold"), command=_edit).grid(row=0, column=2, padx=(0, 5))
                def _del(r=rid, u=uname):
                    from tkinter import messagebox as _mb
                    if _mb.askyesno("Delete", f"Delete profile for '{u}'?", parent=win):
                        try:
                            c = _get_db(); c.execute("DELETE FROM gst_profiles WHERE id=?", (r,)); c.commit(); c.close()
                        except Exception: pass
                        _refresh_list()
                ctk.CTkButton(row, text="🗑", width=34, height=28, fg_color="#7C3AED", hover_color="#6D28D9", font=("Segoe UI", 12), command=_del).grid(row=0, column=3, padx=(0, 8))

        _refresh_list()

        ctk.CTkFrame(win, height=1, fg_color=("#e2e8f0", "#334155")).pack(fill="x", padx=16, pady=(8, 0))
        ctk.CTkLabel(win, text="Add New Profile",
                     font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=20, pady=(10, 4))

        add_frame = ctk.CTkFrame(win, fg_color="transparent")
        add_frame.pack(fill="x", padx=16, pady=(0, 4))

        ctk.CTkLabel(add_frame, text="Client Name (Optional)", font=("Segoe UI", 11)).pack(anchor="w")
        ent_client = ctk.CTkEntry(add_frame, placeholder_text="Enter Client Name", height=34)
        ent_client.pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(add_frame, text="Filing Frequency", font=("Segoe UI", 11)).pack(anchor="w")
        cb_freq = ctk.CTkComboBox(add_frame, values=["Monthly", "Quarterly"], height=34)
        cb_freq.set("Monthly")
        cb_freq.pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(add_frame, text="Username / GST ID", font=("Segoe UI", 11)).pack(anchor="w")
        ent_user = ctk.CTkEntry(add_frame, placeholder_text="Enter GST username", height=34)
        ent_user.pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(add_frame, text="Password", font=("Segoe UI", 11)).pack(anchor="w")
        pass_row = ctk.CTkFrame(add_frame, fg_color="transparent")
        pass_row.pack(fill="x", pady=(2, 0))
        ent_pass = ctk.CTkEntry(pass_row, placeholder_text="Enter password", show="*", height=34)
        ent_pass.pack(side="left", expand=True, fill="x")
        def _toggle():
            ent_pass.configure(show="" if ent_pass.cget("show") == "*" else "*")
        ctk.CTkButton(pass_row, text="👁", width=36, height=34,
                      fg_color="transparent", text_color=("#475569", "#94a3b8"),
                      hover_color=("#e2e8f0", "#334155"), command=_toggle).pack(side="right", padx=(6, 0))

        def _save():
            from tkinter import messagebox as _mb
            c = ent_client.get().strip()
            u = ent_user.get().strip()
            p = ent_pass.get().strip()
            f = cb_freq.get()
            if not u or not p:
                _mb.showerror("Missing", "Please enter both username and password.", parent=win)
                return
            try:
                conn = _get_db()
                existing = conn.execute("SELECT id FROM gst_profiles WHERE username=?", (u,)).fetchone()
                if existing:
                    conn.execute("UPDATE gst_profiles SET password=?, client_name=?, filing_frequency=? WHERE username=?", (p, c, f, u))
                else:
                    conn.execute("INSERT INTO gst_profiles (username, password, client_name, filing_frequency) VALUES (?, ?, ?, ?)", (u, p, c, f))
                conn.commit()
                conn.close()
            except Exception as e:
                _mb.showerror("Error", str(e), parent=win)
                return
            ent_client.delete(0, "end")
            ent_user.delete(0, "end")
            ent_pass.delete(0, "end")
            cb_freq.set("Monthly")
            _refresh_list()

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(10, 16))
        ctk.CTkButton(btn_row, text="✅ Save Profile", command=_save,
                      fg_color="#059669", hover_color="#047857",
                      height=34, font=("Segoe UI", 12, "bold")).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ctk.CTkButton(btn_row, text="Close", command=win.destroy,
                      fg_color="#475569", hover_color="#334155",
                      height=34, font=("Segoe UI", 12, "bold")).pack(side="left", expand=True, fill="x")


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
        elif cat_key == "mail":
            tools, accents = MAIL_GROUP_TOOLS, _MAIL_GROUP_ACCENTS
        elif cat_key == "email":
            tools, accents = EMAIL_TOOLS, _EMAIL_ACCENTS
        elif cat_key == "gmail":
            tools, accents = GMAIL_TOOLS, _GMAIL_ACCENTS
        elif cat_key == "reco":
            tools, accents = RECO_TOOLS,  _RECO_ACCENTS
        elif cat_key == "tally":
            tools, accents = TALLY_TOOLS, _TALLY_ACCENTS
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

        # ── Builtin UIs (no external module loading) ─────────────────────────
        if tool.get("builtin_ui"):
            tab_frame.grid_rowconfigure(0, weight=1)
            tab_frame.grid_columnconfigure(0, weight=1)
            self._build_builtin_ui(tool["builtin_ui"], tab_frame, accent)
            self._loaded[name] = True
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
        self.title("AutomationCafe Suite — Login")
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

        ctk.CTkLabel(wrap, text="AutomationCafe Suite",
                     font=("Segoe UI", 22, "bold"),
                     text_color=("#6366f1", "#818cf8")).pack()

        ctk.CTkLabel(wrap, text="Sign in to continue",
                     font=("Segoe UI", 12),
                     text_color=("#64748b", "#94a3b8")).pack(pady=(4, 24))

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

        pass_outer = ctk.CTkFrame(wrap, fg_color="transparent")
        pass_outer.pack(fill="x", pady=(4, 6))

        self._pass_entry = ctk.CTkEntry(pass_outer, placeholder_text="Enter your password",
                                        show="●", height=40, corner_radius=8,
                                        font=("Segoe UI", 13))
        self._pass_entry.pack(fill="x")
        self._pass_entry.bind("<Return>", lambda e: self._do_login())

        def _toggle_pass():
            if self._pass_entry.cget("show") == "":
                self._pass_entry.configure(show="●")
                eye_btn.configure(text="Show")
            else:
                self._pass_entry.configure(show="")
                eye_btn.configure(text="Hide")

        eye_btn = ctk.CTkButton(pass_outer, text="Show", width=50, height=26,
                                fg_color=("#e2e8f0", "#3a3c3e"),
                                hover_color=("#cbd5e1", "#4a4c4e"),
                                bg_color=("#F9F9FA", "#343638"),
                                text_color=("#475569", "#94a3b8"),
                                border_width=0, corner_radius=13,
                                font=("Segoe UI", 11, "bold"),
                                command=_toggle_pass)
        eye_btn.place(relx=1.0, rely=0.5, anchor="e", x=-8)
        eye_btn.configure(cursor="hand2")

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
        ctk.CTkButton(wrap, text="Try 3-Day Free Trial",
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
                "Your 3-day trial has expired.\nContact our support: Info@studycafe.in | 96250 80264",
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
            # Write a bat that waits for this process to fully exit (so PyInstaller
            # finishes cleaning up _MEI temp dir) before launching the new instance.
            import tempfile
            bat_path = os.path.join(tempfile.gettempdir(), "_ac_restart.bat")
            exe_path = sys.executable
            pid = os.getpid()
            with open(bat_path, "w") as f:
                f.write("@echo off\n")
                f.write(":wait\n")
                f.write(f'tasklist /fi "PID eq {pid}" 2>nul | find "{pid}" >nul\n')
                f.write("if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)\n")
                f.write(f'start "" "{exe_path}"\n')
                f.write('del "%~f0"\n')
            subprocess.Popen(
                ["cmd", "/c", bat_path],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        else:
            subprocess.Popen([sys.executable, os.path.abspath(__file__)])
        sys.exit(0)
    except Exception as e:
        _suite_debug_log(f"relaunch failed: {e}")


def run_app_lifecycle():
    """Single top-level event-loop orchestratfion to avoid nested mainloops."""
    while True:
        # ── Try silent auto-login with saved credentials ───────────────────
        user_info = None
        _saved_email, _saved_password = _load_auth()
        if _saved_email and _saved_password:
            _update_boot_splash("Signing you in...")
            try:
                resp = _call_api("/check_session", {
                    "email":       _saved_email,
                    "password":    _saved_password,
                    "hardware_id": _get_hardware_id(),
                })
                if resp.get("status") == "SESSION_VALID":
                    resp["status"] = "SUCCESS"
                    user_info = resp
            except Exception:
                pass  # network error → fall through to login window

        if user_info:
            _update_boot_splash("Opening automation suite...")
            _close_boot_splash()
        else:
            # ── Show login window ──────────────────────────────────────────
            _update_boot_splash("Opening login screen...")
            _close_boot_splash()
            try:
                _tk._default_root = None
            except Exception:
                pass
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

        _show_boot_splash()
        _update_boot_splash("Opening automation suite...")
        _close_boot_splash()
        try:
            _tk._default_root = None
        except Exception:
            pass
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
            # In-process soft restart: loop back to login without re-launching.
            # Re-launching a PyInstaller onefile exe causes a _MEI temp-dir race
            # condition where cleanup from the old process deletes files the new
            # process needs before it has finished extracting them.
            continue

        # Normal exit (window closed without logout)
        break


if __name__ == "__main__":
    try:
        if _SPLASH is not None:
            _SPLASH.destroy()
    except Exception:
        pass
    try:
        run_app_lifecycle()
    finally:
        _close_boot_splash()




