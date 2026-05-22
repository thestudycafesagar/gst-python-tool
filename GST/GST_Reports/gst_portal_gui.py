import os
import io
import sys
import json
import base64
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from pathlib import Path
from datetime import datetime
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gstr2b_downloader import Gstr2BDownloader, save_result  as save_2b
from gstr3b_downloader import Gstr3BDownloader, save_gstr3b
from gstr2a_downloader import Gstr2ADownloader, save_gstr2a
from gstr1_downloader  import Gstr1Downloader,  save_gstr1

# =============================================================================
# Period constants
# =============================================================================

MONTHS = [
    ("January","1"),("February","2"),("March","3"),("April","4"),
    ("May","5"),("June","6"),("July","7"),("August","8"),
    ("September","9"),("October","10"),("November","11"),("December","12"),
]
QUARTERS = [
    ("Q1 — April to June",       "6"),
    ("Q2 — July to September",   "9"),
    ("Q3 — October to December", "12"),
    ("Q4 — January to March",    "3"),
]
MONTHLY_FY_ORDER   = ["4","5","6","7","8","9","10","11","12","1","2","3"]
QUARTERLY_FY_ORDER = ["6","9","12","3"]
MONTH_NAMES   = {m[1]: m[0] for m in MONTHS}
QUARTER_NAMES = {q[1]: q[0] for q in QUARTERS}
CURRENT_YEAR  = datetime.now().year
YEARS         = [str(y) for y in range(CURRENT_YEAR, CURRENT_YEAR - 6, -1)]

# =============================================================================
# Theme — Modern Dark
# =============================================================================

if CTK_AVAILABLE:
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

# Traditional colors for fallback or specific logic
ACCENT    = "#1f538d"  # Default Blue
GREEN     = "#2ea043"
ORANGE    = "#d83b01"
GREY      = "#605e5c"
DANGER    = "#cf222e"
PURPLE    = "#8250df"
TEAL      = "#0598bc"

# =============================================================================
# Shared session factory
# =============================================================================

def make_session() -> requests.Session:
    sess  = requests.Session()
    retry = Retry(total=5, backoff_factor=1,
                  status_forcelist=[500, 502, 503, 504],
                  allowed_methods=["GET", "POST"])
    sess.mount("https://", HTTPAdapter(max_retries=retry))
    sess.mount("http://",  HTTPAdapter(max_retries=retry))
    return sess

# =============================================================================
# Main Application
# =============================================================================

class GstPortalApp(ctk.CTk if CTK_AVAILABLE else tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("GST Portal Downloader")
        
        # Window Setup
        if CTK_AVAILABLE:
            self._set_appearance_mode("dark")
        
        self.minsize(1000, 680)
        try:
            self.state("zoomed")
        except tk.TclError:
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            self.geometry(f"{int(sw*.9)}x{int(sh*.9)}+{int(sw*.05)}+{int(sh*.05)}")

        # ── Shared state ──────────────────────────────────────────────────────
        self._session: Optional[requests.Session] = None
        self._login_dl: Optional[Gstr3BDownloader] = None
        self._logged_in   = False
        self._otp_pending = False
        self._yearly_stop = False
        self._output_dir  = str(Path(__file__).parent)
        self._profile     = {}
        self._app_config_path = str(Path(__file__).parent / "app_config.json")
        self._app_config  = self._load_app_config()

        self._loaded_credentials = None   # {"Username": ..., "Password": ..., "ClientName": ...}
        self._build_ui()
        self._setup_traces()
        self._set_state("idle")

    def _setup_traces(self):
        """Monitor tab scopes to show/hide consolidated option."""
        for var in [self.var_1_scope, self.var_2a_scope, 
                    self.var_2b_scope, self.var_3b_scope]:
            var.trace_add("write", lambda *a: self._update_excel_mode_visibility())
        self._update_excel_mode_visibility()

    def _update_excel_mode_visibility(self):
        """Shows consolidated option only if 'Full Year' is selected in any tab."""
        any_yearly = any(v.get() == "year" for v in [
            self.var_1_scope, self.var_2a_scope, 
            self.var_2b_scope, self.var_3b_scope
        ])
        if any_yearly:
            self.excel_mode_row.pack(anchor="w", pady=(5, 0), padx=10)
        else:
            self.excel_mode_row.pack_forget()
            self.var_excel_mode.set("individual")

    # =========================================================================
    # UI Components
    # =========================================================================

    def _build_ui(self):
        # Single content row — no header row
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)   # left sidebar
        self.grid_columnconfigure(1, weight=1)   # right panel

        # ── Left sidebar (login + captcha) ────────────────────────────────────
        self._left = ctk.CTkScrollableFrame(self, width=310, corner_radius=0,
                                             fg_color=("#f8fafc", "#111827"))
        self._left.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)

        self._build_login_section()
        self._build_format_section()
        self._build_captcha_section()
        self._build_otp_section()
        self._build_action_buttons()

        # ── Right panel (tabs + log) ──────────────────────────────────────────
        self._right = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self._right.grid(row=0, column=1, sticky="nsew", padx=8, pady=6)

        self._right.grid_columnconfigure(0, weight=1)
        self._right.grid_rowconfigure(0, weight=5)   # Tabview expands most
        self._right.grid_rowconfigure(1, weight=0)   # Status/Progress — fixed
        self._right.grid_rowconfigure(2, weight=1)   # Log — compact

        # Tabview (no fixed height — expands to fill space)
        self.tabview = ctk.CTkTabview(self._right,
                                       fg_color=("gray96", "#1e293b"),
                                       border_width=1,
                                       border_color=("gray85", "#334155"),
                                       segmented_button_fg_color=("gray90", "#111827"),
                                       segmented_button_selected_color=("#8b5cf6", "#7c3aed"),
                                       segmented_button_selected_hover_color=("#7c3aed", "#6d28d9"),
                                       segmented_button_unselected_color=("gray90", "#111827"),
                                       segmented_button_unselected_hover_color=("gray85", "#1e293b"))
        self.tabview.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        
        self._tab_1 = self.tabview.add("GSTR-1")
        self._tab_2a = self.tabview.add("GSTR-2A")
        self._tab_2b = self.tabview.add("GSTR-2B")
        self._tab_3b = self.tabview.add("GSTR-3B")

        self._fill_tab_1()
        self._fill_tab_2a()
        self._fill_tab_2b()
        self._fill_tab_3b()

        # Progress / Status Area
        status_frame = ctk.CTkFrame(self._right, fg_color="transparent")
        status_frame.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        
        self.lbl_status = ctk.CTkLabel(status_frame, text="Ready", font=ctk.CTkFont(weight="bold"))
        self.lbl_status.pack(side="left", padx=5)

        self.progress = ctk.CTkProgressBar(status_frame, orientation="horizontal", height=10)
        self.progress.pack(side="right", fill="x", expand=True, padx=10)
        self.progress.set(0)

        # Log Panel
        self._build_log_panel()

    def _build_login_section(self):
        # Hidden StringVars — populated when a profile is loaded; used by existing login logic
        self.var_username = tk.StringVar()
        self.var_password = tk.StringVar()

        box = ctk.CTkFrame(self._left, fg_color="transparent")
        box.pack(fill="x", pady=(10, 6), padx=16)

        ctk.CTkLabel(box, text="CREDENTIALS",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=("#334155", "#94a3b8")).pack(anchor="w", pady=(0, 6))

        # Profile display (read-only, shows loaded profile name / username)
        self._cred_display_var = tk.StringVar(value="No profile loaded")
        cred_disp = ctk.CTkEntry(box, textvariable=self._cred_display_var,
                                  state="readonly", height=32,
                                  fg_color=("gray93", "#1e293b"),
                                  border_color=("gray80", "#334155"))
        cred_disp.pack(fill="x", pady=(0, 6))

        # Load buttons row
        btn_row = ctk.CTkFrame(box, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 4))

        ctk.CTkButton(btn_row, text="📂  Load ID Pass",
                      height=32, font=ctk.CTkFont(size=12, weight="bold"),
                      fg_color=("#059669", "#047857"),
                      hover_color=("#047857", "#065f46"),
                      command=self._load_id_pass_dialog).pack(side="left", expand=True, fill="x", padx=(0, 4))

        ctk.CTkButton(btn_row, text="▶  Demo",
                      height=32, width=80,
                      font=ctk.CTkFont(size=11, weight="bold"),
                      fg_color=("#DC2626", "#B91C1C"),
                      hover_color=("#B91C1C", "#991B1B"),
                      command=self._open_demo_link).pack(side="right")

        # View loaded cred
        self._btn_view_cred = ctk.CTkButton(
            box, text="👁  View Loaded Profile",
            height=26, width=140,
            fg_color="transparent", border_width=1,
            border_color=("gray75", "#334155"),
            text_color=("#475569", "#94a3b8"),
            font=ctk.CTkFont(size=11),
            state="disabled",
            command=self._view_loaded_cred)
        self._btn_view_cred.pack(anchor="w", pady=(0, 4))

        # Save To
        ctk.CTkLabel(box, text="Save To", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(4, 0))
        out_row = ctk.CTkFrame(box, fg_color="transparent")
        out_row.pack(fill="x", pady=(2, 0))
        self.var_output = tk.StringVar(value=self._output_dir)
        ctk.CTkEntry(out_row, textvariable=self.var_output, height=32).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(out_row, text="...", width=35, height=32,
                      command=self._browse_output).pack(side="right", padx=(5, 0))

    # ── Credential helpers ────────────────────────────────────────────────────

    def _refresh_cred_display(self):
        """Update the display entry and view-button after loading credentials."""
        c = self._loaded_credentials
        if not c:
            self._cred_display_var.set("No profile loaded")
            self._btn_view_cred.configure(state="disabled")
            return
        name = c.get("ClientName") or c.get("Username", "")
        user = c.get("Username", "")
        disp = f"{name} ({user})" if name and name != user else user
        self._cred_display_var.set(disp)
        self._btn_view_cred.configure(state="normal")
        # Push into hidden vars so login logic works without changes
        self.var_username.set(c.get("Username", ""))
        self.var_password.set(c.get("Password", ""))

    def _view_loaded_cred(self):
        c = self._loaded_credentials
        if not c:
            return
        from tkinter import messagebox as _mb
        _mb.showinfo(
            "Loaded Profile",
            f"Username : {c.get('Username','')}\n"
            f"Client   : {c.get('ClientName','')}\n"
            f"Password : {'*' * len(c.get('Password',''))}",
            parent=self
        )

    def _load_id_pass_dialog(self):
        import sqlite3 as _sq
        db_path = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")),
            "GSTSuite", "suite_profiles.db"
        )
        if not os.path.exists(db_path):
            db_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "..", "..", "suite_profiles.db"
            )
        rows = []
        try:
            conn = _sq.connect(db_path)
            cur  = conn.cursor()
            cur.execute("SELECT * FROM gst_profiles ORDER BY username")
            cols = [d[0] for d in cur.description]
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]
            conn.close()
        except Exception:
            pass

        if not rows:
            from tkinter import messagebox as _mb
            _mb.showinfo("No Profiles",
                         "No saved profiles found.\nPlease add profiles via GST Suite settings.",
                         parent=self)
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Load ID Password")
        dlg.geometry("440x520")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        dlg.attributes("-topmost", True)

        ctk.CTkLabel(dlg, text="Select a Profile",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(16, 6))

        search_var = ctk.StringVar()
        ctk.CTkEntry(dlg, placeholder_text="🔍  Search by name or username…",
                     textvariable=search_var, height=34).pack(fill="x", padx=16, pady=(0, 6))

        scroll = ctk.CTkScrollableFrame(dlg, height=280)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 6))

        selected_var = ctk.StringVar()
        data_map  = {}
        row_widgets = {}

        for i, r in enumerate(rows):
            u  = r.get("username", "")
            p  = r.get("password", "")
            c  = r.get("client_name") or ""
            ff = r.get("filing_frequency") or "Monthly"
            disp = f"{c}  ({u})" if c else u
            uid  = f"p{i}"
            data_map[uid] = {"Username": u, "Password": p, "ClientName": c, "FilingFrequency": ff}
            rb = ctk.CTkRadioButton(scroll, text=disp, variable=selected_var, value=uid)
            rb.pack(anchor="w", padx=10, pady=4)
            row_widgets[uid] = (rb, disp)

        def _on_search(*_):
            q = search_var.get().strip().lower()
            for key, (rb, disp) in row_widgets.items():
                if not q or q in disp.lower():
                    rb.pack(anchor="w", padx=10, pady=4)
                else:
                    rb.pack_forget()

        search_var.trace_add("write", _on_search)

        foot = ctk.CTkFrame(dlg, fg_color="transparent")
        foot.pack(fill="x", padx=16, pady=(0, 16))

        def _do_load():
            uid = selected_var.get()
            if not uid or uid not in data_map:
                from tkinter import messagebox as _mb
                _mb.showwarning("No Selection", "Please select a profile.", parent=dlg)
                return
            self._loaded_credentials = data_map[uid]
            self._refresh_cred_display()
            dlg.destroy()

        ctk.CTkButton(foot, text="Cancel", width=110,
                      command=dlg.destroy).pack(side="right")
        ctk.CTkButton(foot, text="Load", width=120,
                      fg_color=("#059669", "#047857"),
                      hover_color=("#047857", "#065f46"),
                      command=_do_load).pack(side="right", padx=(0, 8))

    def _browse_excel_creds(self):
        from tkinter import filedialog as _fd
        path = _fd.askopenfilename(
            title="Select Excel with Username / Password columns",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path:
            return
        try:
            import pandas as pd
            df = pd.read_excel(path)
            clean = {c.lower().strip(): c for c in df.columns}
            u_col = next((clean[k] for k in clean if "user" in k or "name" in k), None)
            p_col = next((clean[k] for k in clean if "pass" in k or "pwd"  in k), None)
            if not u_col or not p_col:
                from tkinter import messagebox as _mb
                _mb.showerror("Column Error",
                              "Need columns: Username / Password\n"
                              f"Found: {list(df.columns)}",
                              parent=self)
                return
            row = df.iloc[0]
            self._loaded_credentials = {
                "Username":   str(row[u_col]).strip(),
                "Password":   str(row[p_col]).strip(),
                "ClientName": str(row.get("ClientName", row.get("Client Name", ""))).strip(),
            }
            self._refresh_cred_display()
        except Exception as ex:
            from tkinter import messagebox as _mb
            _mb.showerror("Excel Error", str(ex), parent=self)

    def _open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://youtu.be/zEggEXMjL-w")

    def _build_format_section(self):
        box = ctk.CTkFrame(self._left, fg_color=("gray93", "#2a2d2e"), corner_radius=10)
        box.pack(fill="x", pady=10, padx=20)

        ctk.CTkLabel(box, text="OUTPUT SETTINGS",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=("#334155", "#94a3b8")).pack(anchor="w", padx=15, pady=(10, 5))

        self.var_fmt_excel = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(box, text="Generate Excel (.xlsx)", variable=self.var_fmt_excel).pack(anchor="w", padx=15, pady=5)

        self.var_excel_mode = tk.StringVar(value="individual")
        self.excel_mode_row = ctk.CTkFrame(box, fg_color="transparent")
        ctk.CTkRadioButton(self.excel_mode_row, text="Individual", variable=self.var_excel_mode, value="individual").pack(side="left")
        ctk.CTkRadioButton(self.excel_mode_row, text="Consolidated", variable=self.var_excel_mode, value="consolidated").pack(side="left", padx=15)

    def _build_captcha_section(self):
        box = ctk.CTkFrame(self._left, fg_color="transparent")
        box.pack(fill="x", pady=(5, 10), padx=20)

        self.btn_captcha = ctk.CTkButton(box, text="GET CAPTCHA", command=self._on_get_captcha, height=35, font=ctk.CTkFont(weight="bold"))
        self.btn_captcha.pack(fill="x", pady=(0, 8))

        self.captcha_label = ctk.CTkLabel(box, text="[ Captcha Image ]", height=65, fg_color=("gray85", "#1a1a1a"), corner_radius=5)
        self.captcha_label.pack(fill="x", pady=(0, 8))

        row = ctk.CTkFrame(box, fg_color="transparent")
        row.pack(fill="x")
        ctk.CTkLabel(row, text="Captcha:").pack(side="left")
        self.var_captcha = tk.StringVar()
        self.entry_captcha = ctk.CTkEntry(row, textvariable=self.var_captcha, height=32, placeholder_text="CODE", font=ctk.CTkFont(size=14, weight="bold"), justify="center")
        self.entry_captcha.pack(side="right", fill="x", expand=True, padx=(10, 0))

    def _build_otp_section(self):
        self.otp_frame = ctk.CTkFrame(self._left, fg_color=("gray93", "#3e3b2e"), border_width=1, border_color="#d4ac0d")
        # Pack/forget dynamically
        
        ctk.CTkLabel(self.otp_frame, text="OTP Verification Required", text_color="#d4ac0d", font=ctk.CTkFont(weight="bold")).pack(pady=5)
        row = ctk.CTkFrame(self.otp_frame, fg_color="transparent")
        row.pack(pady=5, padx=10)
        self.var_otp = tk.StringVar()
        self.entry_otp = ctk.CTkEntry(row, textvariable=self.var_otp, width=120, height=35, justify="center", font=ctk.CTkFont(size=18, weight="bold"))
        self.entry_otp.pack(side="left", padx=5)

    def _build_action_buttons(self):
        self._action_frame = ctk.CTkFrame(self._left, fg_color="transparent")
        self._action_frame.pack(fill="x", pady=(5, 20), padx=20)
        
        self.btn_login = ctk.CTkButton(self._action_frame, text="LOGIN", command=self._on_login, height=40, fg_color=GREEN, hover_color="#268635", font=ctk.CTkFont(weight="bold"))
        self.btn_login.pack(fill="x", pady=(0, 8))
        
        self.btn_logout = ctk.CTkButton(self._action_frame, text="LOGOUT", command=self._on_logout, height=35, fg_color="#444", hover_color="#555")
        self.btn_logout.pack(fill="x")

    def _fill_tab_1(self):
        tab = self.tabview.tab("GSTR-1")
        tab.grid_columnconfigure((1,3), weight=1)

        ctk.CTkLabel(tab, text="Select Period", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", pady=(8,4), padx=10)

        ctk.CTkLabel(tab, text="Month").grid(row=1, column=0, sticky="e", padx=10)
        self.var_1_month = tk.StringVar(value="April")
        self.cb_1_month = ctk.CTkOptionMenu(tab, values=[m[0] for m in MONTHS], variable=self.var_1_month)
        self.cb_1_month.grid(row=1, column=1, sticky="ew", pady=3)

        ctk.CTkLabel(tab, text="Year").grid(row=1, column=2, sticky="e", padx=10)
        self.var_1_year = tk.StringVar(value=str(CURRENT_YEAR))
        ctk.CTkOptionMenu(tab, values=YEARS, variable=self.var_1_year).grid(row=1, column=3, sticky="ew", pady=3)

        ctk.CTkLabel(tab, text="GSTIN (Optional)").grid(row=2, column=0, sticky="e", padx=10)
        self.var_1_gstin = tk.StringVar()
        ctk.CTkEntry(tab, textvariable=self.var_1_gstin, placeholder_text="Lookup B2B Detail").grid(row=2, column=1, columnspan=3, sticky="ew", pady=3)

        sf = ctk.CTkFrame(tab, fg_color="transparent")
        sf.grid(row=3, column=0, columnspan=4, pady=6)
        self.var_1_scope = tk.StringVar(value="single")
        ctk.CTkRadioButton(sf, text="Single Month", variable=self.var_1_scope, value="single").pack(side="left", padx=10)
        ctk.CTkRadioButton(sf, text="Full Financial Year", variable=self.var_1_scope, value="year", text_color="#ffcc00").pack(side="left", padx=10)

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=4, column=0, columnspan=4, pady=6)

        self.btn_dl_1 = ctk.CTkButton(btn_row, text="DOWNLOAD GSTR-1", command=self._on_download_gstr1, width=200, height=38, font=ctk.CTkFont(weight="bold"))
        self.btn_dl_1.pack(side="left", padx=5)
        self.btn_stop_1 = ctk.CTkButton(btn_row, text="STOP", command=self._on_stop, fg_color=DANGER, hover_color="#a01a23", width=80, height=38)
        self.btn_stop_1.pack(side="left", padx=5)

        self.lbl_1_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_1_progress.grid(row=5, column=0, columnspan=2, sticky="w", padx=10)

        self.btn_refresh_names = ctk.CTkButton(tab, text="Refresh Party Names", command=self._on_refresh_names, width=150, height=28, fg_color=("gray75", "#333333"))
        self.btn_refresh_names.grid(row=5, column=2, columnspan=2, sticky="e", padx=10)

    def _fill_tab_2a(self):
        tab = self.tabview.tab("GSTR-2A")
        tab.grid_columnconfigure((1,3), weight=1)

        ctk.CTkLabel(tab, text="Month").grid(row=0, column=0, sticky="e", padx=10, pady=(4,2))
        self.var_2a_month = tk.StringVar(value="April")
        ctk.CTkOptionMenu(tab, values=[m[0] for m in MONTHS], variable=self.var_2a_month).grid(row=0, column=1, sticky="ew", pady=(4,2))

        ctk.CTkLabel(tab, text="Year").grid(row=0, column=2, sticky="e", padx=10)
        self.var_2a_year = tk.StringVar(value=str(CURRENT_YEAR))
        ctk.CTkOptionMenu(tab, values=YEARS, variable=self.var_2a_year).grid(row=0, column=3, sticky="ew", pady=(4,2))

        sf = ctk.CTkFrame(tab, fg_color="transparent")
        sf.grid(row=1, column=0, columnspan=4, pady=2)
        self.var_2a_scope = tk.StringVar(value="single")
        ctk.CTkRadioButton(sf, text="Single Month", variable=self.var_2a_scope, value="single").pack(side="left", padx=10)
        ctk.CTkRadioButton(sf, text="Full Financial Year", variable=self.var_2a_scope, value="year", text_color="#ffcc00").pack(side="left", padx=10)

        sec_box = ctk.CTkFrame(tab, fg_color=("gray93", "#1a1a1a"), corner_radius=8)
        sec_box.grid(row=2, column=0, columnspan=4, sticky="nsew", padx=10, pady=2)

        self._2a_section_vars = {}
        sections_2a = [
            ("B2B","b2b"),   ("B2BA","b2ba"),  ("CDN","cdn"),
            ("CDNA","cdna"), ("TDS","tds"),     ("TCS","tcs"),
            ("ISD","isd"),   ("ISDA","isda"),   ("TDSA","tdsa"),
            ("IMPG","impg"), ("IMPGSEZ","impgsez"),
        ]
        for i, (label, key) in enumerate(sections_2a):
            v = tk.BooleanVar(value=True)
            self._2a_section_vars[key] = v
            ctk.CTkCheckBox(sec_box, text=label, variable=v, font=ctk.CTkFont(size=11)).grid(row=i//4, column=i%4, sticky="w", padx=6, pady=1)

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=3, column=0, columnspan=4, pady=4)
        self.btn_dl_2a = ctk.CTkButton(btn_row, text="DOWNLOAD GSTR-2A", command=self._on_download_gstr2a, width=200, height=34, fg_color=TEAL, hover_color="#047a96")
        self.btn_dl_2a.pack(side="left", padx=5)
        self.btn_stop_2a = ctk.CTkButton(btn_row, text="STOP", command=self._on_stop, fg_color=DANGER, width=80, height=34)
        self.btn_stop_2a.pack(side="left", padx=5)

        self.lbl_2a_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_2a_progress.grid(row=4, column=0, columnspan=4, sticky="w", padx=10)

    def _fill_tab_2b(self):
        tab = self.tabview.tab("GSTR-2B")
        tab.grid_columnconfigure((1,3), weight=1)

        mf = ctk.CTkFrame(tab, fg_color="transparent")
        mf.grid(row=0, column=0, columnspan=4, pady=(8,3))
        self.var_2b_mode = tk.StringVar(value="M")
        ctk.CTkRadioButton(mf, text="Monthly", variable=self.var_2b_mode, value="M", command=self._on_2b_mode_change).pack(side="left", padx=10)
        ctk.CTkRadioButton(mf, text="Quarterly (QRMP)", variable=self.var_2b_mode, value="Q", command=self._on_2b_mode_change).pack(side="left", padx=10)

        sf = ctk.CTkFrame(tab, fg_color="transparent")
        sf.grid(row=1, column=0, columnspan=4, pady=3)
        self.var_2b_scope = tk.StringVar(value="single")
        ctk.CTkRadioButton(sf, text="Single Period", variable=self.var_2b_scope, value="single").pack(side="left", padx=10)
        ctk.CTkRadioButton(sf, text="Full Financial Year", variable=self.var_2b_scope, value="year", text_color="#ffcc00").pack(side="left", padx=10)

        self._2b_period_host = ctk.CTkFrame(tab, fg_color="transparent")
        self._2b_period_host.grid(row=2, column=0, columnspan=4, sticky="ew")

        self._2b_month_f = ctk.CTkFrame(self._2b_period_host, fg_color="transparent")
        ctk.CTkLabel(self._2b_month_f, text="Month").pack(side="left", padx=10)
        self.var_2b_month = tk.StringVar(value="April")
        ctk.CTkOptionMenu(self._2b_month_f, values=[m[0] for m in MONTHS], variable=self.var_2b_month).pack(side="left", fill="x", expand=True)

        self._2b_quarter_f = ctk.CTkFrame(self._2b_period_host, fg_color="transparent")
        ctk.CTkLabel(self._2b_quarter_f, text="Quarter").pack(side="left", padx=10)
        self.var_2b_quarter = tk.StringVar(value=QUARTERS[0][0])
        ctk.CTkOptionMenu(self._2b_quarter_f, values=[q[0] for q in QUARTERS], variable=self.var_2b_quarter).pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(tab, text="Year").grid(row=3, column=0, sticky="e", padx=10, pady=3)
        self.var_2b_year = tk.StringVar(value=str(CURRENT_YEAR))
        ctk.CTkOptionMenu(tab, values=YEARS, variable=self.var_2b_year).grid(row=3, column=1, sticky="ew", pady=3)

        self.var_2b_skip = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(tab, text="Skip existing files", variable=self.var_2b_skip, font=ctk.CTkFont(size=11)).grid(row=3, column=2, columnspan=2, padx=10)

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=4, column=0, columnspan=4, pady=6)
        self.btn_dl_2b = ctk.CTkButton(btn_row, text="DOWNLOAD GSTR-2B", command=self._on_download_gstr2b, width=200, height=38, fg_color=ORANGE, hover_color="#b83201")
        self.btn_dl_2b.pack(side="left", padx=5)
        self.btn_stop_2b = ctk.CTkButton(btn_row, text="STOP", command=self._on_stop, fg_color=DANGER, width=80, height=38)
        self.btn_stop_2b.pack(side="left", padx=5)

        self.lbl_2b_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_2b_progress.grid(row=5, column=0, columnspan=4, sticky="w", padx=10)
        self._on_2b_mode_change()

    def _fill_tab_3b(self):
        tab = self.tabview.tab("GSTR-3B")
        tab.grid_columnconfigure((1,3), weight=1)

        ctk.CTkLabel(tab, text="Month").grid(row=0, column=0, sticky="e", padx=10, pady=(8,3))
        self.var_3b_month = tk.StringVar(value="April")
        ctk.CTkOptionMenu(tab, values=[m[0] for m in MONTHS], variable=self.var_3b_month).grid(row=0, column=1, sticky="ew", pady=(8,3))

        ctk.CTkLabel(tab, text="Year").grid(row=0, column=2, sticky="e", padx=10)
        self.var_3b_year = tk.StringVar(value=str(CURRENT_YEAR))
        ctk.CTkOptionMenu(tab, values=YEARS, variable=self.var_3b_year).grid(row=0, column=3, sticky="ew", pady=(8,3))

        sf = ctk.CTkFrame(tab, fg_color="transparent")
        sf.grid(row=1, column=0, columnspan=4, pady=6)
        self.var_3b_scope = tk.StringVar(value="single")
        ctk.CTkRadioButton(sf, text="Single Month", variable=self.var_3b_scope, value="single").pack(side="left", padx=10)
        ctk.CTkRadioButton(sf, text="Full Financial Year", variable=self.var_3b_scope, value="year", text_color="#ffcc00").pack(side="left", padx=10)

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=2, column=0, columnspan=4, pady=8)
        self.btn_dl_3b = ctk.CTkButton(btn_row, text="DOWNLOAD GSTR-3B", command=self._on_download_gstr3b, width=200, height=38, fg_color=GREEN, hover_color="#268635")
        self.btn_dl_3b.pack(side="left", padx=5)
        self.btn_stop_3b = ctk.CTkButton(btn_row, text="STOP", command=self._on_stop, fg_color=DANGER, width=80, height=38)
        self.btn_stop_3b.pack(side="left", padx=5)

        self.lbl_3b_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_3b_progress.grid(row=3, column=0, columnspan=4, sticky="w", padx=10)

    def _build_log_panel(self):
        log_frame = ctk.CTkFrame(self._right,
                                  fg_color=("gray96", "#1e293b"),
                                  corner_radius=10,
                                  border_width=1,
                                  border_color=("gray85", "#334155"))
        log_frame.grid(row=2, column=0, sticky="nsew")

        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(log_frame, text="ACTIVITY LOG",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=("#475569", "#94a3b8")).grid(row=0, column=0, sticky="w", padx=15, pady=5)
        
        self.txt_log = ctk.CTkTextbox(log_frame, font=("Consolas", 12), border_width=1, border_color=("gray75", "#333333"))
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.txt_log.configure(state="disabled")

        # Tags in CTK Textbox? CTK Textbox is based on tk.Text but might need access to underlying widget
        self.txt_log._textbox.tag_configure("info",    foreground="#888")
        self.txt_log._textbox.tag_configure("success", foreground="#4eacff", font=("Consolas", 12, "bold"))
        self.txt_log._textbox.tag_configure("error",   foreground="#ff5252", font=("Consolas", 12, "bold"))
        self.txt_log._textbox.tag_configure("warn",    foreground="#ffcc00")
        self.txt_log._textbox.tag_configure("step",    foreground="#4eacff", font=("Consolas", 12, "bold"))
        self.txt_log._textbox.tag_configure("skip",    foreground="#555")

        ctk.CTkButton(log_frame, text="Clear Log", width=100, height=25, fg_color="transparent", border_width=1, command=self._clear_log).grid(row=0, column=0, sticky="e", padx=15)

    # =========================================================================
    # State machine
    # =========================================================================

    def _set_state(self, state: str):
        # State mapping for buttons
        s = {
            "idle":        ("normal",   "disabled", "disabled"),
            "captcha_ok":  ("normal",   "normal",   "disabled"),
            "otp_wait":    ("disabled", "normal",   "disabled"),
            "logged_in":   ("disabled", "disabled", "normal"),
            "downloading": ("disabled", "disabled", "disabled"),
            "yearly":      ("disabled", "disabled", "disabled"),
            "done":        ("normal",   "disabled", "normal"),
            "busy":        ("disabled", "disabled", "disabled"),
            "error":       ("normal",   "disabled", "disabled"),
        }.get(state, ("normal", "disabled", "disabled"))

        self.btn_captcha.configure(state=s[0])
        self.btn_login.configure(  state=s[1])
        self.btn_logout.configure( state=s[2])

        dl_st   = "normal"   if state in ("logged_in", "done") else "disabled"
        stop_st = "normal"   if state == "yearly"              else "disabled"
        for b in (self.btn_dl_1, self.btn_dl_2a, self.btn_dl_2b, self.btn_dl_3b):
            b.configure(state=dl_st)
        for b in (self.btn_stop_1, self.btn_stop_2a, self.btn_stop_2b, self.btn_stop_3b):
            b.configure(state=stop_st)

        labels = {
            "idle":        ("Ready — Get Captcha to begin.", "#888"),
            "captcha_ok":  ("Captcha loaded — Enter text and Login.", "#4eacff"),
            "otp_wait":    ("OTP sent — Check your device.", "#ffcc00"),
            "logged_in":   ("Logged in — Select a form to download.", "#2ea043"),
            "downloading": ("Downloading... Please wait.", "#ffcc00"),
            "yearly":      ("Yearly download in progress.", "#8250df"),
            "done":        ("Process complete.", "#2ea043"),
            "busy":        ("Please wait...", "#888"),
            "error":       ("Error occurred.", "#ff5252"),
        }
        txt, clr = labels.get(state, ("", "#888"))
        self.lbl_status.configure(text=txt, text_color=clr)

        if state in ("downloading", "busy"):
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        elif state == "yearly":
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(0)
        else:
            self.progress.stop()
            self.progress.set(0)

        if state == "otp_wait":
            self.otp_frame.pack(fill="x", pady=10, padx=20, before=self._action_frame)
            self.entry_otp.focus_set()
        else:
            self.otp_frame.pack_forget()

    # =========================================================================
    # Events & Logic (Kept Original)
    # =========================================================================

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.var_output.set(folder)

    def _load_app_config(self) -> dict:
        try:
            if os.path.exists(self._app_config_path):
                with open(self._app_config_path, encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_app_config(self):
        try:
            with open(self._app_config_path, "w", encoding="utf-8") as f:
                json.dump(self._app_config, f, indent=2)
        except Exception:
            pass

    def _clear_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _on_stop(self):
        self._yearly_stop = True
        self._log("Stop requested — will stop after current period.", "warn")
        for b in (self.btn_stop_1, self.btn_stop_2a, self.btn_stop_2b, self.btn_stop_3b):
            b.configure(state="disabled")

    def _on_get_captcha(self):
        self._set_state("busy")
        self._log("Connecting to GST portal...", "step")
        threading.Thread(target=self._worker_captcha, daemon=True).start()

    def _worker_captcha(self):
        try:
            out  = self.var_output.get() or str(Path(__file__).parent)
            log  = os.path.join(out, "gst_portal_log.txt")
            self._session  = make_session()
            self._login_dl = Gstr3BDownloader(log_path=log, session=self._session)
            b64 = self._login_dl.get_captcha_base64()
            self.after(0, self._show_captcha, b64)
        except Exception as ex:
            self.after(0, self._log, f"Captcha error: {ex}", "error")
            self.after(0, self._set_state, "error")

    def _show_captcha(self, b64: str):
        try:
            raw = base64.b64decode(b64)
            if PIL_AVAILABLE:
                img   = Image.open(io.BytesIO(raw))
                img   = img.resize((200, 65), Image.LANCZOS)
                photo = ctk.CTkImage(light_image=img, dark_image=img, size=(200, 65))
                self.captcha_label.configure(image=photo, text="")
                self.captcha_label.image = photo
            else:
                out = self.var_output.get()
                os.makedirs(out, exist_ok=True)
                p = os.path.join(out, "captcha.png")
                with open(p, "wb") as f:
                    f.write(raw)
                self.captcha_label.configure(text=f"Saved: {p}\n(install Pillow)", text_color="#ff5252")
            self.var_captcha.set("")
            self.entry_captcha.focus_set()
            self._log("Captcha loaded.", "success")
            self._set_state("captcha_ok")
        except Exception as ex:
            self._log(f"Display error: {ex}", "error")
            self._set_state("error")

    def _on_login(self):
        if self._otp_pending:
            otp = self.var_otp.get().strip()
            if not otp:
                messagebox.showwarning("OTP", "Enter the OTP.")
                return
            self._set_state("busy")
            self._log("Submitting OTP...", "step")
            threading.Thread(target=self._worker_otp, args=(otp,), daemon=True).start()
        else:
            if not self._loaded_credentials:
                messagebox.showwarning("No Profile",
                                       "Please load a GST profile first using 'Load ID Pass'.")
                return
            u = self.var_username.get().strip()
            p = self.var_password.get().strip()
            c = self.var_captcha.get().strip()
            if not u or not p:
                messagebox.showwarning("Missing", "Loaded profile is missing Username or Password.")
                return
            if not c:
                messagebox.showwarning("Missing", "Enter the captcha text.")
                return
            self._set_state("busy")
            self._log(f"Logging in as {u}...", "step")
            threading.Thread(target=self._worker_login, args=(u, p, c), daemon=True).start()

    def _worker_login(self, u, p, c):
        try:
            r = self._login_dl.login(u, p, c)
            self.after(0, self._on_login_result, r)
        except Exception as ex:
            self.after(0, self._log, f"Login error: {ex}", "error")
            self.after(0, self._set_state, "error")

    def _worker_otp(self, otp):
        try:
            r = self._login_dl.login_with_otp(otp)
            self.after(0, self._on_login_result, r)
        except Exception as ex:
            self.after(0, self._log, f"OTP error: {ex}", "error")
            self.after(0, self._set_state, "error")

    def _on_login_result(self, r):
        if r.otp_required:
            self._otp_pending = True
            self._log("OTP required — check mobile/email.", "warn")
            self._set_state("otp_wait")
        elif r.is_success:
            self._otp_pending = False
            self._logged_in   = True
            self._log("Login successful.", "success")
            self._set_state("logged_in")
            
            def _fetch_profile_task():
                try:
                    profile = self._login_dl.fetch_profile()
                    if profile and (profile.get("lgl_nm") or profile.get("trdnm") or profile.get("bname")):
                        self._profile = profile
                        name = profile.get("lgl_nm") or profile.get("trdnm") or profile.get("bname")
                        self.after(0, self._log, f"Profile Linked: {name}", "success")
                    else:
                        self.after(0, self._log, "Profile linked but Name not found.", "warn")
                except Exception as e:
                    self.after(0, self._log, f"Background profile fetch failed: {e}", "info")
            threading.Thread(target=_fetch_profile_task, daemon=True).start()
        else:
            self._otp_pending = False
            self._log(f"Login failed: {r.message}", "error")
            self._set_state("captcha_ok")

    def _on_logout(self):
        if self._login_dl:
            try:
                self._login_dl.logout()
            except Exception: pass
        self._login_dl    = None
        self._session     = None
        self._logged_in   = False
        self._otp_pending = False
        self.var_captcha.set("")
        self.var_otp.set("")
        self.captcha_label.configure(image=None, text="[ Captcha Image ]")
        # Keep credentials loaded so user can re-login without re-selecting
        for lbl_ in (self.lbl_1_progress, self.lbl_2a_progress, self.lbl_2b_progress, self.lbl_3b_progress):
            lbl_.configure(text="")
        self._log("Logged out.", "info")
        self._set_state("idle")

    def _user(self):
        return self.var_username.get().strip() or "GSTN"

    def _month_val(self, var: tk.StringVar) -> str:
        sel = var.get()
        for label, val in MONTHS:
            if label == sel: return val
        return "3"

    def _quarter_val(self) -> str:
        sel = self.var_2b_quarter.get()
        for label, val in QUARTERS:
            if label == sel: return val
        return "6"

    def _on_2b_mode_change(self):
        if self.var_2b_mode.get() == "M":
            self._2b_quarter_f.pack_forget()
            self._2b_month_f.pack(fill="x", padx=10)
        else:
            self._2b_month_f.pack_forget()
            self._2b_quarter_f.pack(fill="x", padx=10)

    def _set_det_progress(self, val):
        if hasattr(self, 'progress'):
            max_val = self.progress.cget("maximum") if hasattr(self.progress, "cget") else 12
            self.progress.set(val / max_val)

    def _finish_yearly(self, progress_lbl, ok, skip, fail, total):
        progress_lbl.configure(text=f"Done — {ok} downloaded, {skip} skipped, {fail} failed")
        self._log(f"Yearly process finished. ({ok} OK, {skip} Skipped, {fail} Failed)", "success" if fail == 0 else "warn")
        self._set_state("done")

    # ── Folder Helpers ────────────────────────────────────────────────────────
    
    def _json_dir(self, form: str, year: int) -> str:
        base = self.var_output.get().strip() or str(Path(__file__).parent)
        fy   = f"FY {year}-{str(year + 1)[2:]}"
        d    = os.path.join(base, self._user(), form, fy, "JSON")
        os.makedirs(d, exist_ok=True)
        return d

    def _excel_dir(self, form: str, year: int) -> str:
        base = self.var_output.get().strip() or str(Path(__file__).parent)
        fy   = f"FY {year}-{str(year + 1)[2:]}"
        d    = os.path.join(base, self._user(), form, fy, "Excel")
        os.makedirs(d, exist_ok=True)
        return d

    # ── Excel & Conversion ───────────────────────────────────────────────────

    def _run_excel_convert(self, form: str, json_path: str, excel_path: str, profile: dict = None):
        try:
            from gst_excel_utils import XLSX_OK
            if not XLSX_OK:
                self.after(0, self._log, "openpyxl not installed — skipping Excel.", "warn")
                return
            with open(json_path, encoding="utf-8") as fh:
                data = json.load(fh)
            if form == "GSTR-1":
                datas = [data]
                self._enrich_datas(datas)
                from gstr1_excel import gstr1_consolidated_to_excel
                gstr1_consolidated_to_excel(datas, excel_path, profile=profile)
            elif form == "GSTR-2A":
                from gstr2a_excel import gstr2a_to_excel
                gstr2a_to_excel(data, excel_path, profile=profile)
            elif form == "GSTR-2B":
                from gstr2b_excel import gstr2b_to_excel
                gstr2b_to_excel(data, excel_path, profile=profile)
            elif form == "GSTR-3B":
                from gstr3b_excel import gstr3b_to_excel
                gstr3b_to_excel(data, excel_path, profile=profile)
            self.after(0, self._log, f"Excel Created: {Path(excel_path).name}", "success")
        except Exception as ex:
            self.after(0, self._log, f"Excel Error: {ex}", "error")

    def _maybe_excel(self, form: str, year: int, json_paths: list):
        if not self.var_fmt_excel.get(): return
        excel_dir = self._excel_dir(form, year)
        def _worker():
            for jp in json_paths:
                ep = os.path.join(excel_dir, Path(jp).stem + ".xlsx")
                self._run_excel_convert(form, jp, ep, profile=self._profile)
        threading.Thread(target=_worker, daemon=True).start()

    def _batch_excel(self, form: str, json_dir: str, json_paths: list):
        if not self.var_fmt_excel.get(): return
        import glob as _glob
        dir_files = sorted(_glob.glob(os.path.join(json_dir, "*.json")))
        session_files = [jp for jp in json_paths if os.path.exists(jp)]
        seen = set()
        all_paths = []
        for p in dir_files + session_files:
            norm = os.path.normcase(os.path.abspath(p))
            if norm not in seen:
                seen.add(norm)
                all_paths.append(p)
        if not all_paths: return
        xl_dir = str(Path(json_dir).parent / "Excel")
        os.makedirs(xl_dir, exist_ok=True)
        mode = self.var_excel_mode.get()
        if mode == "consolidated":
            self.after(0, self._log, f"Generating Consolidated Excel ({len(all_paths)} periods)...", "step")
            self._run_consolidated_excel(form, all_paths, xl_dir)
        else:
            self.after(0, self._log, f"Generating {len(all_paths)} Individual Excel files...", "step")
            for jp in all_paths:
                ep = os.path.join(xl_dir, Path(jp).stem + ".xlsx")
                self._run_excel_convert(form, jp, ep, profile=self._profile)

    def _run_consolidated_excel(self, form: str, json_paths: list, xl_dir: str):
        try:
            datas = []
            for jp in json_paths:
                with open(jp, encoding="utf-8") as f:
                    datas.append(json.load(f))
            
            self._enrich_datas(datas)
            
            fy = ""
            for part in Path(json_paths[0]).parts:
                if part.upper().startswith("FY "):
                    fy = part[3:].strip(); break
            profile = dict(self._profile)
            if fy: profile["fy"] = fy

            user = self._user()
            out_file = f"{form}_{user}_Yearly_Consolidated.xlsx"
            out_path = os.path.join(xl_dir, out_file)

            if form == "GSTR-1": from gstr1_excel import gstr1_consolidated_to_excel as fn
            elif form == "GSTR-2A": from gstr2a_excel import gstr2a_consolidated_to_excel as fn
            elif form == "GSTR-2B": from gstr2b_excel import gstr2b_consolidated_to_excel as fn
            elif form == "GSTR-3B": from gstr3b_excel import gstr3b_consolidated_to_excel as fn
            else: return
            
            fn(datas, out_path, profile=profile)
            self.after(0, self._log, f"Consolidated Saved: {out_file}", "success")
        except Exception as ex:
            self.after(0, self._log, f"Consolidated Error: {ex}", "error")

    def _enrich_datas(self, datas: list):
        name_cache = self._app_config.setdefault("gstin_name_cache", {})
        missing_ctins = set()
        for d in datas:
            if not isinstance(d, dict): continue
            # Handle both ZIP mode (top-level keys) and API mode (sections-level keys)
            secs = d.get("sections") or d
            if not isinstance(secs, dict): continue
            
            for sec_key in ["b2b", "cdnr", "b2ba", "cdnra"]:
                for party in (secs.get(sec_key) or []):
                    if not isinstance(party, dict): continue
                    ctin = party.get("ctin")
                    if ctin and not party.get("trdnm") and ctin not in name_cache:
                        missing_ctins.add(ctin)
        
        if missing_ctins:
            self._log(f"Fetching {len(missing_ctins)} missing customer names from GST portal...", "info")
            newly_found = self._lookup_party_names(list(missing_ctins))
            if newly_found:
                name_cache.update(newly_found); self._save_app_config()
        
        for d in datas:
            if not isinstance(d, dict): continue
            secs = d.get("sections") or d
            if not isinstance(secs, dict): continue
            
            for sec_key in ["b2b", "cdnr", "b2ba", "cdnra"]:
                # Handle both raw lists and JSON-string encoded sections
                sec_val = secs.get(sec_key)
                if isinstance(sec_val, str):
                    try:
                        inner = json.loads(sec_val)
                        party_list = inner.get(sec_key) or inner.get("data", {}).get("cpty") or []
                    except: party_list = []
                else:
                    party_list = sec_val or []

                if not isinstance(party_list, list): continue

                for party in party_list:
                    if not isinstance(party, dict): continue
                    ctin = party.get("ctin")
                    if ctin and not party.get("trdnm"): 
                        party["trdnm"] = name_cache.get(ctin, "")

    def _on_refresh_names(self):
        self._set_state("busy")
        self._log("Refreshing party names...", "step")
        def _task():
            try:
                year = self.var_1_year.get()
                json_dir = self._json_dir("GSTR-1", int(year))
                if not os.path.exists(json_dir): return
                all_ctins = set()
                for f in os.listdir(json_dir):
                    if f.endswith(".json"):
                        with open(os.path.join(json_dir, f), encoding="utf-8") as jf:
                            data = json.load(jf)
                            for sec in ["b2b", "cdnr"]:
                                for p in data.get(sec, []):
                                    if p.get("ctin"): all_ctins.add(p["ctin"])
                missing = [c for c in all_ctins if c not in self._app_config.get("gstin_name_cache", {})]
                if missing:
                    self._log(f"Looking up {len(missing)} names...", "info")
                    res = self._lookup_party_names(missing)
                    self._app_config.setdefault("gstin_name_cache", {}).update(res)
                    self._save_app_config()
                    self._log(f"Updated {len(res)} names.", "success")
            except Exception as e: self._log(f"Refresh failed: {e}", "error")
            finally: self.after(0, self._set_state, "done")
        threading.Thread(target=_task, daemon=True).start()

    def _lookup_party_names(self, ctins: list) -> dict:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        found = {}
        sess = self._session if self._session else make_session()
        is_auth = self._logged_in
        def _fetch_one(ctin):
            if self._yearly_stop: return None
            if is_auth:
                try:
                    resp = sess.post("https://publicservices.gst.gov.in/publicservices/auth/api/search/tp", json={"gstin": ctin}, timeout=10, verify=False, headers={"Referer": "https://services.gst.gov.in/services/auth/searchtp"})
                    if resp.status_code == 200:
                        d = resp.json().get("data", resp.json())
                        return (ctin, d.get("trade_name") or d.get("lgl_nm") or d.get("bname") or "")
                except: pass
            try:
                resp = sess.get(f"https://services.gst.gov.in/services/api/search/taxpayerDetails?gstin={ctin}", timeout=10, verify=False)
                if resp.status_code == 200:
                    d = resp.json().get("data", resp.json())
                    return (ctin, d.get("tradeNam") or d.get("lgnm") or "")
            except: pass
            return None
        with ThreadPoolExecutor(max_workers=10) as executor:
            for future in as_completed({executor.submit(_fetch_one, c): c for c in ctins}):
                res = future.result(); 
                if res: found[res[0]] = res[1]
        return found

    # ── GSTR-1 Downloads ──────────────────────────────────────────────────────

    def _on_download_gstr1(self):
        year = self.var_1_year.get().strip()
        if not year or not year.isdigit():
            messagebox.showwarning("Input Error", "Please select a valid Financial Year.")
            return
        
        self._log(f"Initiating GSTR-1 Download for FY {year}...", "step")
        if self.var_1_scope.get() == "year": 
            self._start_gstr1_yearly(int(year))
        else: 
            self._start_gstr1_single(self._month_val(self.var_1_month), int(year))

    def _start_gstr1_single(self, period, year):
        self._set_state("downloading")
        self._log(f"GSTR-1 — {MONTH_NAMES.get(period,period)} {year}", "step")
        actual_year = year + 1 if int(period) <= 3 else year
        threading.Thread(target=self._worker_gstr1, args=(period, actual_year, self.var_1_gstin.get().strip()), daemon=True).start()

    def _worker_gstr1(self, period, year, gstin):
        try:
            d = Gstr1Downloader(session=self._session, log_callback=lambda m: self.after(0, self._log, m, "info"))
            res = d.download_gstr1(period, year, gstin=gstin)
            self.after(0, self._on_gstr1_done, res, period, year)
        except Exception as e: self.after(0, self._log, f"Error: {e}", "error"); self.after(0, self._set_state, "done")

    def _on_gstr1_done(self, result, period, year):
        try:
            paths = save_gstr1(result, self._json_dir("GSTR-1", year), self._user())
            self._log(f"Saved {len(paths)} files.", "success")
            self._maybe_excel("GSTR-1", year, paths)
        except Exception as e: self._log(f"Save error: {e}", "error")
        self._set_state("done")

    def _start_gstr1_yearly(self, year):
        self._yearly_stop = False; self._set_state("yearly")
        self._log(f"GSTR-1 Full Year — FY {year}", "step")
        threading.Thread(target=self._worker_gstr1_yearly, args=(year, self.var_1_gstin.get().strip()), daemon=True).start()

    def _worker_gstr1_yearly(self, year, gstin):
        from concurrent.futures import ThreadPoolExecutor, as_completed
        out = self._json_dir("GSTR-1", year); user = self._user()
        ok = fail = skip = 0; saved_jsons = []
        
        self._log(f"Starting GSTR-1 Yearly pipeline for {user} (FY {year})...", "info")

        # Phase 1: Pre-triggering (Mass Trigger)
        self._log("Phase 1/2: Triggering offline ZIP generation for all 12 months...", "step")
        trigger_sess = make_session()
        trigger_sess.cookies.update(self._session.cookies)
        trigger_dl = Gstr1Downloader(session=trigger_sess, log_callback=None)
        
        def _trigger_task(p):
            ay = year + 1 if int(p) <= 3 else year
            try: trigger_dl.trigger_offline_gen(p, ay); return True
            except: return False
            
        with ThreadPoolExecutor(max_workers=6) as trigger_exec:
            trigger_exec.map(_trigger_task, MONTHLY_FY_ORDER)
        
        self._log("All months triggered. Moving to download phase...", "success")

        # Phase 2: Concurrent Download
        def _download_month_task(period, idx):
            actual_year = year + 1 if int(period) <= 3 else year
            name = MONTH_NAMES.get(period, period)
            try:
                thread_sess = make_session()
                thread_sess.cookies.update(self._session.cookies)
                d = Gstr1Downloader(session=thread_sess, log_callback=lambda m: self.after(0, self._log, f"[{name}] {m}", "info"))
                d._yearly_stop_source = self
                res = d.download_gstr1(period, actual_year, gstin=gstin)
                res.year = actual_year
                saved = save_gstr1(res, out, user)
                return {"status": "ok", "name": name, "saved": saved, "idx": idx}
            except Exception as e:
                self.after(0, self._log, f"[{name}] Fatal error: {e}", "error")
                return {"status": "fail", "name": name, "error": str(e), "idx": idx}

        self._log("Phase 2/2: Starting concurrent downloads (4 parallel months)...", "info")
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(_download_month_task, p, i) for i, p in enumerate(MONTHLY_FY_ORDER, 1)]
            for future in as_completed(futures):
                if self._yearly_stop: 
                    self._log("Yearly download stop requested by user.", "warn")
                    break
                try:
                    res = future.result()
                    if res["status"] == "ok":
                        ok += 1
                        saved_jsons.extend(res["saved"])
                        if self.var_fmt_excel.get() and self.var_excel_mode.get() == "individual":
                            self._maybe_excel("GSTR-1", year, res["saved"])
                    else:
                        self._log(f"Period {res['name']} failed: {res['error']}", "error")
                        fail += 1
                except Exception as fe:
                    self._log(f"Future execution error: {fe}", "error")
                    fail += 1
                
                self.after(0, self._set_det_progress, ok + fail + skip)
                self.after(0, self.lbl_1_progress.configure, {"text": f"Progress: {ok+fail+skip}/12"})

        if not self._yearly_stop and self.var_excel_mode.get() == "consolidated": 
            self._batch_excel("GSTR-1", out, saved_jsons)
        
        self.after(0, self._finish_yearly, self.lbl_1_progress, ok, skip, fail, 12)

    # ── GSTR-2A Downloads ──────────────────────────────────────────────────────

    def _on_download_gstr2a(self):
        year = self.var_2a_year.get().strip()
        if not year.isdigit(): return
        sections = [k for k, v in self._2a_section_vars.items() if v.get()]
        if not sections: return
        if self.var_2a_scope.get() == "year": self._start_gstr2a_yearly(int(year), sections)
        else: self._start_gstr2a_single(self._month_val(self.var_2a_month), int(year), sections)

    def _start_gstr2a_single(self, period, year, sections):
        self._set_state("downloading")
        self._log(f"GSTR-2A — {MONTH_NAMES.get(period,period)} {year}", "step")
        actual_year = year + 1 if int(period) <= 3 else year
        threading.Thread(target=self._worker_gstr2a, args=(period, actual_year, sections), daemon=True).start()

    def _worker_gstr2a(self, period, year, sections):
        try:
            d = Gstr2ADownloader(session=self._session)
            res = d.download_gstr2a(period, year, sections, progress_callback=lambda m: self.after(0, self._log, m, "info"))
            self.after(0, self._on_gstr2a_done, res, period, year)
        except Exception as e: self.after(0, self._log, f"Error: {e}", "error"); self.after(0, self._set_state, "done")

    def _on_gstr2a_done(self, result, period, year):
        try:
            p = save_gstr2a(result, self._json_dir("GSTR-2A", year), self._user())
            self._log(f"Saved: {p}", "success")
            self._maybe_excel("GSTR-2A", year, [p])
        except Exception as e: self._log(f"Save error: {e}", "error")
        self._set_state("done")

    def _start_gstr2a_yearly(self, year, sections):
        self._yearly_stop = False; self._set_state("yearly")
        threading.Thread(target=self._worker_gstr2a_yearly, args=(year, sections), daemon=True).start()

    def _worker_gstr2a_yearly(self, year, sections):
        out = self._json_dir("GSTR-2A", year); user = self._user()
        ok = fail = skip = 0; saved_jsons = []
        for i, period in enumerate(MONTHLY_FY_ORDER, 1):
            if self._yearly_stop: break
            actual_year = year + 1 if int(period) <= 3 else year
            name = MONTH_NAMES.get(period, period)
            self.after(0, self.lbl_2a_progress.configure, {"text": f"[{i}/12] {name}"})
            try:
                d = Gstr2ADownloader(session=self._session)
                res = d.download_gstr2a(period, actual_year, sections, progress_callback=lambda m: self.after(0, self._log, f"  {m}", "info"))
                res.year = actual_year
                jp = save_gstr2a(res, out, user); saved_jsons.append(jp)
                ok += 1
            except Exception as e: self._log(f"Fail {name}: {e}", "error"); fail += 1
            self.after(0, self._set_det_progress, i)
        self._batch_excel("GSTR-2A", out, saved_jsons)
        self.after(0, self._finish_yearly, self.lbl_2a_progress, ok, skip, fail, 12)

    # ── GSTR-2B Downloads ──────────────────────────────────────────────────────

    def _on_download_gstr2b(self):
        mode = self.var_2b_mode.get(); year = self.var_2b_year.get().strip()
        if not year.isdigit(): return
        period = self._month_val(self.var_2b_month) if mode == "M" else self._quarter_val()
        if self.var_2b_scope.get() == "year":
            self._start_gstr2b_yearly(int(year), mode, MONTHLY_FY_ORDER if mode == "M" else QUARTERLY_FY_ORDER)
        else: self._start_gstr2b_single(period, int(year), mode)

    def _start_gstr2b_single(self, period, year, mode):
        self._set_state("downloading")
        self._log(f"GSTR-2B — {period} {year}", "step")
        actual_year = year + 1 if int(period) <= 3 else year
        threading.Thread(target=self._worker_gstr2b, args=(period, actual_year, mode), daemon=True).start()

    def _worker_gstr2b(self, period, year, mode):
        try:
            d = Gstr2BDownloader(session=self._session)
            res = d.download_gstr2b(period, year, mode)
            self.after(0, self._on_gstr2b_done, res, period, year, mode)
        except Exception as e: self.after(0, self._log, f"Error: {e}", "error"); self.after(0, self._set_state, "done")

    def _on_gstr2b_done(self, result, period, year, mode):
        try:
            jdir = self._json_dir("GSTR-2B", year); save_2b(result, jdir, self._user(), period, year)
            prd = str(period).zfill(2)
            paths = [os.path.join(jdir, f"GSTR2B_Return_{self._user()}_{year}_{prd}.json")]
            self._maybe_excel("GSTR-2B", year, [p for p in paths if os.path.exists(p)])
        except Exception as e: self._log(f"Save error: {e}", "error")
        self._set_state("done")

    def _start_gstr2b_yearly(self, year, mode, periods):
        self._yearly_stop = False; self._set_state("yearly")
        threading.Thread(target=self._worker_gstr2b_yearly, args=(year, mode, periods), daemon=True).start()

    def _worker_gstr2b_yearly(self, year, mode, periods):
        out = self._json_dir("GSTR-2B", year); user = self._user(); ok = fail = skip = 0; saved_jsons = []
        for i, period in enumerate(periods, 1):
            if self._yearly_stop: break
            actual_year = year + 1 if int(period) <= 3 else year
            name = (MONTH_NAMES if mode == "M" else QUARTER_NAMES).get(period, period)
            self.after(0, self.lbl_2b_progress.configure, {"text": f"[{i}/{len(periods)}] {name}"})
            try:
                d = Gstr2BDownloader(session=self._session)
                res = d.download_gstr2b(period, actual_year, mode)
                save_2b(res, out, user, period, actual_year)
                prd = str(period).zfill(2)
                saved_jsons.append(os.path.join(out, f"GSTR2B_Return_{user}_{actual_year}_{prd}.json"))
                ok += 1
            except Exception as e: self._log(f"Fail {name}: {e}", "error"); fail += 1
            self.after(0, self._set_det_progress, i)
        self._batch_excel("GSTR-2B", out, saved_jsons)
        self.after(0, self._finish_yearly, self.lbl_2b_progress, ok, skip, fail, len(periods))

    # ── GSTR-3B Downloads ──────────────────────────────────────────────────────

    def _on_download_gstr3b(self):
        year = self.var_3b_year.get().strip()
        if not year.isdigit(): return
        if self.var_3b_scope.get() == "year": self._start_gstr3b_yearly(int(year))
        else: self._start_gstr3b_single(self._month_val(self.var_3b_month), int(year))

    def _start_gstr3b_single(self, period, year):
        self._set_state("downloading")
        actual_year = year + 1 if int(period) <= 3 else year
        threading.Thread(target=self._worker_gstr3b, args=(period, actual_year), daemon=True).start()

    def _worker_gstr3b(self, period, year):
        try:
            res = self._login_dl.download_gstr3b(period, year)
            self.after(0, self._on_gstr3b_done, res, period, year)
        except Exception as e: self.after(0, self._log, f"Error: {e}", "error"); self.after(0, self._set_state, "done")

    def _on_gstr3b_done(self, result, period, year):
        try:
            p = save_gstr3b(result, self._json_dir("GSTR-3B", year), self._user())
            self._maybe_excel("GSTR-3B", year, [p])
        except Exception as e: self._log(f"Save error: {e}", "error")
        self._set_state("done")

    def _start_gstr3b_yearly(self, year):
        self._yearly_stop = False; self._set_state("yearly")
        threading.Thread(target=self._worker_gstr3b_yearly, args=(year,), daemon=True).start()

    def _worker_gstr3b_yearly(self, year):
        out = self._json_dir("GSTR-3B", year); user = self._user(); ok = fail = skip = 0; saved_jsons = []
        for i, period in enumerate(MONTHLY_FY_ORDER, 1):
            if self._yearly_stop: break
            actual_year = year + 1 if int(period) <= 3 else year
            name = MONTH_NAMES.get(period, period)
            self.after(0, self.lbl_3b_progress.configure, {"text": f"[{i}/12] {name}"})
            try:
                res = self._login_dl.download_gstr3b(period, actual_year)
                res.year = actual_year; jp = save_gstr3b(res, out, user); saved_jsons.append(jp); ok += 1
            except Exception as e: self._log(f"Fail {name}: {e}", "error"); fail += 1
            self.after(0, self._set_det_progress, i)
        self._batch_excel("GSTR-3B", out, saved_jsons)
        self.after(0, self._finish_yearly, self.lbl_3b_progress, ok, skip, fail, 12)

    # ── Log ──────────────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = "info"):
        now = datetime.now().strftime("%H:%M:%S")
        line = f"[{now}] {msg}\n"
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", line, tag)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

if __name__ == "__main__":
    app = GstPortalApp()
    app.mainloop()
