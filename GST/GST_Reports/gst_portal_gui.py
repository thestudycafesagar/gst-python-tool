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
YEARS         = [f"{y}-{str(y+1)[2:]}" for y in range(CURRENT_YEAR - 5, CURRENT_YEAR + 1)]

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
        # Default to the current user's Documents folder
        _default_out = os.path.join(
            os.path.expanduser("~"), "Documents"
        )
        os.makedirs(_default_out, exist_ok=True)
        self._output_dir  = _default_out
        self._profile     = {}
        _cfg_dir = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")), "GSTSuite"
        )
        os.makedirs(_cfg_dir, exist_ok=True)
        self._app_config_path = os.path.join(_cfg_dir, "gst_reports_config.json")
        self._app_config  = self._load_app_config()

        self._loaded_credentials = None   # {"Username": ..., "Password": ..., "ClientName": ...}
        self._build_ui()
        self._set_state("idle")
    # =========================================================================
    # UI Components
    # =========================================================================

    def _build_ui(self):
        # Single content row — no header row
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)   # left sidebar
        self.grid_columnconfigure(1, weight=1)   # right panel

        self._left = ctk.CTkScrollableFrame(self, width=330, corner_radius=0,
                                             fg_color=("#f8fafc", "#111827"))
        self._left.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)

        self.card_cred = ctk.CTkFrame(self._left, border_color=("#e2e8f0", "#334155"), border_width=1, corner_radius=6, fg_color=("gray98", "#1e293b"))
        self.card_cred.pack(fill="x", pady=15, padx=15)
        
        ctk.CTkLabel(self.card_cred, text="📂 Credentials Source", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=15, pady=(15, 10))

        # Silently initialize format vars
        self.var_fmt_excel = tk.BooleanVar(value=True)
        self.var_excel_mode = tk.StringVar(value="individual")

        self._build_login_section(self.card_cred)
        self._build_captcha_section(self.card_cred)
        self._build_otp_section(self.card_cred)
        self._build_action_buttons(self.card_cred)

        # ── Right panel (tabs + log) ──────────────────────────────────────────
        self._right = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self._right.grid(row=0, column=1, sticky="nsew", padx=8, pady=6)

        self._right.grid_columnconfigure(0, weight=1)
        self._right.grid_rowconfigure(0, weight=1)   # Tabview expands most
        self._right.grid_rowconfigure(1, weight=0)   # Status/Progress — fixed
        self._right.grid_rowconfigure(2, weight=0)   # Log — fixed height

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
        self.tabview.grid(row=0, column=0, sticky="nsew", pady=(0, 2))
        
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
        status_frame.grid(row=1, column=0, sticky="ew", pady=(0, 2))
        
        self.lbl_status = ctk.CTkLabel(status_frame, text="Ready", font=ctk.CTkFont(weight="bold"))
        self.lbl_status.pack(side="left", padx=5)

        self.progress = ctk.CTkProgressBar(status_frame, orientation="horizontal", height=10)
        self.progress.pack(side="right", fill="x", expand=True, padx=10)
        self.progress.set(0)

        # Log Panel
        self._build_log_panel()

    def _build_login_section(self, parent):
        # Hidden StringVars — populated when a profile is loaded; used by existing login logic
        self.var_username = tk.StringVar()
        self.var_password = tk.StringVar()

        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.pack(fill="x", pady=(0, 10), padx=15)

        # Profile display (read-only, shows loaded profile name / username)
        self._cred_display_var = tk.StringVar(value="Add ID/Password manually...")
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

        # View loaded cred removed per user request

        # Save To (Programmatically set to user's Documents folder, removed from UI)
        self.var_output = tk.StringVar(value=self._output_dir)

    # ── Credential helpers ────────────────────────────────────────────────────

    def _refresh_cred_display(self):
        """Update the display entry and view-button after loading credentials."""
        c = self._loaded_credentials
        if not c:
            self._cred_display_var.set("No profile loaded")
            return
        name = c.get("ClientName") or c.get("Username", "")
        user = c.get("Username", "")
        disp = f"{name} ({user})" if name and name != user else user
        self._cred_display_var.set(disp)
        # Push into hidden vars so login logic works without changes
        self.var_username.set(c.get("Username", ""))
        self.var_password.set(c.get("Password", ""))

        # Dynamically set and freeze filing frequency based on DB profile
        freq = c.get("FilingFrequency")
        if freq in ["Monthly", "Quarterly"]:
            for prefix in ["1", "2a", "2b", "3b"]:
                var_mode = getattr(self, f"var_{prefix}_mode", None)
                tabs = getattr(self, f"_{prefix}_mode_tabs", None)
                toggle = getattr(self, f"_{prefix}_toggle_inputs", None)
                if var_mode and tabs and toggle:
                    var_mode.set(freq)
                    toggle()
                    tabs.configure(state="disabled")

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
            base_disp = f"{c}  ({u})" if c else u
            disp = f"{base_disp} [{ff}]"
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
        # Checkbox removed; Excel is generated by default

        self.var_excel_mode = tk.StringVar(value="individual")
        self.excel_mode_row = ctk.CTkFrame(box, fg_color="transparent")
        self.excel_mode_row.pack(anchor="w", pady=(0, 10), padx=10)
        ctk.CTkRadioButton(self.excel_mode_row, text="Individual", variable=self.var_excel_mode, value="individual").pack(side="left")
        ctk.CTkRadioButton(self.excel_mode_row, text="Consolidated", variable=self.var_excel_mode, value="consolidated").pack(side="left", padx=15)

    def _build_captcha_section(self, parent):
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.pack(fill="x", pady=(5, 10), padx=15)

        self.btn_captcha = ctk.CTkButton(box, text="GET CAPTCHA", command=self._on_get_captcha, height=35, font=ctk.CTkFont(weight="bold"))
        self.btn_captcha.pack(fill="x", pady=(0, 8))

        self.captcha_label = ctk.CTkLabel(box, text="[ Captcha Image ]", height=65, fg_color=("gray85", "#111827"), corner_radius=5)
        self.captcha_label.pack(fill="x", pady=(0, 8))

        row = ctk.CTkFrame(box, fg_color="transparent")
        row.pack(fill="x")
        ctk.CTkLabel(row, text="Captcha:").pack(side="left")
        self.var_captcha = tk.StringVar()
        self.entry_captcha = ctk.CTkEntry(row, textvariable=self.var_captcha, height=32, placeholder_text="CODE", font=ctk.CTkFont(size=14, weight="bold"), justify="center")
        self.entry_captcha.pack(side="right", fill="x", expand=True, padx=(10, 0))

    def _build_otp_section(self, parent):
        self.otp_frame = ctk.CTkFrame(parent, fg_color=("gray93", "#3e3b2e"), border_width=1, border_color="#d4ac0d")
        # Pack/forget dynamically
        
        ctk.CTkLabel(self.otp_frame, text="OTP Verification Required", text_color="#d4ac0d", font=ctk.CTkFont(weight="bold")).pack(pady=5)
        row = ctk.CTkFrame(self.otp_frame, fg_color="transparent")
        row.pack(pady=5, padx=10)
        self.var_otp = tk.StringVar()
        self.entry_otp = ctk.CTkEntry(row, textvariable=self.var_otp, width=120, height=35, justify="center", font=ctk.CTkFont(size=18, weight="bold"))
        self.entry_otp.pack(side="left", padx=5)

    def _build_action_buttons(self, parent):
        self._action_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._action_frame.pack(fill="x", pady=(5, 15), padx=15)
        
        self.btn_login = ctk.CTkButton(self._action_frame, text="LOGIN", command=self._on_login, height=40, fg_color=GREEN, hover_color="#268635", font=ctk.CTkFont(weight="bold"))
        self.btn_login.pack(fill="x", pady=(0, 8))
        
        self.btn_logout = ctk.CTkButton(self._action_frame, text="LOGOUT", command=self._on_logout, height=35, fg_color="#444", hover_color="#555")
        self.btn_logout.pack(fill="x")

    def _build_period_selection(self, tab, prefix):
        card = ctk.CTkFrame(tab, border_color=("#e2e8f0", "#334155"), border_width=1, corner_radius=6, fg_color=("gray95", "#1e293b"))
        
        ctk.CTkLabel(card, text="📅 Period Selection", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=15, pady=(4, 1))

        # Single Row for Parameters
        frm_params = ctk.CTkFrame(card, fg_color="transparent")
        frm_params.pack(fill="x", padx=15, pady=(1, 2))

        # Financial Year
        ctk.CTkLabel(frm_params, text="Financial Year:", width=110, anchor="w").pack(side="left")
        var_year = tk.StringVar(value=YEARS[-1])
        setattr(self, f"var_{prefix}_year", var_year)
        ctk.CTkOptionMenu(frm_params, values=YEARS, variable=var_year, width=160).pack(side="left", padx=(0, 30))
        
        # Filing Frequency
        var_mode = tk.StringVar(value="Monthly")
        setattr(self, f"var_{prefix}_mode", var_mode)
        ctk.CTkLabel(frm_params, text="Filing Frequency:", width=120, anchor="w").pack(side="left")
        
        # Dynamic Checkbox Frame
        frm_checkboxes = ctk.CTkFrame(card, fg_color="transparent")
        frm_checkboxes.pack(fill="both", expand=True, padx=15, pady=1)
        
        checkbox_vars = {}
        setattr(self, f"_{prefix}_period_vars", checkbox_vars)

        def toggle_inputs(mode_choice=None):
            mode = var_mode.get()
            for w in frm_checkboxes.winfo_children():
                w.destroy()
            checkbox_vars.clear()

            top_bar = ctk.CTkFrame(frm_checkboxes, fg_color="transparent")
            top_bar.pack(fill="x", pady=(0, 1))
            
            select_all_var = tk.BooleanVar(value=False)
            def toggle_select_all():
                state = select_all_var.get()
                for v in checkbox_vars.values():
                    v.set(state)
                # Toggle visibility of the external opt_frame
                external_opt = getattr(self, f"_opt_frame_{prefix}", None)
                if external_opt:
                    if state:
                        external_opt.grid(row=2, column=0, sticky="ew", padx=10, pady=(2, 0))
                    else:
                        external_opt.grid_forget()
                        self.var_excel_mode.set("individual")
            
            ctk.CTkCheckBox(top_bar, text="Select All", variable=select_all_var, command=toggle_select_all, font=("Segoe UI", 12, "bold"), text_color="#10B981").pack(side="left")

            chk_grid = ctk.CTkFrame(frm_checkboxes, fg_color="transparent")
            chk_grid.pack(fill="both", expand=True)

            if mode == "Monthly":
                items = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
                labels = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
                cols = 6
            else:
                items = ["6", "9", "12", "3"]
                labels = ["Q1 (Apr-Jun)", "Q2 (Jul-Sep)", "Q3 (Oct-Dec)", "Q4 (Jan-Mar)"]
                cols = 4
            
            for i, (item, label) in enumerate(zip(items, labels)):
                var = tk.BooleanVar(value=False)
                checkbox_vars[item] = var
                chk = ctk.CTkCheckBox(chk_grid, text=label, variable=var, font=("Segoe UI", 12))
                chk.grid(row=i // cols, column=i % cols, padx=5, pady=1, sticky="w")

        mode_tabs = ctk.CTkSegmentedButton(
            frm_params,
            values=["Monthly", "Quarterly"],
            variable=var_mode,
            command=toggle_inputs,
            width=180
        )
        mode_tabs.pack(side="left", padx=10)
        
        setattr(self, f"_{prefix}_mode_tabs", mode_tabs)
        setattr(self, f"_{prefix}_toggle_inputs", toggle_inputs)

        toggle_inputs()
        return card

    def _fill_tab_1(self):
        tab = self.tabview.tab("GSTR-1")
        tab.grid_columnconfigure(0, weight=1)

        # --- 1. Period Selection Card ---
        self._build_period_selection(tab, "1").grid(row=0, column=0, sticky="ew", pady=(4, 2), padx=10)

        # Silent initialization for backend compatibility
        self.var_1_gstin = tk.StringVar()

        # --- Refresh Names Button ---
        # self.btn_refresh_names = ctk.CTkButton(tab, text="Refresh Party Names", command=self._on_refresh_names, width=150, height=28, fg_color=("gray75", "#333333"))
        # self.btn_refresh_names.grid(row=2, column=0, sticky="e", padx=10, pady=5)

        # --- Output Options Row ---
        self._opt_frame_1 = ctk.CTkFrame(tab, fg_color="transparent")
        ctk.CTkLabel(self._opt_frame_1, text="Output Mode:", font=ctk.CTkFont(size=12, weight="bold"), text_color="#94a3b8").pack(side="left", padx=5)
        ctk.CTkRadioButton(self._opt_frame_1, text="Consolidated", variable=self.var_excel_mode, value="consolidated", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)
        ctk.CTkRadioButton(self._opt_frame_1, text="Individual", variable=self.var_excel_mode, value="individual", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)

        # --- 4. Standard Action Row ---
        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=3, column=0, sticky="ew", pady=(2, 2))
        btn_row_center = ctk.CTkFrame(btn_row, fg_color="transparent")
        btn_row_center.pack(anchor="center")

        self.btn_dl_1 = ctk.CTkButton(btn_row_center, text="DOWNLOAD GSTR-1", command=self._on_download_gstr1, width=220, height=38, font=ctk.CTkFont(weight="bold"))
        self.btn_dl_1.pack(side="left", padx=10)
        self.btn_stop_1 = ctk.CTkButton(btn_row_center, text="STOP", command=self._on_stop, fg_color=DANGER, hover_color="#a01a23", width=80, height=38)
        self.btn_stop_1.pack(side="left", padx=10)
        self.btn_open_1 = ctk.CTkButton(btn_row_center, text="📂 OPEN FOLDER", command=lambda: self._open_output_folder("GSTR-1", self.var_1_year), width=120, height=38, fg_color="#475569", hover_color="#334155")
        self.btn_open_1.pack(side="left", padx=10)

        # --- Progress Label ---
        self.lbl_1_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_1_progress.grid(row=4, column=0, sticky="w", padx=10)

    def _fill_tab_2a(self):
        tab = self.tabview.tab("GSTR-2A")
        tab.grid_columnconfigure(0, weight=1)

        # --- 1. Period Selection Card ---
        self._build_period_selection(tab, "2a").grid(row=0, column=0, sticky="ew", pady=(4, 2), padx=10)

        # --- 3. Checkbox Logic (Hidden, default True) ---
        self._2a_section_vars = {}
        sections_2a = [
            ("B2B","b2b"),   ("B2BA","b2ba"),  ("CDN","cdn"),
            ("CDNA","cdna"), ("TDS","tds"),     ("TCS","tcs"),
            ("ISD","isd"),   ("ISDA","isda"),   ("TDSA","tdsa"),
            ("IMPG","impg"), ("IMPGSEZ","impgsez"),
        ]
        for _, key in sections_2a:
            self._2a_section_vars[key] = tk.BooleanVar(value=True)

        # --- Output Options Row ---
        self._opt_frame_2a = ctk.CTkFrame(tab, fg_color="transparent")
        ctk.CTkLabel(self._opt_frame_2a, text="Output Mode:", font=ctk.CTkFont(size=12, weight="bold"), text_color="#94a3b8").pack(side="left", padx=5)
        ctk.CTkRadioButton(self._opt_frame_2a, text="Consolidated", variable=self.var_excel_mode, value="consolidated", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)
        ctk.CTkRadioButton(self._opt_frame_2a, text="Individual", variable=self.var_excel_mode, value="individual", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)

        # --- 4. Standard Action Row ---
        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=3, column=0, sticky="ew", pady=(2, 2))
        btn_row_center = ctk.CTkFrame(btn_row, fg_color="transparent")
        btn_row_center.pack(anchor="center")

        self.btn_dl_2a = ctk.CTkButton(btn_row_center, text="DOWNLOAD GSTR-2A", command=self._on_download_gstr2a, width=220, height=38, fg_color=TEAL, hover_color="#047a96", font=ctk.CTkFont(weight="bold"))
        self.btn_dl_2a.pack(side="left", padx=10)
        self.btn_stop_2a = ctk.CTkButton(btn_row_center, text="STOP", command=self._on_stop, fg_color=DANGER, hover_color="#a01a23", width=80, height=38)
        self.btn_stop_2a.pack(side="left", padx=10)
        self.btn_open_2a = ctk.CTkButton(btn_row_center, text="📂 OPEN FOLDER", command=lambda: self._open_output_folder("GSTR-2A", self.var_2a_year), width=120, height=38, fg_color="#475569", hover_color="#334155")
        self.btn_open_2a.pack(side="left", padx=10)

        # --- Progress Label ---
        self.lbl_2a_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_2a_progress.grid(row=4, column=0, sticky="w", padx=10)

    def _fill_tab_2b(self):
        tab = self.tabview.tab("GSTR-2B")
        tab.grid_columnconfigure(0, weight=1)

        # --- 1. Period Selection Card ---
        self._build_period_selection(tab, "2b").grid(row=0, column=0, sticky="ew", pady=(4, 2), padx=10)

        # Skip Checkbox (Specific to GSTR-2B)
        self.var_2b_skip = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(tab, text="Skip if JSON exists", variable=self.var_2b_skip).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 4))

        # --- Output Options Row ---
        self._opt_frame_2b = ctk.CTkFrame(tab, fg_color="transparent")
        ctk.CTkLabel(self._opt_frame_2b, text="Output Mode:", font=ctk.CTkFont(size=12, weight="bold"), text_color="#94a3b8").pack(side="left", padx=5)
        ctk.CTkRadioButton(self._opt_frame_2b, text="Consolidated", variable=self.var_excel_mode, value="consolidated", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)
        ctk.CTkRadioButton(self._opt_frame_2b, text="Individual", variable=self.var_excel_mode, value="individual", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)

        # --- 4. Standard Action Row ---
        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=3, column=0, sticky="ew", pady=(2, 2))
        btn_row_center = ctk.CTkFrame(btn_row, fg_color="transparent")
        btn_row_center.pack(anchor="center")

        self.btn_dl_2b = ctk.CTkButton(btn_row_center, text="DOWNLOAD GSTR-2B", command=self._on_download_gstr2b, width=220, height=38, fg_color=ORANGE, hover_color="#b83201", font=ctk.CTkFont(weight="bold"))
        self.btn_dl_2b.pack(side="left", padx=10)
        self.btn_stop_2b = ctk.CTkButton(btn_row_center, text="STOP", command=self._on_stop, fg_color=DANGER, hover_color="#a01a23", width=80, height=38)
        self.btn_stop_2b.pack(side="left", padx=10)
        self.btn_open_2b = ctk.CTkButton(btn_row_center, text="📂 OPEN FOLDER", command=lambda: self._open_output_folder("GSTR-2B", self.var_2b_year), width=120, height=38, fg_color="#475569", hover_color="#334155")
        self.btn_open_2b.pack(side="left", padx=10)

        # --- Progress Label ---
        self.lbl_2b_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_2b_progress.grid(row=4, column=0, sticky="w", padx=10)

    def _fill_tab_3b(self):
        tab = self.tabview.tab("GSTR-3B")
        tab.grid_columnconfigure(0, weight=1)

        # --- 1. Period Selection Card ---
        self._build_period_selection(tab, "3b").grid(row=0, column=0, sticky="ew", pady=(4, 2), padx=10)

        # --- Output Options Row ---
        self._opt_frame_3b = ctk.CTkFrame(tab, fg_color="transparent")
        ctk.CTkLabel(self._opt_frame_3b, text="Output Mode:", font=ctk.CTkFont(size=12, weight="bold"), text_color="#94a3b8").pack(side="left", padx=5)
        ctk.CTkRadioButton(self._opt_frame_3b, text="Consolidated", variable=self.var_excel_mode, value="consolidated", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)
        ctk.CTkRadioButton(self._opt_frame_3b, text="Individual", variable=self.var_excel_mode, value="individual", font=ctk.CTkFont(size=12)).pack(side="left", padx=10)

        # --- 4. Standard Action Row ---
        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=3, column=0, sticky="ew", pady=(2, 2))
        btn_row_center = ctk.CTkFrame(btn_row, fg_color="transparent")
        btn_row_center.pack(anchor="center")

        self.btn_dl_3b = ctk.CTkButton(btn_row_center, text="DOWNLOAD GSTR-3B", command=self._on_download_gstr3b, width=220, height=38, fg_color=GREEN, hover_color="#268635", font=ctk.CTkFont(weight="bold"))
        self.btn_dl_3b.pack(side="left", padx=10)
        self.btn_stop_3b = ctk.CTkButton(btn_row_center, text="STOP", command=self._on_stop, fg_color=DANGER, hover_color="#a01a23", width=80, height=38)
        self.btn_stop_3b.pack(side="left", padx=10)
        self.btn_open_3b = ctk.CTkButton(btn_row_center, text="📂 OPEN FOLDER", command=lambda: self._open_output_folder("GSTR-3B", self.var_3b_year), width=120, height=38, fg_color="#475569", hover_color="#334155")
        self.btn_open_3b.pack(side="left", padx=10)

        # --- Progress Label ---
        self.lbl_3b_progress = ctk.CTkLabel(tab, text="", text_color=PURPLE)
        self.lbl_3b_progress.grid(row=4, column=0, sticky="w", padx=10)

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
                     text_color=("#475569", "#94a3b8")).grid(row=0, column=0, sticky="w", padx=15, pady=(3, 1))
        
        self.txt_log = ctk.CTkTextbox(log_frame, font=("Consolas", 12), border_width=1, border_color=("gray75", "#333333"), height=100)
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 6))
        self.txt_log.configure(state="disabled")

        self.txt_log._textbox.tag_configure("info",    foreground="#888")
        self.txt_log._textbox.tag_configure("success", foreground="#4eacff", font=("Consolas", 12, "bold"))
        self.txt_log._textbox.tag_configure("error",   foreground="#ff5252", font=("Consolas", 12, "bold"))
        self.txt_log._textbox.tag_configure("warn",    foreground="#ffcc00")
        self.txt_log._textbox.tag_configure("step",    foreground="#4eacff", font=("Consolas", 12, "bold"))
        self.txt_log._textbox.tag_configure("skip",    foreground="#555")

    def _clear_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

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
        from tkinter import messagebox as _mb
        if not _mb.askyesno("Confirm Logout", "Are you sure you want to log out?", parent=self):
            return
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

    def _set_det_progress(self, val):
        if hasattr(self, 'progress'):
            max_val = self.progress.cget("maximum") if hasattr(self.progress, "cget") else 12
            self.progress.set(val / max_val)

    def _finish_yearly(self, progress_lbl, ok, skip, fail, total):
        progress_lbl.configure(text=f"Done — {ok} downloaded, {skip} skipped, {fail} failed")
        self._log(f"Batch process finished. ({ok} OK, {skip} Skipped, {fail} Failed)", "success" if fail == 0 else "warn")
        self._set_state("done")
        from tkinter import messagebox
        messagebox.showinfo("Process Complete", f"Batch download completed.\n\nSuccessfully downloaded: {ok}\nSkipped: {skip}\nFailed: {fail}")

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

    def _open_output_folder(self, form: str, year_var):
        year_str = year_var.get().split("-")[0].strip()
        if not year_str.isdigit(): return
        d = self._json_dir(form, int(year_str))
        if os.path.exists(d):
            os.startfile(d)
        else:
            base = self.var_output.get().strip() or str(Path(__file__).parent)
            out_base = os.path.join(base, self._user(), form)
            if os.path.exists(out_base): os.startfile(out_base)

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
                year = self.var_1_year.get().split("-")[0]
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
        year = self.var_1_year.get().split("-")[0].strip()
        if not year.isdigit(): return
        
        periods = [p for p, v in self._1_period_vars.items() if v.get()]
        if not periods:
            from tkinter import messagebox
            messagebox.showwarning("Input Error", "Please select at least one period.")
            return

        self._log(f"Initiating GSTR-1 Download for FY {year}...", "step")
        self._start_gstr1_batch(int(year), periods)

    def _start_gstr1_batch(self, year, periods):
        self._yearly_stop = False; self._set_state("yearly")
        self._log(f"GSTR-1 Batch ({len(periods)} periods) — FY {year}", "step")
        import threading
        threading.Thread(target=self._worker_gstr1_batch, args=(year, periods, self.var_1_gstin.get().strip()), daemon=True).start()

    def _worker_gstr1_batch(self, year, periods, gstin):
        from concurrent.futures import ThreadPoolExecutor, as_completed
        out = self._json_dir("GSTR-1", year); user = self._user()
        ok = fail = skip = 0; saved_jsons = []
        
        self._log(f"Starting GSTR-1 Batch pipeline for {user} (FY {year})...", "info")

        # Phase 1: Pre-triggering
        self._log(f"Phase 1/2: Triggering offline ZIP generation for {len(periods)} periods...", "step")
        trigger_sess = make_session()
        trigger_sess.cookies.update(self._session.cookies)
        trigger_dl = Gstr1Downloader(session=trigger_sess, log_callback=None)
        
        def _trigger_task(p):
            ay = year + 1 if int(p) <= 3 else year
            try: trigger_dl.trigger_offline_gen(p, ay); return True
            except: return False
            
        with ThreadPoolExecutor(max_workers=6) as trigger_exec:
            trigger_exec.map(_trigger_task, periods)
        
        self._log("Trigger phase complete. Moving to download phase...", "success")

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

        self._log("Phase 2/2: Starting concurrent downloads...", "info")
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(_download_month_task, p, i) for i, p in enumerate(periods, 1)]
            for future in as_completed(futures):
                if self._yearly_stop: 
                    self._log("Batch download stop requested by user.", "warn")
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
                self.after(0, self.lbl_1_progress.configure, {"text": f"Progress: {ok+fail+skip}/{len(periods)}"})

        if not self._yearly_stop and self.var_excel_mode.get() == "consolidated": 
            self._batch_excel("GSTR-1", out, saved_jsons)
        
        self.after(0, self._finish_yearly, self.lbl_1_progress, ok, skip, fail, len(periods))

    # ── GSTR-2A Downloads ──────────────────────────────────────────────────────

    def _on_download_gstr2a(self):
        year = self.var_2a_year.get().split("-")[0].strip()
        if not year.isdigit(): return
        
        periods = [p for p, v in self._2a_period_vars.items() if v.get()]
        sections = [k for k, v in self._2a_section_vars.items() if v.get()]
        if not periods or not sections:
            from tkinter import messagebox
            messagebox.showwarning("Input Error", "Please select at least one period.")
            return

        if self.var_2a_mode.get() == "Quarterly":
            expanded = []
            for p in periods:
                if p == "6": expanded.extend(["4", "5", "6"])
                elif p == "9": expanded.extend(["7", "8", "9"])
                elif p == "12": expanded.extend(["10", "11", "12"])
                elif p == "3": expanded.extend(["1", "2", "3"])
            periods = expanded

        self._start_gstr2a_batch(int(year), periods, sections)

    def _start_gstr2a_batch(self, year, periods, sections):
        self._yearly_stop = False; self._set_state("yearly")
        import threading
        threading.Thread(target=self._worker_gstr2a_batch, args=(year, periods, sections), daemon=True).start()

    def _worker_gstr2a_batch(self, year, periods, sections):
        out = self._json_dir("GSTR-2A", year); user = self._user()
        ok = fail = skip = 0; saved_jsons = []
        for i, period in enumerate(periods, 1):
            if self._yearly_stop: break
            actual_year = year + 1 if int(period) <= 3 else year
            name = MONTH_NAMES.get(period, period)
            self.after(0, self.lbl_2a_progress.configure, {"text": f"[{i}/{len(periods)}] {name}"})
            try:
                d = Gstr2ADownloader(session=self._session)
                res = d.download_gstr2a(period, actual_year, sections, progress_callback=lambda m: self.after(0, self._log, f"  {m}", "info"))
                res.year = actual_year
                jp = save_gstr2a(res, out, user); saved_jsons.append(jp)
                ok += 1
            except Exception as e: self._log(f"Fail {name}: {e}", "error"); fail += 1
            self.after(0, self._set_det_progress, i)
        self._batch_excel("GSTR-2A", out, saved_jsons)
        self.after(0, self._finish_yearly, self.lbl_2a_progress, ok, skip, fail, len(periods))

    # ── GSTR-2B Downloads ──────────────────────────────────────────────────────

    def _on_download_gstr2b(self):
        mode = self.var_2b_mode.get(); year = self.var_2b_year.get().split("-")[0].strip()
        if not year.isdigit(): return
        
        periods = [p for p, v in self._2b_period_vars.items() if v.get()]
        if not periods:
            from tkinter import messagebox
            messagebox.showwarning("Input Error", "Please select at least one period.")
            return

        self._start_gstr2b_batch(int(year), mode, periods)

    def _start_gstr2b_batch(self, year, mode, periods):
        self._yearly_stop = False; self._set_state("yearly")
        import threading
        threading.Thread(target=self._worker_gstr2b_batch, args=(year, mode, periods), daemon=True).start()

    def _worker_gstr2b_batch(self, year, mode, periods):
        out = self._json_dir("GSTR-2B", year); user = self._user(); ok = fail = skip = 0; saved_jsons = []
        for i, period in enumerate(periods, 1):
            if self._yearly_stop: break
            actual_year = year + 1 if int(period) <= 3 else year
            name = (MONTH_NAMES if mode == "Monthly" else QUARTER_NAMES).get(period, period)
            self.after(0, self.lbl_2b_progress.configure, {"text": f"[{i}/{len(periods)}] {name}"})
            try:
                prd = str(period).zfill(2)
                fpath = os.path.join(out, f"GSTR2B_Return_{user}_{actual_year}_{prd}.json")
                if self.var_2b_skip.get() and os.path.exists(fpath):
                    self._log(f"  {name} skipped (exists)", "info")
                    saved_jsons.append(fpath); skip += 1
                else:
                    d = Gstr2BDownloader(session=self._session)
                    res = d.download_gstr2b(period, actual_year, "M" if mode == "Monthly" else "Q")
                    save_2b(res, out, user, period, actual_year)
                    saved_jsons.append(fpath)
                    ok += 1
            except Exception as e: self._log(f"Fail {name}: {e}", "error"); fail += 1
            self.after(0, self._set_det_progress, i)
        self._batch_excel("GSTR-2B", out, saved_jsons)
        self.after(0, self._finish_yearly, self.lbl_2b_progress, ok, skip, fail, len(periods))

    # ── GSTR-3B Downloads ──────────────────────────────────────────────────────

    def _on_download_gstr3b(self):
        year = self.var_3b_year.get().split("-")[0].strip()
        if not year.isdigit(): return
        
        periods = [p for p, v in self._3b_period_vars.items() if v.get()]
        if not periods:
            from tkinter import messagebox
            messagebox.showwarning("Input Error", "Please select at least one period.")
            return

        self._start_gstr3b_batch(int(year), periods)

    def _start_gstr3b_batch(self, year, periods):
        self._yearly_stop = False; self._set_state("yearly")
        import threading
        threading.Thread(target=self._worker_gstr3b_batch, args=(year, periods), daemon=True).start()

    def _worker_gstr3b_batch(self, year, periods):
        out = self._json_dir("GSTR-3B", year); user = self._user(); ok = fail = skip = 0; saved_jsons = []
        for i, period in enumerate(periods, 1):
            if self._yearly_stop: break
            actual_year = year + 1 if int(period) <= 3 else year
            name = (MONTH_NAMES if self.var_3b_mode.get() == "Monthly" else QUARTER_NAMES).get(period, period)
            self.after(0, self.lbl_3b_progress.configure, {"text": f"[{i}/{len(periods)}] {name}"})
            try:
                res = self._login_dl.download_gstr3b(period, actual_year)
                res.year = actual_year; jp = save_gstr3b(res, out, user); saved_jsons.append(jp); ok += 1
            except Exception as e: self._log(f"Fail {name}: {e}", "error"); fail += 1
            self.after(0, self._set_det_progress, i)
        self._batch_excel("GSTR-3B", out, saved_jsons)
        self.after(0, self._finish_yearly, self.lbl_3b_progress, ok, skip, fail, len(periods))

    # ── Log ──────────────────────────────────────────────────────────────────

    def _set_det_progress(self, current_val, max_val=12):
        pass

    def _log(self, msg: str, tag: str = "info"):
        # UI Log
        now_ui = datetime.now().strftime("%H:%M:%S")
        if hasattr(self, 'txt_log'):
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", f"[{now_ui}] {msg}\n", tag)
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")

        # File Log
        now_file = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{now_file}] [{tag.upper()}] {msg}\n"
        log_dir = os.path.join(os.path.dirname(__file__), "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"session_{datetime.now().strftime('%Y-%m-%d')}.log")
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(line)
        except: pass

if __name__ == "__main__":
    app = GstPortalApp()
    app.mainloop()
