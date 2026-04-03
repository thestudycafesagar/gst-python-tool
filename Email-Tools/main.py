import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import win32com.client
import pythoncom
import calendar
from datetime import date, datetime, timedelta

try:
    import openpyxl
    _OPENPYXL = True
except ImportError:
    _OPENPYXL = False

# ═══════════════════════════════════════════════════════════════
#  TEMPLATE DEFINITIONS 
# ═══════════════════════════════════════════════════════════════

class Template:
    def __init__(self, name, icon, config_fields, recipient_cols,
                 defaults, build_subject, build_body, has_attachment=False):
        self.name            = name
        self.icon            = icon
        self.config_fields   = config_fields    # [(label, key), ...]
        self.recipient_cols  = recipient_cols   # extra cols after Name, Email
        self.defaults        = defaults
        self.build_subject   = build_subject
        self.build_body      = build_body
        self.has_attachment  = has_attachment

    @property
    def all_cols(self):
        return ["Name", "Email"] + self.recipient_cols


# ── Template 1: GST Return ──────────────────────────────────────
def _gst_subject(cfg, row):
    return f'Share "{cfg["return_type"]}" Return Data for the Month of "{cfg["month"]}"'

def _gst_body(cfg, row):
    name = row.get("Name", "Recipient")
    return (
        f"Dear {name} Ji,\n\n"
        f"Please provide {cfg['return_type']} Data for the month of {cfg['month']} to File the Return.\n\n"
        f"Kindly send data upto {cfg['data_deadline']} so that we can also get time to verify the data.\n\n"
        f"Please Note that Last Date to file the Return is {cfg['filing_deadline']} so Request you to "
        f"send the Data Maximum By {cfg['data_deadline']}.\n\n"
        f"In Case of Any Query please feel free to call on +91-{cfg['phone']}.\n\n"
        f"Thanks & Regards\n{cfg['sender']}"
    )

TEMPLATE_GST = Template(
    name="GST Return Data Request",
    icon="📋",
    config_fields=[
        ("Month of Return:",          "month"),
        ("Return Type:",              "return_type"),
        ("Data Submission Deadline:", "data_deadline"),
        ("Last Date of Filing:",      "filing_deadline"),
        ("Sender Name:",              "sender"),
        ("Contact Phone:",            "phone"),
        ("CC Email Address:",         "cc"),
    ],
    recipient_cols=[],
    defaults={
        "month":           "March",
        "return_type":     "GST",
        "data_deadline":   "20 Mar 2026",
        "filing_deadline": "31 Mar 2026",
        "sender":          "Rohit Sharma",
        "phone":           "9876543210",
        "cc":              "cc@yourfirm.com",
    },
    build_subject=_gst_subject,
    build_body=_gst_body,
)


# ── Template 2: Invoice Sender ──────────────────────────────────
def _inv_subject(cfg, row):
    return f"Invoice for {row.get('Service','')} for the Period {row.get('Period','')}"

def _inv_body(cfg, row):
    name   = row.get("Name", "Recipient")
    svc    = row.get("Service", "")
    period = row.get("Period", "")
    amount = row.get("Invoice Amount", "")
    return (
        f"Dear {name},\n\n"
        f"I trust this message finds you well. We appreciate your continued partnership with us "
        f"and we are grateful for the opportunity to provide our services to you.\n\n"
        f"As per our agreement, we are pleased to send you the invoice for the {svc} availed for "
        f"the Period {period}. Please note that Invoice Amount is Rs. {amount}.\n\n"
        f"Please ensure that payment is made by the due date mentioned in the invoice. You can make "
        f"the payment through [Bank/Cheque/UPI etc.], and if you have any questions or require further "
        f"clarification regarding the invoice, please do not hesitate to reach out to our Email id.\n\n"
        f"Your satisfaction is our top priority, and we are committed to delivering exceptional service. "
        f"We look forward to your prompt payment and to continue serving your needs.\n\n"
        f"Thank you for choosing us as your service provider. We value your business and are here to "
        f"assist you with any requirements you may have.\n\n"
        f"Please find attached the invoice for your reference.\n\n"
        f"Should you have any queries or require further assistance, please feel free to contact us "
        f"at +91-{cfg['phone']}.\n\n"
        f"Sincerely,\n{cfg['org_name']}"
    )

TEMPLATE_INVOICE = Template(
    name="Invoice Sender",
    icon="🧾",
    config_fields=[
        ("CC Email Address:", "cc"),
        ("Organisation Name:", "org_name"),
        ("Contact Phone:",     "phone"),
    ],
    recipient_cols=["Invoice Amount", "Period", "Service", "Attachment File"],
    defaults={
        "cc":       "cc@yourfirm.com",
        "org_name": "Your Firm Name",
        "phone":    "9876543210",
    },
    build_subject=_inv_subject,
    build_body=_inv_body,
    has_attachment=True,
)


# ── Template 3: Outstanding Payment ────────────────────────────
def _pay_subject(cfg, row):
    name   = row.get("Name", "")
    period = row.get("Period", "")
    amount = row.get("Outstanding Amount", "")
    return f"{name}, Outstanding Payment Request for Services Utilized till {period} of Rs. {amount}."

def _pay_body(cfg, row):
    name    = row.get("Name", "Recipient")
    period  = row.get("Period", "")
    amount  = row.get("Outstanding Amount", "")
    service = row.get("Service Type", "")
    return (
        f"Dear {name},\n\n"
        f"I hope this email finds you well.\n\n"
        f"I am writing on behalf of {cfg['org_name']} to follow up on an outstanding payment against "
        f"the {service} services rendered to Your Organisation. As of {period}, there remains an unpaid "
        f"balance of Rs. {amount}.\n\n"
        f"We value our relationship with you and have always strived to offer our best services. "
        f"We would appreciate it if the pending amount could be cleared at the earliest to maintain "
        f"the smooth functioning of our business relationship.\n\n"
        f"Kindly verify this against your records, and if there are any discrepancies or clarifications "
        f"required, please let us know within the next {cfg['days_to_respond']}. If you have already "
        f"processed this payment, kindly share the payment details or the remittance advice, so we can "
        f"reconcile it in our system.\n\n"
        f"Please share the payment confirmation with us once completed.\n\n"
        f"We appreciate your prompt attention to this matter and look forward to continuing our collaboration.\n\n"
        f"Please Note that if Payment is not Received by {cfg['pay_deadline']} date then Interest of "
        f"{cfg['interest_rate']} will be levied on the Payment.\n\n"
        f"Therefore, Request you to clear the Payment Maximum By {cfg['pay_deadline']}.\n\n"
        f"In Case of Any Query please feel free to call on +91-{cfg['phone']}.\n\n"
        f"Thank you for your understanding and cooperation.\n\n"
        f"Regards\n{cfg['org_name']}"
    )

TEMPLATE_PAYMENT = Template(
    name="Outstanding Payment Reminder",
    icon="💰",
    config_fields=[
        ("CC Email Address:",    "cc"),
        ("Days to Respond:",     "days_to_respond"),
        ("Payment Deadline:",    "pay_deadline"),
        ("Interest Rate:",       "interest_rate"),
        ("Organisation Name:",   "org_name"),
        ("Contact Phone:",       "phone"),
    ],
    recipient_cols=["Period", "Outstanding Amount", "Service Type"],
    defaults={
        "cc":              "cc@yourfirm.com",
        "days_to_respond": "7 days",
        "pay_deadline":    "31 Mar 2026",
        "interest_rate":   "18% p.a.",
        "org_name":        "Your Firm Name",
        "phone":           "9876543210",
    },
    build_subject=_pay_subject,
    build_body=_pay_body,
)

ALL_TEMPLATES = [TEMPLATE_GST, TEMPLATE_INVOICE, TEMPLATE_PAYMENT]


# ═══════════════════════════════════════════════════════════════
#  SEND ENGINE
# ═══════════════════════════════════════════════════════════════

def send_emails(template, cfg, recipients, attachment_folder, log_cb, done_cb):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:
        log_cb(f"[ERROR] Could not connect to Outlook: {exc}")
        done_cb(success=False)
        pythoncom.CoUninitialize()
        return

    failed = []
    for idx, row in enumerate(recipients, 1):
        name  = str(row.get("Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        if not email:
            log_cb(f"[SKIP]  Row {idx} — empty email.")
            continue
        try:
            mail            = outlook.CreateItem(0)
            mail.To         = email
            mail.CC         = cfg.get("cc", "")
            mail.Subject    = template.build_subject(cfg, row)
            body_text       = template.build_body(cfg, row)
            # Use HTMLBody so Outlook doesn't override with its default HTML/signature
            html_body       = "<html><body><pre style='font-family:Calibri,sans-serif;font-size:11pt;white-space:pre-wrap'>" \
                              + body_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") \
                              + "</pre></body></html>"
            mail.HTMLBody   = html_body

            if template.has_attachment and attachment_folder:
                fname = str(row.get("Attachment File", "")).strip()
                if fname:
                    import os
                    base = fname[:-4] if fname.lower().endswith(".pdf") else fname
                    path = os.path.join(attachment_folder, base + ".pdf")
                    if os.path.exists(path):
                        mail.Attachments.Add(path)
                    else:
                        log_cb(f"[WARN]  {name} — attachment not found: {base}.pdf")

            mail.Send()
            log_cb(f"[SENT]  {name} <{email}>")
        except Exception as exc:
            log_cb(f"[ERROR] {name} <{email}> — {exc}")
            failed.append(email)

    pythoncom.CoUninitialize()
    done_cb(success=True, failed=failed)


# ═══════════════════════════════════════════════════════════════
#  GUI  —  Design Tokens
# ═══════════════════════════════════════════════════════════════

# Palette
ACCENT      = "#2563EB"   # Blue-600
ACCENT_H    = "#1D4ED8"   # Blue-700 (hover)
ACCENT_L    = "#DBEAFE"   # Blue-100 (selection / tint)
BG          = "#F8FAFC"   # Slate-50  (page background)
BG2         = "#F1F5F9"   # Slate-100 (panel / sidebar)
BORDER      = "#E2E8F0"   # Slate-200 (subtle border)
BORDER2     = "#CBD5E1"   # Slate-300 (strong border)
CARD        = "#FFFFFF"   # Pure white surface
TEXT        = "#0F172A"   # Slate-900
TEXT2       = "#475569"   # Slate-600
MUTED       = "#94A3B8"   # Slate-400
HDR_BG      = "#1E3A5F"   # Navy header (light-mode default)
HDR_TEXT    = "#F1F5F9"   # Slate-100
HDR_MUTED   = "#94A3B8"   # Slate-400
COL_HDR_BG  = "#2D4A6E"   # Navy-medium (data grid column headers)
COL_HDR_FG  = "#CBD5E1"   # Slate-300
GREEN       = "#10B981"
RED         = "#EF4444"
YELLOW      = "#F59E0B"
LOG_BG      = "#0D1117"   # GitHub dark
LOG_FG      = "#8B949E"   # GitHub dark muted

# ── Theme palettes — bg/fg kept separate to avoid colour-key collisions ───────
_EMAIL_THEMES = {
    "Light": dict(
        bg="#F8FAFC",  bg2="#F1F5F9",  card="#FFFFFF",
        border="#E2E8F0", border2="#CBD5E1", accent_l="#DBEAFE",
        text="#0F172A", text2="#475569", muted="#94A3B8",
        hdr_bg="#1E3A5F", hdr_text="#F1F5F9", hdr_muted="#94A3B8",
        col_hdr_bg="#2D4A6E", col_hdr_fg="#CBD5E1",
    ),
    "Dark": dict(
        bg="#1e293b",  bg2="#162032",  card="#1e2d3d",
        border="#334155", border2="#475569", accent_l="#1e3a6e",
        text="#f1f5f9", text2="#94a3b8", muted="#64748b",
        hdr_bg="#0F172A", hdr_text="#F1F5F9", hdr_muted="#64748B",
        col_hdr_bg="#1E293B", col_hdr_fg="#94A3B8",
    ),
}

# Fonts
F_TITLE    = ("Segoe UI", 16, "bold")
F_BADGE    = ("Segoe UI", 12, "bold")
F_LABEL    = ("Segoe UI", 13)
F_LABEL_SML= ("Segoe UI", 11)
F_INPUT    = ("Segoe UI", 13)
F_BTN      = ("Segoe UI", 11, "bold")
F_BTN_LG   = ("Segoe UI", 12, "bold")
F_MONO     = ("Consolas", 12)
F_MONO_LG  = ("Consolas", 13)
F_TAB      = ("Segoe UI", 11, "bold")
F_COL_HDR  = ("Segoe UI", 11, "bold")


# ═══════════════════════════════════════════════════════════════
#  GUI  —  Date Picker
# ═══════════════════════════════════════════════════════════════

class CustomDatePicker(tk.Toplevel):
    def __init__(self, parent, target_var):
        super().__init__(parent)
        self.target_var = target_var
        self.title("Select Date")
        self.geometry("350x420")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.current_date = datetime.now()
        
        self.configure(bg=CARD)
        
        self._build_ui()
        self._update_calendar()
        
    def _build_ui(self):
        header = tk.Frame(self, bg=ACCENT, pady=10)
        header.pack(fill="x")
        
        ttk.Button(header, text="<", width=3, command=self._prev_month).pack(side="left", padx=10)
        self.month_lbl = tk.Label(header, text="", font=F_BTN_LG, bg=ACCENT, fg="white")
        self.month_lbl.pack(side="left", expand=True)
        ttk.Button(header, text=">", width=3, command=self._next_month).pack(side="right", padx=10)
        
        self.cal_frame = tk.Frame(self, bg=CARD, padx=10, pady=10)
        self.cal_frame.pack(fill="both", expand=True)
        
        days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for c, day in enumerate(days):
            tk.Label(self.cal_frame, text=day, font=F_LABEL_SML, bg=CARD, fg=MUTED).grid(row=0, column=c, padx=3, pady=3)
            
        self.day_buttons = []
        for r in range(6):
            row_btns = []
            for c in range(7):
                btn = tk.Button(self.cal_frame, text="", width=3, relief="flat", bg=CARD, fg=TEXT, font=F_LABEL,
                                command=lambda r=r, c=c: self._select_day(r, c))
                btn.grid(row=r+1, column=c, padx=2, pady=2)
                row_btns.append(btn)
            self.day_buttons.append(row_btns)
            
    def _update_calendar(self):
        self.month_lbl.config(text=self.current_date.strftime("%B %Y"))
        
        cal = calendar.monthcalendar(self.current_date.year, self.current_date.month)
        
        for r in range(6):
            for c in range(7):
                if r < len(cal) and cal[r][c] != 0:
                    day = cal[r][c]
                    self.day_buttons[r][c].config(text=str(day), state="normal")
                else:
                    self.day_buttons[r][c].config(text="", state="disabled")
                    
    def _prev_month(self):
        first = self.current_date.replace(day=1)
        prev_month = first - timedelta(days=1)
        self.current_date = prev_month.replace(day=1)
        self._update_calendar()
        
    def _next_month(self):
        days_in_month = calendar.monthrange(self.current_date.year, self.current_date.month)[1]
        last = self.current_date.replace(day=days_in_month)
        next_month = last + timedelta(days=1)
        self.current_date = next_month.replace(day=1)
        self._update_calendar()
        
    def _select_day(self, r, c):
        day = int(self.day_buttons[r][c].cget("text"))
        selected = datetime(self.current_date.year, self.current_date.month, day)
        self.target_var.set(selected.strftime("%d %b %Y"))
        self.destroy()

# ═══════════════════════════════════════════════════════════════
#  GUI  —  Application
# ═══════════════════════════════════════════════════════════════

class BulkMailApp(tk.Tk):
    def __init__(self, hide_switcher=False):
        super().__init__()
        self.title("Bulk Outlook Mailer")
        self.geometry("860x680")
        self.minsize(740, 560)
        self.configure(bg=BG)
        self._hide_switcher = hide_switcher

        self._active_template  = ALL_TEMPLATES[0]
        self._cfg_vars         = {}
        self._attachment_folder = ""
        self._sending          = False
        self._add_vars         = {}

        self._apply_styles()
        self._build_header()
        self._build_bottom_bar()   # must pack before notebook so it isn't pushed out
        self._build_notebook()     # fills remaining space after bottom bar is reserved

        self._switch_template(0)

    # ── Styles ──────────────────────────────────────────────────
    def _apply_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")

        # Notebook
        s.configure("TNotebook",
                    background=BG2,
                    borderwidth=0,
                    tabmargins=[0, 0, 0, 0])
        s.configure("TNotebook.Tab",
                    background=BG2,
                    foreground=MUTED,
                    padding=[20, 10],
                    font=F_TAB,
                    borderwidth=0,
                    focuscolor=BG2)
        s.map("TNotebook.Tab",
              background=[("selected", CARD), ("active", BG)],
              foreground=[("selected", ACCENT), ("active", TEXT2)],
              focuscolor=[("selected", CARD)])

        # Frames
        s.configure("TFrame",      background=BG)
        s.configure("Card.TFrame", background=CARD)

        # Labels
        s.configure("TLabel", background=BG, foreground=TEXT, font=F_LABEL)

        # Entry
        s.configure("TEntry",
                    fieldbackground=CARD,
                    foreground=TEXT,
                    font=F_INPUT,
                    padding=[12, 9],
                    relief="flat",
                    borderwidth=1,
                    bordercolor=BORDER2)
        s.map("TEntry",
              bordercolor=[("focus", ACCENT)],
              lightcolor=[("focus", ACCENT)],
              darkcolor=[("focus", ACCENT)])

        # Default button
        s.configure("TButton",
                    background=CARD,
                    foreground=TEXT2,
                    font=F_BTN,
                    padding=[18, 11],
                    relief="flat",
                    borderwidth=1,
                    bordercolor=BORDER2,
                    focuscolor=CARD)
        s.map("TButton",
              background=[("active", BG2), ("pressed", BG2)],
              bordercolor=[("active", BORDER2)])

        # Primary send button
        s.configure("Send.TButton",
                    background=ACCENT,
                    foreground="white",
                    font=F_BTN_LG,
                    padding=[28, 14],
                    relief="flat",
                    borderwidth=0,
                    focuscolor=ACCENT)
        s.map("Send.TButton",
              background=[("active", ACCENT_H), ("pressed", ACCENT_H)],
              foreground=[("active", "white")])

        # Import button (accent tint)
        s.configure("Import.TButton",
                    background=ACCENT_L,
                    foreground=ACCENT,
                    font=F_BTN,
                    padding=[12, 7],
                    relief="flat",
                    borderwidth=0,
                    focuscolor=ACCENT_L)
        s.map("Import.TButton",
              background=[("active", "#BFDBFE"), ("pressed", "#BFDBFE")])

        # Scrollbars
        s.configure("TScrollbar",
                    background=BG2,
                    troughcolor=BG,
                    borderwidth=0,
                    arrowsize=12,
                    relief="flat")
        s.map("TScrollbar",
              background=[("active", BORDER2), ("pressed", MUTED)])

        # Treeview
        s.configure("Treeview",
                    background=CARD,
                    fieldbackground=CARD,
                    foreground=TEXT,
                    rowheight=32,
                    font=F_LABEL,
                    borderwidth=0)
        s.configure("Treeview.Heading",
                    background=BG2,
                    foreground=TEXT2,
                    font=("Segoe UI", 11, "bold"),
                    padding=[10, 8],
                    relief="flat")
        s.map("Treeview",
              background=[("selected", ACCENT_L)],
              foreground=[("selected", ACCENT)])

        # Body combobox (used in config tab)
        s.configure("TCombobox",
                    fieldbackground=CARD,
                    background=CARD,
                    foreground=TEXT,
                    arrowcolor=TEXT2,
                    font=F_INPUT,
                    padding=[8, 8])
        s.map("TCombobox",
              fieldbackground=[("readonly", CARD)],
              selectbackground=[("readonly", CARD)],
              selectforeground=[("readonly", TEXT)])

        # Header combobox — dark background, white text, no border flash
        s.configure("Header.TCombobox",
                    fieldbackground=HDR_BG,
                    background=HDR_BG,
                    foreground=HDR_TEXT,
                    selectbackground=HDR_BG,
                    selectforeground=HDR_TEXT,
                    arrowcolor=HDR_TEXT,
                    font=F_LABEL,
                    padding=[10, 8],
                    relief="flat",
                    borderwidth=0)
        s.map("Header.TCombobox",
              fieldbackground=[("readonly", HDR_BG)],
              background=[("readonly", HDR_BG)],
              selectbackground=[("readonly", HDR_BG)],
              foreground=[("readonly", HDR_TEXT)],
              selectforeground=[("readonly", HDR_TEXT)],
              arrowcolor=[("readonly", HDR_TEXT)])
        self.option_add("*Header.TCombobox*Listbox.background", CARD)
        self.option_add("*Header.TCombobox*Listbox.foreground", TEXT)
        self.option_add("*Header.TCombobox*Listbox.font", F_INPUT)
        self.option_add("*Header.TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*Header.TCombobox*Listbox.selectForeground", "#ffffff")

        # Separator
        s.configure("TSeparator", background=BORDER)

    # ── Called by GST Suite on Dark / Light toggle ──────────────
    def set_theme(self, mode: str):
        global BG, BG2, CARD, BORDER, BORDER2, ACCENT_L, TEXT, TEXT2, MUTED
        global HDR_BG, HDR_TEXT, HDR_MUTED, COL_HDR_BG, COL_HDR_FG

        t   = _EMAIL_THEMES.get(mode, _EMAIL_THEMES["Light"])
        old = _EMAIL_THEMES["Dark" if mode == "Light" else "Light"]

        # 1) Update module-level palette so future widget builds use new colours
        BG, BG2, CARD      = t["bg"],     t["bg2"],    t["card"]
        BORDER, BORDER2    = t["border"], t["border2"]
        ACCENT_L           = t["accent_l"]
        TEXT, TEXT2, MUTED = t["text"],   t["text2"],  t["muted"]
        HDR_BG, HDR_TEXT, HDR_MUTED   = t["hdr_bg"],  t["hdr_text"],  t["hdr_muted"]
        COL_HDR_BG, COL_HDR_FG        = t["col_hdr_bg"], t["col_hdr_fg"]

        # 2) Re-apply all ttk styles
        self._apply_styles()

        # 3) Rebuild dynamic tabs so they pick up new globals
        self._rebuild_config_tab()

        # 4) Separate swap maps — bg colours and fg colours diverge in light mode
        #    (e.g. TEXT=#0F172A == HDR_BG, so they must not be swapped together)
        _BG_KEYS = ("bg", "bg2", "card", "border", "border2", "accent_l",
                    "hdr_bg", "col_hdr_bg")
        _FG_KEYS = ("text", "text2", "muted", "hdr_text", "hdr_muted", "col_hdr_fg")

        bg_swap = {old[k].lower(): t[k] for k in _BG_KEYS
                   if old[k].lower() != t[k].lower()}
        fg_swap = {old[k].lower(): t[k] for k in _FG_KEYS
                   if old[k].lower() != t[k].lower()}

        def _recolor(w):
            for attr in ("bg", "background", "highlightbackground", "activebackground"):
                try:
                    val = str(w.cget(attr)).lower()
                    if val in bg_swap:
                        w.configure(**{attr: bg_swap[val]})
                except Exception:
                    pass
            for attr in ("fg", "foreground", "activeforeground"):
                try:
                    val = str(w.cget(attr)).lower()
                    if val in fg_swap:
                        w.configure(**{attr: fg_swap[val]})
                except Exception:
                    pass
            for child in w.winfo_children():
                _recolor(child)

        _recolor(self)

    # ── Header ──────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=HDR_BG)
        hdr.pack(fill="x")

        inner = tk.Frame(hdr, bg=HDR_BG, padx=18, pady=13)
        inner.pack(fill="x")

        # Left: dynamic template title (updates when template changes)
        left = tk.Frame(inner, bg=HDR_BG)
        left.pack(side="left")
        self._hdr_title = tk.Label(left, text="✉  Bulk Outlook Mailer",
                                   bg=HDR_BG, fg=HDR_TEXT,
                                   font=F_TITLE)
        self._hdr_title.pack(side="left")

        self._tmpl_var = tk.StringVar()
        self._tmpl_cb  = None

        if not self._hide_switcher:
            # Right: template switcher (compact — title already shows active template)
            right = tk.Frame(inner, bg=HDR_BG)
            right.pack(side="right")
            tk.Label(right, text="SWITCH:",
                     bg=HDR_BG, fg=HDR_MUTED,
                     font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 8))

            names = [f"{t.icon}  {t.name}" for t in ALL_TEMPLATES]
            cb = ttk.Combobox(right, textvariable=self._tmpl_var, values=names,
                              state="readonly", width=30,
                              style="Header.TCombobox")
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: self._on_template_change())
            self._tmpl_cb = cb

        # Accent underline
        tk.Frame(hdr, bg=ACCENT, height=2).pack(fill="x")

    # ── Notebook ────────────────────────────────────────────────
    def _build_notebook(self):
        self._nb = ttk.Notebook(self)
        self._nb.pack(fill="both", expand=True, padx=12, pady=(10, 0))

        self._tab_config     = ttk.Frame(self._nb, padding=0)
        self._tab_recipients = ttk.Frame(self._nb, padding=0)
        self._tab_log        = ttk.Frame(self._nb, padding=0)

        self._nb.add(self._tab_config,     text="  Configuration  ")
        self._nb.add(self._tab_recipients, text="  Recipients  ")
        self._nb.add(self._tab_log,        text="  Log  ")

        self._build_log_tab()

    # ── Bottom bar ──────────────────────────────────────────────
    def _build_bottom_bar(self):
        bar = tk.Frame(self, bg=CARD, pady=10)
        bar.pack(fill="x", side="bottom")

        # Top border line
        tk.Frame(bar, bg=BORDER, height=1).pack(fill="x", side="top")

        inner = tk.Frame(bar, bg=CARD)
        inner.pack(fill="x", padx=14, pady=(8, 0))

        # Status label (hidden — template is already shown in header combobox)
        self._status_lbl = tk.Label(inner, text="",
                                    bg=CARD, fg=MUTED, font=F_LABEL_SML)

        # Buttons on the right
        ttk.Button(inner, text="  Send All Emails  →",
                   style="Send.TButton",
                   command=self._start_send).pack(side="right", padx=(6, 0))
        ttk.Button(inner, text="Preview Email",
                   command=self._show_preview).pack(side="right", padx=(4, 0))
        ttk.Button(inner, text="Check Data",
                   command=self._check_data).pack(side="right", padx=(4, 0))

    # ── Config tab ──────────────────────────────────────────────
    def _rebuild_config_tab(self):
        for w in self._tab_config.winfo_children():
            w.destroy()
        self._cfg_vars = {}
        t = self._active_template

        outer = tk.Frame(self._tab_config, bg=BG)
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 12))

        # Form card
        form_border = tk.Frame(outer, bg=BORDER, padx=1, pady=1)
        form_border.pack(fill="x")
        form_card = tk.Frame(form_border, bg=CARD, padx=32, pady=24)
        form_card.pack(fill="both", expand=True)
        form_card.columnconfigure(1, weight=1)

        def validate_phone(P):
            if P == "": return True
            if P.isdigit() and len(P) <= 10: return True
            return False
        vcmd_phone = (self.register(validate_phone), "%P")

        for i, (label, key) in enumerate(t.config_fields):
            tk.Label(form_card, text=label, anchor="w",
                     bg=CARD, fg=TEXT2,
                     font=F_LABEL).grid(row=i, column=0, sticky="w",
                                        pady=12, padx=(0, 28))
            var = tk.StringVar(value=t.defaults.get(key, ""))

            # Entry with focus-ring border frame
            ef = tk.Frame(form_card, bg=BORDER2, padx=1, pady=1)
            ef.grid(row=i, column=1, sticky="ew", pady=12)

            if key == "phone":
                inner_f = tk.Frame(ef, bg=CARD)
                inner_f.pack(fill="x")
                tk.Label(inner_f, text="+91", bg=CARD, fg=TEXT2, font=F_INPUT).pack(side="left", padx=(12, 0))
                e = tk.Entry(inner_f, textvariable=var,
                             font=F_INPUT, bg=CARD, fg=TEXT,
                             relief="flat", bd=0,
                             insertbackground=ACCENT,
                             validate="key", validatecommand=vcmd_phone)
                e.pack(side="left", fill="x", expand=True, padx=(4, 12), pady=11)
            elif "deadline" in key.lower() or "date" in key.lower():
                inner_f = tk.Frame(ef, bg=CARD)
                inner_f.pack(fill="x")
                e = tk.Entry(inner_f, textvariable=var,
                             font=F_INPUT, bg=CARD, fg=TEXT,
                             relief="flat", bd=0,
                             insertbackground=ACCENT)
                e.pack(side="left", fill="x", expand=True, padx=(12, 4), pady=11)
                btn = tk.Label(inner_f, text="📅", bg=CARD, fg=TEXT2, font=F_INPUT, cursor="hand2")
                btn.pack(side="right", padx=(0, 12))
                btn.bind("<Button-1>", lambda event, v=var: CustomDatePicker(self, v))
            else:
                e = tk.Entry(ef, textvariable=var,
                             font=F_INPUT, bg=CARD, fg=TEXT,
                             relief="flat", bd=0,
                             insertbackground=ACCENT)
                e.pack(fill="x", padx=12, pady=11)
                
            e.bind("<FocusIn>",  lambda ev, f=ef: f.config(bg=ACCENT))
            e.bind("<FocusOut>", lambda ev, f=ef: f.config(bg=BORDER2))
            
            self._cfg_vars[key] = var

    # ── Recipients tab ──────────────────────────────────────────
    def _rebuild_recipients_tab(self):
        for w in self._tab_recipients.winfo_children():
            w.destroy()
        t = self._active_template
        cols = t.all_cols

        # Top toolbar
        toolbar = tk.Frame(self._tab_recipients, bg=BG, padx=12, pady=8)
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="📂  Import from Excel",
                   style="Import.TButton",
                   command=self._import_excel).pack(side="left")

        ttk.Button(toolbar, text="🗑  Clear All",
                   command=self._clear_recipients).pack(side="left", padx=(8, 0))

        # Attachment folder (Invoice only)
        if t.has_attachment:
            tk.Frame(toolbar, bg=BORDER2, width=1).pack(
                side="left", fill="y", padx=(12, 8))
            tk.Label(toolbar, text="Attachment Folder:",
                     fg=TEXT2, font=F_LABEL, bg=BG).pack(side="left")
            self._folder_lbl = tk.Label(toolbar,
                                        text="  (not selected)",
                                        bg=BG2, fg=MUTED,
                                        font=F_LABEL,
                                        padx=8, pady=5,
                                        width=32, anchor="w",
                                        relief="flat",
                                        highlightbackground=BORDER2,
                                        highlightthickness=1)
            self._folder_lbl.pack(side="left", padx=6)
            ttk.Button(toolbar, text="Browse…",
                       command=self._pick_folder).pack(side="left")

        # ── Multi-column editor ───────────────────────────────
        # Bordered container
        editor_border = tk.Frame(self._tab_recipients, bg=BORDER2, padx=1, pady=1)
        editor_border.pack(fill="both", expand=True, padx=12, pady=(0, 2))

        editor_outer = tk.Frame(editor_border, bg=BG)
        editor_outer.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(editor_outer, orient="vertical")
        vsb.pack(side="right", fill="y")

        h_canvas = tk.Canvas(editor_outer, bg=BG, highlightthickness=0)
        hsb = ttk.Scrollbar(editor_outer, orient="horizontal",
                             command=h_canvas.xview)
        h_canvas.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        h_canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(h_canvas, bg=BORDER)
        _canvas_win = h_canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_resize(e):
            h_canvas.configure(scrollregion=h_canvas.bbox("all"))
        def _on_canvas_resize(e):
            h_canvas.itemconfig(_canvas_win, height=e.height)
        inner.bind("<Configure>", _on_inner_resize)
        h_canvas.bind("<Configure>", _on_canvas_resize)

        self._col_texts = []
        self._lineno_widgets = []

        def _make_yview(*args):
            for tw in self._col_texts:
                tw.yview(*args)

        def _on_scroll(*args):
            vsb.set(*args)
            for tw in self._col_texts:
                tw.yview_moveto(args[0])
            _refresh_linenos()

        vsb.config(command=_make_yview)

        def _refresh_linenos():
            for tw, ln_w in zip(self._col_texts, self._lineno_widgets):
                ln_w.config(state="normal")
                ln_w.delete("1.0", "end")
                total = int(tw.index("end-1c").split(".")[0])
                nums  = "\n".join(str(i) for i in range(1, total + 1))
                ln_w.insert("1.0", nums)
                ln_w.config(state="disabled")
                ln_w.yview_moveto(tw.yview()[0])

        def _on_any_change():
            _refresh_linenos()
            self._update_count()
            self._check_line_mismatch()

        COL_MIN_W = 220
        for col_name in cols:
            col_frame = tk.Frame(inner, bg=BG, width=COL_MIN_W)
            col_frame.pack(side="left", fill="y", padx=(0, 1))
            col_frame.pack_propagate(False)

            # Dark column header (data-grid style)
            tk.Label(col_frame, text=col_name.upper(),
                     bg=COL_HDR_BG, fg=COL_HDR_FG,
                     font=F_COL_HDR,
                     anchor="w", padx=12, pady=10).pack(fill="x")

            row = tk.Frame(col_frame, bg=CARD)
            row.pack(fill="both", expand=True)

            # Line-number gutter
            ln_w = tk.Text(row, width=3, font=("Consolas", 11),
                           bg="#F1F5F9", fg="#CBD5E1",
                           relief="flat", state="disabled",
                           cursor="arrow", takefocus=False,
                           padx=2,
                           yscrollcommand=lambda *a: None)
            ln_w.pack(side="left", fill="y")
            tk.Frame(row, bg=BORDER, width=1).pack(side="left", fill="y")
            self._lineno_widgets.append(ln_w)

            tw = tk.Text(row, font=F_MONO_LG,
                         bg=CARD, fg=TEXT,
                         insertbackground=ACCENT,
                         selectbackground=ACCENT_L,
                         selectforeground=TEXT,
                         relief="flat", wrap="none", undo=True,
                         padx=6, pady=2,
                         yscrollcommand=_on_scroll)
            tw.pack(side="left", fill="both", expand=True)
            tw.bind("<KeyRelease>", lambda e: _on_any_change())
            tw.bind("<ButtonRelease>", lambda e: _refresh_linenos())
            self._col_texts.append(tw)

        # Intercept Ctrl+V → distribute tab-separated clipboard
        def _on_paste(event):
            try:
                raw = self.clipboard_get()
            except tk.TclError:
                return "break"
            lines = [ln for ln in raw.splitlines() if ln.strip()]
            if not lines:
                return "break"
            has_tabs = any("\t" in ln for ln in lines)
            if has_tabs:
                for line in lines:
                    parts = line.split("\t")
                    parts = (parts + [""] * len(cols))[:len(cols)]
                    for ci, tw in enumerate(self._col_texts):
                        existing = tw.get("1.0", "end").rstrip("\n")
                        if existing:
                            tw.insert("end", "\n")
                        tw.insert("end", parts[ci])
            else:
                focused = event.widget
                for line in lines:
                    existing = focused.get("1.0", "end").rstrip("\n")
                    if existing:
                        focused.insert("end", "\n")
                    focused.insert("end", line)
            _on_any_change()
            return "break"

        for tw in self._col_texts:
            tw.bind("<Control-v>", _on_paste)
            tw.bind("<Control-V>", _on_paste)

        # Bottom status bar
        bot = tk.Frame(self._tab_recipients, bg=BG2, pady=5)
        bot.pack(fill="x")

        self._count_lbl = tk.Label(bot, text="0 recipients",
                                   bg=BG2, fg=TEXT2,
                                   font=("Segoe UI", 11, "bold"))
        self._count_lbl.pack(side="left", padx=12)

        self._mismatch_lbl = tk.Label(bot, text="",
                                      bg=BG2, fg=RED,
                                      font=("Segoe UI", 8, "bold"))
        self._mismatch_lbl.pack(side="left", padx=8)

        tk.Label(bot, text="Ctrl+V  ·  paste from Excel",
                 bg=BG2, fg=MUTED, font=("Segoe UI", 7)).pack(side="right", padx=12)

        _refresh_linenos()
        self._rec_text = self._col_texts[0]
        self._update_count()

    # ── Log tab ─────────────────────────────────────────────────
    def _build_log_tab(self):
        # Header row
        hdr = tk.Frame(self._tab_log, bg=CARD,
                       highlightbackground=BORDER,
                       highlightthickness=1)
        hdr.pack(fill="x", padx=12, pady=(12, 0))

        tk.Label(hdr, text="  Send Log", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 9, "bold"),
                 pady=9).pack(side="left")
        ttk.Button(hdr, text="Clear",
                   command=self._clear_log).pack(side="right", padx=6, pady=4)

        # Log area in bordered container
        log_border = tk.Frame(self._tab_log, bg=BORDER2, padx=1, pady=1)
        log_border.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self._log_box = scrolledtext.ScrolledText(
            log_border,
            font=F_MONO,
            state="disabled",
            bg=LOG_BG,
            fg=LOG_FG,
            insertbackground="white",
            relief="flat",
            padx=14, pady=12,
            borderwidth=0)
        self._log_box.pack(fill="both", expand=True)

        self._log_box.tag_configure("sent",  foreground="#4ADE80")
        self._log_box.tag_configure("error", foreground="#F87171")
        self._log_box.tag_configure("skip",  foreground="#FBBF24")
        self._log_box.tag_configure("warn",  foreground="#FB923C")
        self._log_box.tag_configure("info",  foreground="#60A5FA")

    # ── Template switching ───────────────────────────────────────
    def _on_template_change(self):
        if self._tmpl_cb is None:
            return
        idx = self._tmpl_cb.current()
        if idx >= 0:
            self._switch_template(idx)

    def _switch_template(self, idx):
        self._active_template = ALL_TEMPLATES[idx]
        self._tmpl_var.set(f"{self._active_template.icon}  {self._active_template.name}")
        self._hdr_title.config(
            text=f"{self._active_template.icon}  {self._active_template.name}")
        self._attachment_folder = ""
        self._rebuild_config_tab()
        self._rebuild_recipients_tab()
        self._status_lbl.config(text="")

    # ── Helpers ─────────────────────────────────────────────────
    def _get_config(self):
        return {k: v.get() for k, v in self._cfg_vars.items()}

    def _get_recipients(self):
        cols     = self._active_template.all_cols
        col_lines = [tw.get("1.0", "end").splitlines()
                     for tw in self._col_texts]
        max_len  = max((len(cl) for cl in col_lines), default=0)
        col_lines = [cl + [""] * (max_len - len(cl)) for cl in col_lines]
        rows = []
        for values in zip(*col_lines):
            if not any(v.strip() for v in values):
                continue
            rows.append(dict(zip(cols, values)))
        return rows

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="Select Attachment Folder")
        if folder:
            self._attachment_folder = folder
            short = folder if len(folder) < 45 else "…" + folder[-42:]
            self._folder_lbl.config(text=f"  {short}", fg=TEXT)

    def _clear_recipients(self):
        if not hasattr(self, "_col_texts") or not self._col_texts:
            return
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all recipients?", parent=self):
            for tw in self._col_texts:
                tw.delete("1.0", "end")
            if hasattr(self, "_update_count"):
                self._update_count()
            if hasattr(self, "_check_line_mismatch"):
                self._check_line_mismatch()
            for tw in self._col_texts:
                tw.event_generate("<KeyRelease>")

    def _import_excel(self):
        if not _OPENPYXL:
            messagebox.showerror("Missing Library",
                "openpyxl is not installed.\n\nRun: pip install openpyxl")
            return
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
        except Exception as exc:
            messagebox.showerror("Read Error", f"Could not read Excel file:\n{exc}")
            return
        if not rows:
            messagebox.showwarning("Empty File", "The Excel file has no data.")
            return

        cols = self._active_template.all_cols
        cols_lower = [c.lower() for c in cols]

        first_row = [str(v).strip() if v is not None else "" for v in rows[0]]
        header_map = {}
        for xi, hdr in enumerate(first_row):
            hl = hdr.lower()
            if hl in cols_lower:
                header_map[xi] = cols_lower.index(hl)

        if header_map:
            data_rows = rows[1:]
        else:
            data_rows = rows
            header_map = {xi: xi for xi in range(min(len(first_row), len(cols)))}

        for tw in self._col_texts:
            tw.delete("1.0", "end")

        for row in data_rows:
            if not any(v for v in row if v is not None and str(v).strip()):
                continue
            for xi, ci in header_map.items():
                val = row[xi] if xi < len(row) else ""
                val = "" if val is None else str(val).strip()
                existing = self._col_texts[ci].get("1.0", "end").rstrip("\n")
                if existing:
                    self._col_texts[ci].insert("end", "\n")
                self._col_texts[ci].insert("end", val)

        if hasattr(self, "_rec_text"):
            self._update_count()
        self._check_line_mismatch()
        for tw in self._col_texts:
            tw.event_generate("<KeyRelease>")

    def _check_line_mismatch(self):
        if not hasattr(self, "_col_texts") or not hasattr(self, "_mismatch_lbl"):
            return
        cols = self._active_template.all_cols
        counts = []
        for tw in self._col_texts:
            n = sum(1 for ln in tw.get("1.0", "end").splitlines() if ln.strip())
            counts.append(n)
        non_zero = [n for n in counts if n > 0]
        if len(set(non_zero)) > 1:
            detail = "  ".join(
                f"{c}={n}" for c, n in zip(cols, counts) if n > 0)
            self._mismatch_lbl.config(
                text=f"⚠  Line mismatch: {detail} — please fix before sending")
        else:
            self._mismatch_lbl.config(text="")

    def _update_count(self):
        if not hasattr(self, "_col_texts") or not hasattr(self, "_count_lbl"):
            return
        raw = self._col_texts[0].get("1.0", "end")
        n = sum(1 for line in raw.splitlines() if line.strip())
        self._count_lbl.config(text=f"{n} recipient{'s' if n != 1 else ''}")

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _clear_recipients(self):
        for tw in self._col_texts:
            tw.delete("1.0", "end")
        self._update_count()

    # ── Preview popup ───────────────────────────────────────────
    def _show_preview(self):
        cfg  = self._get_config()
        recs = self._get_recipients()
        t    = self._active_template

        sample_row = recs[0] if recs else {c: f"<{c}>" for c in t.all_cols}
        subject    = t.build_subject(cfg, sample_row)
        body       = t.build_body(cfg, sample_row)

        win = tk.Toplevel(self)
        win.title("Email Preview")
        win.geometry("660x560")
        win.configure(bg=BG)
        win.grab_set()

        # Header
        hdr = tk.Frame(win, bg=HDR_BG, padx=18, pady=13)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"{t.icon}  {t.name}",
                 bg=HDR_BG, fg=HDR_TEXT, font=F_BADGE).pack(side="left")
        if recs:
            tk.Label(hdr, text=f"Preview: {sample_row.get('Name', '—')}",
                     bg=HDR_BG, fg=HDR_MUTED, font=F_LABEL_SML).pack(side="right")
        tk.Frame(win, bg=ACCENT, height=2).pack(fill="x")

        # Content area
        content = tk.Frame(win, bg=BG, padx=16, pady=14)
        content.pack(fill="both", expand=True)

        # Subject
        tk.Label(content, text="SUBJECT", bg=BG, fg=MUTED,
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(0, 3))
        subj_border = tk.Frame(content, bg=BORDER2, padx=1, pady=1)
        subj_border.pack(fill="x", pady=(0, 14))
        subj_inner = tk.Frame(subj_border, bg=CARD, padx=12, pady=9)
        subj_inner.pack(fill="x")
        tk.Label(subj_inner, text=subject, bg=CARD, fg=TEXT,
                 font=("Segoe UI", 10), anchor="w",
                 wraplength=580).pack(fill="x")

        # Body
        tk.Label(content, text="BODY", bg=BG, fg=MUTED,
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(0, 3))
        body_border = tk.Frame(content, bg=BORDER2, padx=1, pady=1)
        body_border.pack(fill="both", expand=True, pady=(0, 8))
        body_box = scrolledtext.ScrolledText(body_border,
                                             font=("Segoe UI", 9),
                                             bg=CARD, fg=TEXT,
                                             wrap="word",
                                             relief="flat",
                                             padx=12, pady=10)
        body_box.pack(fill="both", expand=True)
        body_box.insert("end", body)
        body_box.config(state="disabled")

        # Footer
        foot = tk.Frame(win, bg=BG2, pady=8)
        foot.pack(fill="x")
        ttk.Button(foot, text="Close", command=win.destroy).pack()

    # ── Send ────────────────────────────────────────────────────
    def _log(self, message: str):
        self._log_box.configure(state="normal")
        tag = ("sent"  if "[SENT]"  in message else
               "error" if "[ERROR]" in message else
               "skip"  if "[SKIP]"  in message else
               "warn"  if "[WARN]"  in message else "info")
        self._log_box.insert("end", message + "\n", tag)
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _log_safe(self, msg):
        self.after(0, self._log, msg)

    def _check_data(self):
        recs = self._get_recipients()
        cols = self._active_template.all_cols

        win = tk.Toplevel(self)
        win.title("Check Data — Parsed Recipients")
        win.geometry("720x480")
        win.configure(bg=BG)
        win.grab_set()

        # Header
        hdr = tk.Frame(win, bg=HDR_BG, padx=18, pady=13)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Recipient Preview",
                 bg=HDR_BG, fg=HDR_TEXT, font=F_BADGE).pack(side="left")
        tk.Label(hdr, text=f"  {len(recs)} row(s)  ",
                 bg=ACCENT, fg="white",
                 font=("Segoe UI", 8, "bold"),
                 padx=8, pady=3).pack(side="right")
        tk.Frame(win, bg=ACCENT, height=2).pack(fill="x")

        # Table
        frame = tk.Frame(win, bg=BG, padx=12, pady=10)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=cols, show="headings",
                            selectmode="browse")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=max(90, 640 // len(cols)), minwidth=80)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        for r in recs:
            tree.insert("", "end", values=[r.get(c, "") for c in cols])

        if not recs:
            tk.Label(frame,
                     text="No rows found. Make sure Name and Email columns have data.",
                     bg=BG, fg=RED, font=F_LABEL).grid(
                row=2, column=0, columnspan=2, pady=8)

        # Footer
        foot = tk.Frame(win, bg=BG2, pady=8)
        foot.pack(fill="x")
        ttk.Button(foot, text="Close", command=win.destroy).pack()

    def _start_send(self):
        if self._sending:
            messagebox.showwarning("Busy", "Already sending. Please wait.")
            return
        recs = self._get_recipients()
        if not recs:
            messagebox.showwarning(
                "No Recipients",
                "No recipient rows found.\n\n"
                "Tip: Use 'Check Data' button to verify your data is parsed correctly."
            )
            return
        if hasattr(self, "_mismatch_lbl") and self._mismatch_lbl.cget("text"):
            proceed = messagebox.askyesno(
                "Line Mismatch",
                self._mismatch_lbl.cget("text").replace("⚠  ", "")
                + "\n\nSome rows may have missing data.\n\nSend anyway?"
            )
            if not proceed:
                self._nb.select(self._tab_recipients)
                return
        t = self._active_template
        if t.has_attachment and not self._attachment_folder:
            if not messagebox.askyesno("No folder",
                    "No attachment folder selected.\nContinue anyway (no attachments)?"):
                return
                
        # Confirm send
        msg = f"Send {len(recs)} email(s) using template:\n'{t.name}'?\n\n"
            
        if not messagebox.askyesno("Confirm Send", msg):
            return

        self._sending = True
        self._nb.select(self._tab_log)
        self._log(f"── {t.icon} {t.name} — {len(recs)} recipient(s) ──")
        threading.Thread(
            target=send_emails,
            args=(t, self._get_config(), recs,
                  self._attachment_folder,
                  self._log_safe, self._on_done),
            daemon=True
        ).start()

    def _on_done(self, success=True, failed=None):
        failed = failed or []
        def _show():
            self._sending = False
            if success:
                msg = "All emails are sent."
                if failed:
                    msg += f"\n\nFailed ({len(failed)}):\n" + "\n".join(failed)
                self._log("── Done ──")
                messagebox.showinfo("Done", msg)
            else:
                messagebox.showerror("Outlook Error",
                    "Could not connect to Outlook.\n\n"
                    "1. Make sure Outlook is installed and open.\n"
                    "2. Note: The 'New Outlook' (Modern App) does not support COM automation. "
                    "You must use the classic Outlook Desktop App.")
        self.after(0, _show)


# ═══════════════════════════════════════════════════════════════
#  Wrapper classes expected by GST_Suite.py
# ═══════════════════════════════════════════════════════════════

class GSTReturnMailApp(BulkMailApp):
    def __init__(self):
        super().__init__(hide_switcher=True)
        self._switch_template(0)


class InvoiceSenderMailApp(BulkMailApp):
    def __init__(self):
        super().__init__(hide_switcher=True)
        self._switch_template(1)


class PaymentReminderMailApp(BulkMailApp):
    def __init__(self):
        super().__init__(hide_switcher=True)
        self._switch_template(2)


# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = BulkMailApp()
    app.mainloop()
