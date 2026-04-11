import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import win32com.client
import pythoncom
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
        f"In Case of Any Query please feel free to call on {cfg['phone']}.\n\n"
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
        "data_deadline":   "20th March 2026",
        "filing_deadline": "31st March 2026",
        "sender":          "Rohit Sharma",
        "phone":           "+91-98765-43210",
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
        f"at {cfg['phone']}.\n\n"
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
        "phone":    "+91-98765-43210",
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
        f"In Case of Any Query please feel free to call on {cfg['phone']}.\n\n"
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
        "pay_deadline":    "31st March 2026",
        "interest_rate":   "18% p.a.",
        "org_name":        "Your Firm Name",
        "phone":           "+91-98765-43210",
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
#  GUI
# ═══════════════════════════════════════════════════════════════

ACCENT   = "#2563eb"
ACCENT_H = "#1d4ed8"
BG       = "#f8fafc"
BG2      = "#e2e8f0"
CARD     = "#ffffff"
TEXT     = "#1e293b"
MUTED    = "#64748b"
GREEN    = "#16a34a"
RED      = "#dc2626"
YELLOW   = "#d97706"


class BulkMailApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bulk Outlook Mailer")
        self.geometry("820x640")
        self.minsize(720, 540)
        self.configure(bg=BG)

        self._active_template  = ALL_TEMPLATES[0]
        self._cfg_vars         = {}
        self._attachment_folder = ""
        self._sending          = False
        self._add_vars          = {}   # entry vars for the add-recipient form

        self._apply_styles()
        self._build_header()
        self._build_notebook()
        self._build_bottom_bar()

        self._switch_template(0)       # load first template

    # ── Styles ──────────────────────────────────────────────────
    def _apply_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TNotebook",        background=BG,  borderwidth=0)
        s.configure("TNotebook.Tab",    background=BG2, foreground=MUTED,
                    padding=[14, 7], font=("Segoe UI", 9, "bold"))
        s.map("TNotebook.Tab",
              background=[("selected", CARD)],
              foreground=[("selected", ACCENT)])
        s.configure("TFrame",           background=BG)
        s.configure("TLabel",           background=BG,  foreground=TEXT)
        s.configure("TEntry",           fieldbackground=CARD, foreground=TEXT)
        s.configure("Treeview",         background=CARD, fieldbackground=CARD,
                    foreground=TEXT, rowheight=24)
        s.configure("Treeview.Heading", background=BG2, foreground=MUTED,
                    font=("Segoe UI", 8, "bold"))
        s.configure("Send.TButton",     background=ACCENT, foreground="white",
                    font=("Segoe UI", 10, "bold"), padding=[14, 7])
        s.map("Send.TButton",           background=[("active", ACCENT_H)])
        s.configure("Card.TFrame",      background=CARD, relief="flat")

    # ── Header ──────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=ACCENT, pady=12)
        hdr.pack(fill="x")

        tk.Label(hdr, text="  Bulk Outlook Mailer", bg=ACCENT, fg="white",
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=4)

        # Template picker on the right
        right = tk.Frame(hdr, bg=ACCENT)
        right.pack(side="right", padx=14)
        tk.Label(right, text="Template:", bg=ACCENT, fg="#bfdbfe",
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

        self._tmpl_var = tk.StringVar()
        names = [f"{t.icon}  {t.name}" for t in ALL_TEMPLATES]
        cb = ttk.Combobox(right, textvariable=self._tmpl_var, values=names,
                          state="readonly", width=28,
                          font=("Segoe UI", 9))
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self._on_template_change())
        self._tmpl_cb = cb

    # ── Notebook ────────────────────────────────────────────────
    def _build_notebook(self):
        self._nb = ttk.Notebook(self)
        self._nb.pack(fill="both", expand=True, padx=12, pady=(10, 0))

        self._tab_config     = ttk.Frame(self._nb, padding=14)
        self._tab_recipients = ttk.Frame(self._nb, padding=14)
        self._tab_log        = ttk.Frame(self._nb, padding=14)

        self._nb.add(self._tab_config,     text="  Configuration  ")
        self._nb.add(self._tab_recipients, text="  Recipients  ")
        self._nb.add(self._tab_log,        text="  Log  ")

        self._build_log_tab()

    # ── Bottom bar ──────────────────────────────────────────────
    def _build_bottom_bar(self):
        bar = tk.Frame(self, bg=BG2, pady=9)
        bar.pack(fill="x", side="bottom")

        self._status_lbl = tk.Label(bar, text="", bg=BG2, fg=MUTED,
                                    font=("Segoe UI", 8))
        self._status_lbl.pack(side="left", padx=14)

        ttk.Button(bar, text="  Send All Emails  ", style="Send.TButton",
                   command=self._start_send).pack(side="right", padx=12)
        ttk.Button(bar, text="  Preview Email  ",
                   command=self._show_preview).pack(side="right", padx=4)
        ttk.Button(bar, text="Check Data",
                   command=self._check_data).pack(side="right", padx=4)

    # ── Config tab ──────────────────────────────────────────────
    def _rebuild_config_tab(self):
        for w in self._tab_config.winfo_children():
            w.destroy()
        self._cfg_vars = {}
        t = self._active_template

        # Template description badge
        badge = tk.Frame(self._tab_config, bg=CARD, pady=6, padx=10)
        badge.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        tk.Label(badge, text=f"{t.icon}  {t.name}", bg=CARD,
                 font=("Segoe UI", 10, "bold"), fg=ACCENT).pack(anchor="w")

        for i, (label, key) in enumerate(t.config_fields, start=1):
            tk.Label(self._tab_config, text=label, anchor="w",
                     font=("Segoe UI", 9), fg=MUTED).grid(
                row=i, column=0, sticky="w", pady=5, padx=(0, 14))
            var = tk.StringVar(value=t.defaults.get(key, ""))
            e = ttk.Entry(self._tab_config, textvariable=var, width=44)
            e.grid(row=i, column=1, sticky="ew", pady=5)
            self._cfg_vars[key] = var

        self._tab_config.columnconfigure(1, weight=1)

    # ── Recipients tab ──────────────────────────────────────────
    def _rebuild_recipients_tab(self):
        for w in self._tab_recipients.winfo_children():
            w.destroy()
        t = self._active_template
        cols = t.all_cols

        # Import from Excel button (all templates)
        imp_row = tk.Frame(self._tab_recipients, bg=BG)
        imp_row.pack(fill="x", pady=(0, 4))
        ttk.Button(imp_row, text="📂  Import from Excel",
                   command=self._import_excel).pack(side="left")

        # Attachment folder row (Invoice only)
        if t.has_attachment:
            att_row = tk.Frame(self._tab_recipients, bg=BG)
            att_row.pack(fill="x", pady=(0, 6))
            tk.Label(att_row, text="Attachment Folder:", fg=MUTED,
                     font=("Segoe UI", 9), bg=BG).pack(side="left")
            self._folder_lbl = tk.Label(att_row, text="  (not selected)",
                                        bg=CARD, relief="groove", padx=6,
                                        font=("Segoe UI", 9), fg=MUTED, width=38,
                                        anchor="w")
            self._folder_lbl.pack(side="left", padx=6)
            ttk.Button(att_row, text="Browse…",
                       command=self._pick_folder).pack(side="left")

        # ── Multi-column editor with line numbers + dividers ───
        editor_outer = tk.Frame(self._tab_recipients, bg=BG)
        editor_outer.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(editor_outer, orient="vertical")
        vsb.pack(side="right", fill="y")

        # Canvas for horizontal scrolling
        h_canvas = tk.Canvas(editor_outer, bg=BG, highlightthickness=0)
        hsb = ttk.Scrollbar(editor_outer, orient="horizontal",
                             command=h_canvas.xview)
        h_canvas.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        h_canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(h_canvas, bg="#c0c8d8")
        _canvas_win = h_canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_resize(e):
            h_canvas.configure(scrollregion=h_canvas.bbox("all"))
        def _on_canvas_resize(e):
            # Stretch inner to canvas height
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
                # Sync scroll position
                ln_w.yview_moveto(tw.yview()[0])

        def _on_any_change():
            _refresh_linenos()
            self._update_count()
            self._check_line_mismatch()

        COL_MIN_W = 220   # px per column — all columns always visible
        for col_name in cols:
            col_frame = tk.Frame(inner, bg=BG, width=COL_MIN_W)
            col_frame.pack(side="left", fill="y", padx=(0, 1))
            col_frame.pack_propagate(False)

            # Header
            tk.Label(col_frame, text=col_name,
                     bg="#dde3ed", fg=MUTED,
                     font=("Consolas", 8, "bold"),
                     anchor="w", padx=8, pady=4).pack(fill="x")

            # Row: line-number gutter + text area
            row = tk.Frame(col_frame, bg=CARD)
            row.pack(fill="both", expand=True)

            ln_w = tk.Text(row, width=3, font=("Consolas", 10),
                           bg="#f0f4f8", fg="#94a3b8",
                           relief="flat", state="disabled",
                           cursor="arrow", takefocus=False,
                           yscrollcommand=lambda *a: None)
            ln_w.pack(side="left", fill="y")
            tk.Frame(row, bg="#dde3ed", width=1).pack(side="left", fill="y")
            self._lineno_widgets.append(ln_w)

            tw = tk.Text(row, font=("Consolas", 10),
                         bg=CARD, fg=TEXT,
                         insertbackground=ACCENT,
                         selectbackground="#bfdbfe",
                         relief="flat", wrap="none", undo=True,
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
        bot = tk.Frame(self._tab_recipients, bg=BG2, pady=3)
        bot.pack(fill="x")
        self._count_lbl = tk.Label(bot, text="0 recipients",
                                   bg=BG2, fg=MUTED, font=("Segoe UI", 8))
        self._count_lbl.pack(side="left", padx=8)
        self._mismatch_lbl = tk.Label(bot, text="",
                                      bg=BG2, fg="#dc2626",
                                      font=("Segoe UI", 8, "bold"))
        self._mismatch_lbl.pack(side="left", padx=12)
        tk.Label(bot, text="Ctrl+V pastes from Excel",
                 bg=BG2, fg=MUTED, font=("Segoe UI", 7)).pack(side="right", padx=8)

        _refresh_linenos()

        # Expose a single _rec_text alias for _get_recipients / _update_count
        self._rec_text = self._col_texts[0]   # primary column for line counting
        self._update_count()

    # ── Log tab ─────────────────────────────────────────────────
    def _build_log_tab(self):
        toolbar = tk.Frame(self._tab_log, bg=BG)
        toolbar.pack(fill="x", pady=(0, 6))
        tk.Label(toolbar, text="Send Log", bg=BG, font=("Segoe UI", 9, "bold"),
                 fg=TEXT).pack(side="left")
        ttk.Button(toolbar, text="Clear", command=self._clear_log).pack(side="right")

        self._log_box = scrolledtext.ScrolledText(
            self._tab_log, height=20, font=("Consolas", 9),
            state="disabled", bg="#0f172a", fg="#94a3b8",
            insertbackground="white", relief="flat")
        self._log_box.pack(fill="both", expand=True)
        self._log_box.tag_configure("sent",  foreground="#4ade80")
        self._log_box.tag_configure("error", foreground="#f87171")
        self._log_box.tag_configure("skip",  foreground="#fbbf24")
        self._log_box.tag_configure("warn",  foreground="#fb923c")
        self._log_box.tag_configure("info",  foreground="#93c5fd")

    # ── Template switching ───────────────────────────────────────
    def _on_template_change(self):
        idx = self._tmpl_cb.current()
        if idx >= 0:
            self._switch_template(idx)

    def _switch_template(self, idx):
        self._active_template = ALL_TEMPLATES[idx]
        self._tmpl_var.set(f"{self._active_template.icon}  {self._active_template.name}")
        self._attachment_folder = ""
        self._rebuild_config_tab()
        self._rebuild_recipients_tab()
        self._status_lbl.config(text=f"Template: {self._active_template.name}")

    # ── Helpers ─────────────────────────────────────────────────
    def _get_config(self):
        return {k: v.get() for k, v in self._cfg_vars.items()}

    def _get_recipients(self):
        cols     = self._active_template.all_cols
        col_lines = [tw.get("1.0", "end").splitlines()
                     for tw in self._col_texts]
        # Pad all columns to the same length
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

        # Try to match first row as headers
        first_row = [str(v).strip() if v is not None else "" for v in rows[0]]
        header_map = {}   # excel_col_index → col_texts_index
        for xi, hdr in enumerate(first_row):
            hl = hdr.lower()
            if hl in cols_lower:
                header_map[xi] = cols_lower.index(hl)

        if header_map:
            data_rows = rows[1:]   # skip header row
        else:
            # fallback: positional mapping
            data_rows = rows
            header_map = {xi: xi for xi in range(min(len(first_row), len(cols)))}

        # Clear all columns
        for tw in self._col_texts:
            tw.delete("1.0", "end")

        # Populate
        for row in data_rows:
            if not any(v for v in row if v is not None and str(v).strip()):
                continue  # skip blank rows
            for xi, ci in header_map.items():
                val = row[xi] if xi < len(row) else ""
                val = "" if val is None else str(val).strip()
                existing = self._col_texts[ci].get("1.0", "end").rstrip("\n")
                if existing:
                    self._col_texts[ci].insert("end", "\n")
                self._col_texts[ci].insert("end", val)

        # Refresh UI
        if hasattr(self, "_rec_text"):
            self._update_count()
        self._check_line_mismatch()
        # Trigger line number refresh via keyrelease simulation
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
        # Only flag columns that have SOME data (ignore fully-empty optional columns)
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
        win.geometry("640x520")
        win.configure(bg=BG)
        win.grab_set()

        # Badge
        badge = tk.Frame(win, bg=CARD, padx=12, pady=8)
        badge.pack(fill="x", padx=12, pady=(12, 0))
        tk.Label(badge, text=f"{t.icon}  {t.name}", bg=CARD,
                 font=("Segoe UI", 10, "bold"), fg=ACCENT).pack(side="left")
        if recs:
            tk.Label(badge, text=f"(preview: {sample_row.get('Name','—')})",
                     bg=CARD, fg=MUTED, font=("Segoe UI", 8)).pack(side="right")

        # Subject
        subj_frame = tk.Frame(win, bg=BG, padx=12, pady=6)
        subj_frame.pack(fill="x")
        tk.Label(subj_frame, text="Subject:", fg=MUTED,
                 font=("Segoe UI", 8, "bold"), bg=BG).pack(anchor="w")
        subj_entry = tk.Entry(subj_frame, font=("Segoe UI", 9),
                              fg=TEXT, bg=CARD, relief="groove")
        subj_entry.pack(fill="x", pady=(2, 0))
        subj_entry.insert(0, subject)
        subj_entry.config(state="readonly")

        # Body
        body_frame = tk.Frame(win, bg=BG, padx=12, pady=6)
        body_frame.pack(fill="both", expand=True)
        tk.Label(body_frame, text="Body:", fg=MUTED,
                 font=("Segoe UI", 8, "bold"), bg=BG).pack(anchor="w")
        body_box = scrolledtext.ScrolledText(body_frame, font=("Segoe UI", 9),
                                             bg=CARD, fg=TEXT, wrap="word",
                                             relief="groove")
        body_box.pack(fill="both", expand=True, pady=(2, 0))
        body_box.insert("end", body)
        body_box.config(state="disabled")

        ttk.Button(win, text="Close", command=win.destroy).pack(pady=10)

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
        """Show parsed recipient rows so user can verify before sending."""
        recs = self._get_recipients()
        cols = self._active_template.all_cols

        win = tk.Toplevel(self)
        win.title("Check Data — Parsed Recipients")
        win.geometry("700x420")
        win.configure(bg=BG)
        win.grab_set()

        tk.Label(win, text=f"{len(recs)} recipient(s) will be used for sending",
                 bg=BG, font=("Segoe UI", 9, "bold"), fg=ACCENT).pack(pady=(10, 4))

        frame = tk.Frame(win, bg=BG)
        frame.pack(fill="both", expand=True, padx=10)

        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=max(80, 600 // len(cols)))
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        for r in recs:
            tree.insert("", "end", values=[r.get(c, "") for c in cols])

        if not recs:
            tk.Label(win, text="No rows found. Make sure Name and Email columns have data.",
                     bg=BG, fg=RED, font=("Segoe UI", 9)).pack(pady=4)

        ttk.Button(win, text="Close", command=win.destroy).pack(pady=8)

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
        # Warn (not block) if line counts differ — let user decide
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
        if not messagebox.askyesno("Confirm Send",
                f"Send {len(recs)} email(s) using template:\n'{t.name}'?"):
            return

        self._sending = True
        self._nb.select(self._tab_log)
        self._log(f"── {t.icon} {t.name} — {len(recs)} recipient(s) ──")
        threading.Thread(
            target=send_emails,
            args=(t, self._get_config(), recs,
                  self._attachment_folder, self._log_safe, self._on_done),
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
                    "Could not connect to Outlook.\n"
                    "Make sure Outlook is installed and signed in.")
        self.after(0, _show)


# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = BulkMailApp()
    app.mainloop()
