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
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    import os
    import json

    cred_path = os.path.join(os.path.dirname(__file__), "gmail_credentials.json")
    sender_email = ""
    app_password = ""
    try:
        if os.path.exists(cred_path):
            with open(cred_path, "r") as f:
                creds = json.load(f)
                sender_email = creds.get("email", "")
                app_password = creds.get("password", "")
    except Exception:
        pass

    if not sender_email or not app_password:
        log_cb("[ERROR] Sender Gmail or App Password not configured.")
        done_cb(success=False)
        return

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, app_password)
    except Exception as exc:
        log_cb(f"[ERROR] Could not connect to Gmail SMTP: {exc}")
        done_cb(success=False)
        return

    failed = []
    for idx, row in enumerate(recipients, 1):
        name  = str(row.get("Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        if not email:
            log_cb(f"[SKIP]  Row {idx} - empty email.")
            continue
        try:
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = email

            cc_email = str(cfg.get("cc", "")).strip()
            if cc_email:
                msg["Cc"] = cc_email

            msg["Subject"] = template.build_subject(cfg, row)
            body_text = template.build_body(cfg, row)
            html_body = (
                "<html><body><pre style='font-family:Calibri,sans-serif;font-size:11pt;white-space:pre-wrap'>"
                + body_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                + "</pre></body></html>"
            )
            msg.attach(MIMEText(html_body, "html"))

            if template.has_attachment and attachment_folder:
                fname = str(row.get("Attachment File", "")).strip()
                if fname:
                    base = fname[:-4] if fname.lower().endswith(".pdf") else fname
                    path = os.path.join(attachment_folder, base + ".pdf")
                    if os.path.exists(path):
                        with open(path, "rb") as att:
                            part = MIMEApplication(att.read(), Name=os.path.basename(path))
                        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(path)}"'
                        msg.attach(part)
                    else:
                        log_cb(f"[WARN]  {name} - attachment not found: {base}.pdf")

            all_recipients = [email]
            if cc_email:
                all_recipients.extend([rcpt.strip() for rcpt in cc_email.split(",") if rcpt.strip()])

            server.send_message(msg, from_addr=sender_email, to_addrs=all_recipients)
            log_cb(f"[SENT]  {name} <{email}>")

        except Exception as exc:
            log_cb(f"[ERROR] {name} <{email}> - {exc}")
            failed.append(email)

    try:
        server.quit()
    except:
        pass
    
    if failed:
        log_cb(f"\\n[DONE] Finished with {len(failed)} errors.")
        done_cb(success=False)
    else:
        log_cb("\\n[DONE] All emails sent successfully!")
        done_cb(success=True)

        return

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, app_password)
    except Exception as exc:
        log_cb(f"[ERROR] Could not connect to Gmail SMTP: {exc}")
        done_cb(success=False)
        return

    failed = []
    for idx, row in enumerate(recipients, 1):
        name  = str(row.get("Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        if not email:
            log_cb(f"[SKIP]  Row {idx} - empty email.")
            continue
        try:
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = email

            cc_email = str(cfg.get("cc", "")).strip()
            if cc_email:
                msg["Cc"] = cc_email

            msg["Subject"] = template.build_subject(cfg, row)
            body_text = template.build_body(cfg, row)
            html_body = (
                "<html><body><pre style='font-family:Calibri,sans-serif;font-size:11pt;white-space:pre-wrap'>"
                + body_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                + "</pre></body></html>"
            )
            msg.attach(MIMEText(html_body, "html"))

            if template.has_attachment and attachment_folder:
                fname = str(row.get("Attachment File", "")).strip()
                if fname:
                    base = fname[:-4] if fname.lower().endswith(".pdf") else fname
                    path = os.path.join(attachment_folder, base + ".pdf")
                    if os.path.exists(path):
                        with open(path, "rb") as att:
                            part = MIMEApplication(att.read(), Name=os.path.basename(path))
                        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(path)}"'
                        msg.attach(part)
                    else:
                        log_cb(f"[WARN]  {name} - attachment not found: {base}.pdf")

            all_recipients = [email]
            if cc_email:
                all_recipients.extend([rcpt.strip() for rcpt in cc_email.split(",") if rcpt.strip()])

            server.send_message(msg, from_addr=sender_email, to_addrs=all_recipients)
            log_cb(f"[SENT]  {name} <{email}>")

        except Exception as exc:
            log_cb(f"[ERROR] {name} <{email}> - {exc}")
            failed.append(email)

    try:
        server.quit()
    except:
        pass
    
    if failed:
        log_cb(f"\\n[DONE] Finished with {len(failed)} errors.")
        done_cb(success=False)
    else:
        log_cb("\\n[DONE] All emails sent successfully!")
        done_cb(success=True)

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
#  GUI  —  Design Tokens & Setup
# ═══════════════════════════════════════════════════════════════
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import calendar
import os
import threading

# Set default theme
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

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

class CustomDatePicker(ctk.CTkToplevel):
    def __init__(self, parent, target_var):
        super().__init__(parent)
        self.target_var = target_var
        self.title("Select Date")
        self.geometry("320x360")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.current_date = datetime.now()
        
        self._build_ui()
        self._update_calendar()
        
    def _build_ui(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkButton(header, text="<", width=30, command=self._prev_month).pack(side="left")
        self.month_lbl = ctk.CTkLabel(header, text="", font=F_BTN_LG)
        self.month_lbl.pack(side="left", expand=True)
        ctk.CTkButton(header, text=">", width=30, command=self._next_month).pack(side="right")
        
        self.cal_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.cal_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for c, day in enumerate(days):
            ctk.CTkLabel(self.cal_frame, text=day, font=F_LABEL_SML, text_color="gray").grid(row=0, column=c, padx=3, pady=3)
            
        self.day_buttons = []
        for r in range(6):
            row_btns = []
            for c in range(7):
                btn = ctk.CTkButton(self.cal_frame, text="", width=35, height=35, fg_color="transparent", text_color=("black", "white"), font=F_LABEL, hover_color=("gray80", "gray30"), command=lambda r=r, c=c: self._select_day(r, c))
                btn.grid(row=r+1, column=c, padx=2, pady=2)
                row_btns.append(btn)
            self.day_buttons.append(row_btns)
            
    def _update_calendar(self):
        self.month_lbl.configure(text=self.current_date.strftime("%B %Y"))
        
        cal = calendar.monthcalendar(self.current_date.year, self.current_date.month)
        
        for r in range(6):
            for c in range(7):
                if r < len(cal) and cal[r][c] != 0:
                    day = cal[r][c]
                    self.day_buttons[r][c].configure(text=str(day), state="normal")
                else:
                    self.day_buttons[r][c].configure(text="", state="disabled")
                    
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

class BulkMailApp(ctk.CTk):
    def __init__(self, hide_switcher=False):
        super().__init__()
        self.title("Bulk Gmail Mailer")
        self.geometry("900x700")
        self.minsize(740, 560)
        self._hide_switcher = hide_switcher

        self._active_template  = ALL_TEMPLATES[0]
        self._cfg_vars         = {}
        self._attachment_folder = ""
        self._sending          = False
        self._add_vars         = {}

        self._build_ui()
        self._switch_template(0)

    def set_theme(self, mode: str):
        ctk.set_appearance_mode(mode)
        style = ttk.Style()
        is_dark = mode == "Dark"
        bg_color = "#1e2d3d" if is_dark else "white"
        fg_color = "white" if is_dark else "black"
        hl_color = "#1e3a6e" if is_dark else "#DBEAFE"
        hdr_bg = "#1e293b" if is_dark else "#F1F5F9"
        style.configure("Custom.Treeview", background=bg_color, foreground=fg_color, fieldbackground=bg_color)
        style.map("Custom.Treeview", background=[("selected", hl_color)], foreground=[("selected", fg_color)])
        style.configure("Custom.Treeview.Heading", background=hdr_bg, foreground=fg_color)

    def _build_ui(self):
        self.bottom_bar = ctk.CTkFrame(self, height=60, corner_radius=0)
        self.bottom_bar.pack(fill="x", side="bottom")

        self._status_lbl = ctk.CTkLabel(self.bottom_bar, text="", font=F_LABEL_SML, text_color="gray")
        self._status_lbl.pack(side="left", padx=20)

        ctk.CTkButton(self.bottom_bar, text="  Send All Emails  →", font=F_BTN_LG, height=45, fg_color="#2563EB", hover_color="#1D4ED8", command=self._start_send).pack(side="right", padx=(10, 20), pady=10)
        ctk.CTkButton(self.bottom_bar, text="Preview Email", height=45, fg_color="transparent", border_width=1, text_color=("black", "white"), command=self._show_preview).pack(side="right", padx=(5, 0), pady=10)
        ctk.CTkButton(self.bottom_bar, text="Check Data", height=45, fg_color="transparent", border_width=1, text_color=("black", "white"), command=self._check_data).pack(side="right", padx=(5, 0), pady=10)

        self._nb = ctk.CTkTabview(self)
        self._nb.pack(fill="both", expand=True, padx=20, pady=(10, 10))

        self._tab_config     = self._nb.add("Configuration")
        self._tab_recipients = self._nb.add("Recipients")
        self._tab_log        = self._nb.add("Log")

        self._build_log_tab()


    def _open_email_settings(self):
        import os
        import json
        cred_path = os.path.join(os.path.dirname(__file__), "gmail_credentials.json")
        email_val = ""
        pass_val = ""
        if os.path.exists(cred_path):
            try:
                with open(cred_path, "r") as f:
                    data = json.load(f)
                    email_val = data.get("email", "")
                    pass_val = data.get("password", "")
            except:
                pass

        top = ctk.CTkToplevel(self)
        top.title("⚙ Email Configuration")
        top.geometry("440x280")
        top.transient(self)
        top.grab_set()

        title_fr = ctk.CTkFrame(top, fg_color="transparent")
        title_fr.pack(fill="x", pady=(20, 10))
        ctk.CTkLabel(title_fr, text="⚙", font=("Segoe UI Emoji", 16), text_color="#4da8da").pack(side="left", padx=(30, 10))
        ctk.CTkLabel(title_fr, text="Gmail Details", font=("Segoe UI", 14, "bold")).pack(side="left")

        frame = ctk.CTkFrame(top, corner_radius=8, border_width=1)
        frame.pack(padx=25, fill="x", pady=10)

        ctk.CTkLabel(frame, text="Gmail Address:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5, padx=10)
        e_email = ctk.CTkEntry(frame, width=200)
        e_email.grid(row=0, column=1, sticky="w", pady=5, padx=10)
        e_email.insert(0, email_val)

        ctk.CTkLabel(frame, text="App Password:", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", pady=(15, 5), padx=10)
        e_pass = ctk.CTkEntry(frame, width=200, show="*")
        e_pass.grid(row=1, column=1, sticky="w", pady=(15, 5), padx=10)
        e_pass.insert(0, pass_val)

        def save():
            with open(cred_path, "w") as f:
                json.dump({"email": e_email.get().strip(), "password": e_pass.get().strip()}, f)
            top.destroy()
            from tkinter import messagebox
            messagebox.showinfo("Saved", "Settings saved successfully!", parent=self)

        ctk.CTkButton(top, text="Save Settings", command=save).pack(pady=(10, 0))

    def _rebuild_config_tab(self):
        for w in self._tab_config.winfo_children():
            w.destroy()
        self._cfg_vars = {}
        t = self._active_template

        scroll = ctk.CTkScrollableFrame(self._tab_config, fg_color="transparent", corner_radius=0)
        scroll.pack(fill="both", expand=True)
        scroll.columnconfigure(0, weight=1)

        form_card = ctk.CTkFrame(scroll, corner_radius=10)
        form_card.pack(fill="x", anchor="n", padx=30, pady=20)
        form_card.columnconfigure(1, weight=1)

        self._row_i = 0
        lbl_text = "Email Template" if not self._hide_switcher else f"{t.icon}  {t.name} Config"
        ctk.CTkLabel(form_card, text=lbl_text, font=F_TITLE).grid(row=self._row_i, column=0, sticky="w", pady=20, padx=30)

        cb_frame = ctk.CTkFrame(form_card, fg_color="transparent")
        cb_frame.grid(row=self._row_i, column=1, sticky="w", pady=20)

        if not self._hide_switcher:
            self._tmpl_var = ctk.StringVar(value=f"{t.icon}  {t.name}")
            names = [f"{tmpl.icon}  {tmpl.name}" for tmpl in ALL_TEMPLATES]
            
            cb = ctk.CTkComboBox(cb_frame, variable=self._tmpl_var, values=names, state="readonly", width=300, command=self._on_template_change)
            cb.pack(side="left")
            self._tmpl_cb = cb

        settings_btn = ctk.CTkButton(cb_frame, text="⚙ Email Settings", command=self._open_email_settings)
        settings_btn.pack(side="right", padx=(0, 0) if self._hide_switcher else (14, 0))
        
        import webbrowser
        link = ctk.CTkLabel(cb_frame, text="How to setup email 🔗", font=("Segoe UI", 10, "underline"), cursor="hand2", text_color="#4da8da")
        link.pack(side="right", padx=(14, 14))
        link.bind("<Button-1>", lambda e: webbrowser.open("https://www.youtube.com/watch?v=MkLX85XU5rU"))


        self._row_i += 1

        def validate_phone(P):
            if P == "": return True
            if P.isdigit() and len(P) <= 10: return True
            return False
        vcmd_phone = (self.register(validate_phone), "%P")

        for i, (label, key) in enumerate(t.config_fields):
            ctk.CTkLabel(form_card, text=label, font=F_LABEL).grid(row=self._row_i+i, column=0, sticky="w", pady=10, padx=30)
            var = ctk.StringVar(value=t.defaults.get(key, ""))
            self._cfg_vars[key] = var

            if key == "phone":
                inner_f = ctk.CTkFrame(form_card, fg_color="transparent")
                inner_f.grid(row=self._row_i+i, column=1, sticky="ew", pady=10, padx=(0, 30))
                ctk.CTkLabel(inner_f, text="+91", font=F_INPUT).pack(side="left", padx=(0, 10))
                e = ctk.CTkEntry(inner_f, textvariable=var, font=F_INPUT, validate="key", validatecommand=vcmd_phone)
                e.pack(side="left", fill="x", expand=True)
            elif "deadline" in key.lower() or "date" in key.lower():
                inner_f = ctk.CTkFrame(form_card, fg_color="transparent")
                inner_f.grid(row=self._row_i+i, column=1, sticky="ew", pady=10, padx=(0, 30))
                e = ctk.CTkEntry(inner_f, textvariable=var, font=F_INPUT)
                e.pack(side="left", fill="x", expand=True)
                btn = ctk.CTkButton(inner_f, text="📅", width=40, fg_color="transparent", text_color=("black", "white"), border_width=1, command=lambda v=var: CustomDatePicker(self, v))
                btn.pack(side="right", padx=(5, 0))
            else:
                e = ctk.CTkEntry(form_card, textvariable=var, font=F_INPUT)
                e.grid(row=self._row_i+i, column=1, sticky="ew", pady=10, padx=(0, 30))

    def _rebuild_recipients_tab(self):
        for w in self._tab_recipients.winfo_children():
            w.destroy()
        t = self._active_template
        cols = t.all_cols

        toolbar = ctk.CTkFrame(self._tab_recipients, fg_color="transparent")
        toolbar.pack(fill="x", pady=10)

        ctk.CTkButton(toolbar, text="📂  Import from Excel", command=self._import_excel, width=150).pack(side="left", padx=(0, 10))
        ctk.CTkButton(toolbar, text="⬇  Download Sample", fg_color="transparent", border_width=1, text_color=("black", "white"), command=self._download_sample, width=150).pack(side="left", padx=(0, 10))
        ctk.CTkButton(toolbar, text="🗑  Clear All", fg_color="#EF4444", hover_color="#DC2626", text_color="white", command=self._clear_recipients, width=100).pack(side="left", padx=(0, 10))

        if t.has_attachment:
            self._folder_lbl = ctk.CTkLabel(toolbar, text="(not selected)", fg_color=("gray90", "gray20"), corner_radius=5, width=200)
            self._folder_lbl.pack(side="left", padx=(10, 10))
            ctk.CTkButton(toolbar, text="Browse…", width=80, fg_color="transparent", border_width=1, text_color=("black", "white"), command=self._pick_folder).pack(side="left")

        self._tree_frame = ctk.CTkFrame(self._tab_recipients, fg_color="transparent")
        self._tree_frame.pack(fill="both", expand=True, pady=(10, 0))

        style = ttk.Style()
        style.theme_use("clam")
        
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg_color = "#1e2d3d" if is_dark else "white"
        fg_color = "white" if is_dark else "black"
        hl_color = "#1e3a6e" if is_dark else "#DBEAFE"
        hdr_bg = "#1e293b" if is_dark else "#F1F5F9"
        
        style.configure("Custom.Treeview", background=bg_color, foreground=fg_color, fieldbackground=bg_color, rowheight=30, borderwidth=0, font=F_LABEL_SML)
        style.map("Custom.Treeview", background=[("selected", hl_color)], foreground=[("selected", "white")])
        style.configure("Custom.Treeview.Heading", background=hdr_bg, foreground=fg_color, font=F_BTN, padding=5)

        self._tree = ttk.Treeview(self._tree_frame, columns=cols, show="headings", style="Custom.Treeview")
        for col in cols:
            self._tree.heading(col, text=col)
            self._tree.column(col, width=150, minwidth=100)

        vsb = ttk.Scrollbar(self._tree_frame, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(self._tree_frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self._tree_frame.rowconfigure(0, weight=1)
        self._tree_frame.columnconfigure(0, weight=1)

        self._tree.bind("<Control-v>", self._on_paste)
        self._tree.bind("<Control-V>", self._on_paste)
        self._tree.bind("<Double-1>", self._on_double_click)
        self._tree.bind("<Delete>", self._on_delete)
        self._tree.bind("<BackSpace>", self._on_delete)

        bot = ctk.CTkFrame(self._tab_recipients, fg_color="transparent")
        bot.pack(fill="x", pady=(5, 0))
        self._count_lbl = ctk.CTkLabel(bot, text="0 recipients", font=F_LABEL_SML)
        self._count_lbl.pack(side="left")
        ctk.CTkLabel(bot, text="Ctrl+V: Paste | Double-click: Edit | Del: Remove", font=F_LABEL_SML, text_color="gray").pack(side="right")

    def _on_paste(self, event):
        try:
            raw = self.clipboard_get()
        except tk.TclError:
            return "break"
        lines = [ln for ln in raw.splitlines() if ln.strip()]
        if not lines:
            return "break"
        cols = self._active_template.all_cols
        for line in lines:
            parts = line.split("\t")
            parts = (parts + [""] * len(cols))[:len(cols)]
            self._tree.insert("", "end", values=parts)
        self._update_count()
        return "break"

    def _on_delete(self, event):
        selected = self._tree.selection()
        for item in selected:
            self._tree.delete(item)
        self._update_count()

    def _on_double_click(self, event):
        region = self._tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self._tree.identify_column(event.x)
        row_id = self._tree.identify_row(event.y)
        if not column or not row_id: return
        
        col_index = int(column[1:]) - 1
        x, y, width, height = self._tree.bbox(row_id, column)
        val = self._tree.set(row_id, column)

        entry = tk.Entry(self._tree, font=F_LABEL_SML)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, val)
        entry.focus()

        def save_edit(event=None):
            new_val = entry.get()
            self._tree.set(row_id, column, new_val)
            entry.destroy()
            self._update_count()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    def _build_log_tab(self):
        hdr = ctk.CTkFrame(self._tab_log, fg_color="transparent")
        hdr.pack(fill="x", pady=(10, 5))

        ctk.CTkLabel(hdr, text="Send Log", font=F_BTN).pack(side="left")
        ctk.CTkButton(hdr, text="Clear Log", width=100, fg_color="transparent", border_width=1, text_color=("black", "white"), command=self._clear_log).pack(side="right")

        self._log_box = ctk.CTkTextbox(self._tab_log, font=F_MONO, state="disabled")
        self._log_box.pack(fill="both", expand=True)

    def _on_template_change(self, value=None):
        idx = -1
        names = [f"{tmpl.icon}  {tmpl.name}" for tmpl in ALL_TEMPLATES]
        val = self._tmpl_var.get()
        if val in names:
            idx = names.index(val)
        if idx >= 0:
            self._switch_template(idx)

    def _switch_template(self, idx):
        self._active_template = ALL_TEMPLATES[idx]
        self._attachment_folder = ""
        self._rebuild_config_tab()
        self._rebuild_recipients_tab()
        if hasattr(self, "_status_lbl") and hasattr(self._status_lbl, "configure"):
            self._status_lbl.configure(text="")

    def _get_config(self):
        return {k: v.get() for k, v in self._cfg_vars.items()}

    def _get_recipients(self):
        rows = []
        cols = self._active_template.all_cols
        for item in self._tree.get_children():
            values = self._tree.item(item, 'values')
            row = dict(zip(cols, values))
            if any(row.values()):
                rows.append(row)
        return rows

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="Select Attachment Folder")
        if folder:
            self._attachment_folder = folder
            short = folder if len(folder) < 30 else "…" + folder[-27:]
            self._folder_lbl.configure(text=short)

    def _clear_recipients(self):
        if self._tree.get_children():
            if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all recipients?", parent=self):
                for item in self._tree.get_children():
                    self._tree.delete(item)
                self._update_count()

    def _download_sample(self):
        import shutil
        t = self._active_template
        mapping = {
            "GST Return Data Request": "GST_Return_Recipients.xlsx",
            "Invoice Sender": "Invoice_Sender_Recipients.xlsx",
            "Outstanding Payment Reminder": "Outstanding_Payment_Recipients.xlsx"
        }
        filename = mapping.get(t.name)
        if not filename:
            messagebox.showerror("Error", "No sample file configured for this template.")
            return

        base_dir = os.path.dirname(os.path.abspath(__file__))
        src_path = os.path.join(base_dir, "Input For Email Upload", filename)

        if not os.path.exists(src_path):
            messagebox.showerror("Not Found", f"Sample file not found at:\n{src_path}")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save Sample Excel File",
            initialfile=filename,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            parent=self
        )
        if save_path:
            try:
                shutil.copy2(src_path, save_path)
                messagebox.showinfo("Success", f"Sample saved successfully to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

    def _import_excel(self):
        if not _OPENPYXL:
            messagebox.showerror("Missing Library", "openpyxl is not installed.\n\nRun: pip install openpyxl")
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

        for item in self._tree.get_children():
            self._tree.delete(item)

        for row in data_rows:
            if not any(v for v in row if v is not None and str(v).strip()):
                continue
            item_vals = [""] * len(cols)
            for xi, ci in header_map.items():
                val = row[xi] if xi < len(row) else ""
                item_vals[ci] = "" if val is None else str(val).strip()
            self._tree.insert("", "end", values=item_vals)

        self._update_count()

    def _update_count(self):
        n = len(self._tree.get_children())
        self._count_lbl.configure(text=f"{n} recipient{'s' if n != 1 else ''}")

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _show_preview(self):
        cfg  = self._get_config()
        recs = self._get_recipients()
        t    = self._active_template

        sample_row = recs[0] if recs else {c: f"<{c}>" for c in t.all_cols}
        subject    = t.build_subject(cfg, sample_row)
        body       = t.build_body(cfg, sample_row)

        win = ctk.CTkToplevel(self)
        win.title("Email Preview")
        win.geometry("660x560")
        win.grab_set()

        hdr = ctk.CTkFrame(win, corner_radius=0)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text=f"{t.icon}  {t.name}", font=F_BADGE).pack(side="left", padx=20, pady=10)

        content = ctk.CTkFrame(win, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=10)

        ctk.CTkLabel(content, text="SUBJECT", font=("Segoe UI", 10, "bold"), text_color="gray").pack(anchor="w")
        ctk.CTkLabel(content, text=subject, font=F_LABEL, wraplength=600, justify="left").pack(anchor="w", pady=(0, 20))

        ctk.CTkLabel(content, text="BODY", font=("Segoe UI", 10, "bold"), text_color="gray").pack(anchor="w")
        body_box = ctk.CTkTextbox(content, font=("Segoe UI", 12))
        body_box.pack(fill="both", expand=True)
        body_box.insert("end", body)
        body_box.configure(state="disabled")

        ctk.CTkButton(win, text="Close", command=win.destroy).pack(pady=10)

    def _log(self, message: str, tag="info"):
        self._log_box.configure(state="normal")
        if tag == "sent" or "[SENT]" in message:
            color = "#4ADE80"
            tag_name = "sent"
        elif tag == "error" or "[ERROR]" in message:
            color = "#F87171"
            tag_name = "error"
        elif tag == "skip" or "[SKIP]" in message:
            color = "#FBBF24"
            tag_name = "skip"
        elif tag == "warn" or "[WARN]" in message:
            color = "#FB923C"
            tag_name = "warn"
        else:
            color = None
            tag_name = "info"
            
        if color:
            self._log_box.tag_config(tag_name, foreground=color)
            self._log_box.insert("end", message + "\n", tag_name)
        else:
            self._log_box.insert("end", message + "\n")
            
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _log_safe(self, msg):
        self.after(0, self._log, msg)

    def _check_data(self):
        recs = self._get_recipients()
        cols = self._active_template.all_cols

        win = ctk.CTkToplevel(self)
        win.title("Check Data")
        win.geometry("720x480")
        win.grab_set()

        hdr = ctk.CTkFrame(win, corner_radius=0)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text="Recipient Preview", font=F_BADGE).pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(hdr, text=f"{len(recs)} row(s)", font=F_BADGE).pack(side="right", padx=20)

        frame = ctk.CTkFrame(win)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        tree = ttk.Treeview(frame, columns=cols, show="headings", style="Custom.Treeview")
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

        ctk.CTkButton(win, text="Close", command=win.destroy).pack(pady=10)

    def _start_send(self):
        if self._sending:
            messagebox.showwarning("Busy", "Already sending. Please wait.")
            return
        recs = self._get_recipients()
        if not recs:
            messagebox.showwarning("No Recipients", "No recipient rows found.\n\nTip: Use 'Check Data' button to verify your data is parsed correctly.")
            return

        t = self._active_template
        if t.has_attachment and not self._attachment_folder:
            if not messagebox.askyesno("No folder", "No attachment folder selected.\nContinue anyway (no attachments)?"):
                return
                
        msg = f"Send {len(recs)} email(s) using template:\n'{t.name}'?\n\n"
            
        if not messagebox.askyesno("Confirm Send", msg):
            return

        self._sending = True
        self._nb.set("Log")
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
                messagebox.showerror("Outlook Error", "Could not connect to Outlook.\n\n1. Make sure Outlook is installed and open.\n2. Note: The 'New Outlook' (Modern App) does not support COM automation. You must use the classic Outlook Desktop App.")
        self.after(0, _show)

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

if __name__ == "__main__":
    app = BulkMailApp()
    app.mainloop()
