import re, os
with open("Gmail-Tools/main.py", "r", encoding="utf-8") as f:
    text = f.read()

new_send_emails = """def send_emails(template, cfg, recipients, attachment_folder, log_cb, done_cb):
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
                "<html><body><pre style=\'font-family:Calibri,sans-serif;font-size:11pt;white-space:pre-wrap\'>"
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
                        part["Content-Disposition"] = f\'attachment; filename="{os.path.basename(path)}"\'
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
"""

text = re.sub(r"def send_emails\(template, cfg.+?done_cb\(success=(True|False)\)", new_send_emails, text, flags=re.DOTALL)

with open("Gmail-Tools/main.py", "w", encoding="utf-8") as f:
    f.write(text)