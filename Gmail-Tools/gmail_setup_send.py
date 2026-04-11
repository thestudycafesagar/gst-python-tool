with open(r"c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\Gmail-Tools\main.py", "r", encoding="utf-8") as f:
    text = f.read()

start_idx = text.find("def send_emails(template")
end_idx = text.find("\n# \u2550", start_idx) 

old_send = text[start_idx:end_idx]

# using regular ascii
new_send = """def send_emails(template, cfg, recipients, attachment_folder, log_cb, done_cb):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    import os

    sender_email = str(cfg.get("sender_email", "")).strip()
    app_password = str(cfg.get("app_password", "")).strip()

    if not sender_email or not app_password:
        log_cb("[ERROR] Sender Gmail or App Password not configured.")
        done_cb(success=False)
        return

    try:
        # Establish connection once for the entire batch
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
                        # Attach and keep original format
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path)}"'
                        msg.attach(part)
                    else:
                        log_cb(f"[WARN]  {name} - attachment not found: {base}.pdf")

            all_recipients = [email]
            if cc_email:
                all_recipients.extend([rcpt.strip() for rcpt in cc_email.split(',') if rcpt.strip()])

            # Send via SMTP
            server.send_message(msg, from_addr=sender_email, to_addrs=all_recipients)
            log_cb(f"[SENT]  {name} <{email}>")
            
        except Exception as exc:
            log_cb(f"[ERROR] {name} <{email}> - {exc}")
            failed.append(email)

    try:
        server.quit()
    except:
        pass
        
    done_cb(success=True, failed=failed)
"""

if len(old_send) > 200:
    text = text.replace(old_send, new_send)
    with open(r"c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\Gmail-Tools\main.py", "w", encoding="utf-8") as f:
        f.write(text)
    print("SEND REPLACED")
else:
    print(f"NOT FOUND OLD SEND {start_idx} {end_idx}")
