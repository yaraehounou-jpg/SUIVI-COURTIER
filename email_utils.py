# SUIVI_AUTO/email_utils.py
from __future__ import annotations

import smtplib
from email.message import EmailMessage


def send_pdf_email_smtp(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_pass: str,
    mail_from: str,
    to_email: str,
    subject: str,
    body: str,
    pdf_bytes: bytes,
    pdf_filename: str,
) -> None:
    msg = EmailMessage()
    msg["From"] = mail_from
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_filename,
    )

    with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
