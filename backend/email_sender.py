"""
Email Sender - Sends follow-up emails via SMTP with Gmail support.
"""

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime


def send_followup_email(
    to_email: str,
    subject: str,
    body: str,
    cc_email: str = "",
    sender_name: str = "",
) -> tuple[bool, str]:
    smtp_host = os.getenv("EMAIL_SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("EMAIL_SMTP_PORT", 587))
    sender_email = os.getenv("EMAIL_SENDER", "")
    password = os.getenv("EMAIL_PASSWORD", "")

    if not sender_email or not password:
        return False, "EMAIL_SENDER or EMAIL_PASSWORD not configured in .env"

    if not to_email or "@" not in to_email:
        return False, f"Invalid recipient email: {to_email}"

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
        msg["To"] = to_email
        if cc_email and "@" in cc_email:
            msg["Cc"] = cc_email

        msg.attach(MIMEText(body, "plain"))

        # Build HTML version
        html_body = body.replace("\n", "<br>")
        html = f"""
        <html><body style="font-family: Arial, sans-serif; color: #222; max-width: 600px; margin: auto;">
        <p>{html_body}</p>
        <hr style="border: none; border-top: 1px solid #eee; margin: 24px 0;">
        <p style="font-size: 11px; color: #888;">Sent via Auto Follow-Up Tool</p>
        </body></html>
        """
        msg.attach(MIMEText(html, "html"))

        recipients = [to_email]
        if cc_email and "@" in cc_email:
            recipients.append(cc_email)

        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, recipients, msg.as_string())

        return True, f"Email sent to {to_email} at {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    except smtplib.SMTPAuthenticationError:
        return False, "SMTP authentication failed — check EMAIL_SENDER and EMAIL_PASSWORD in .env"
    except smtplib.SMTPRecipientsRefused:
        return False, f"Recipient {to_email} refused by SMTP server"
    except Exception as e:
        return False, f"Email sending failed: {str(e)}"


def parse_email_draft(draft: str) -> tuple[str, str]:
    """Split a draft into (subject, body). Subject line must start with 'Subject: '."""
    lines = draft.strip().split("\n")
    subject = ""
    body_lines = []
    found_subject = False

    for i, line in enumerate(lines):
        if line.startswith("Subject:") and not found_subject:
            subject = line.replace("Subject:", "").strip()
            found_subject = True
        else:
            body_lines.append(line)

    body = "\n".join(body_lines).strip()
    if not subject:
        subject = "Following up on my application"

    return subject, body
