# email_handler.py
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from typing import Optional

def send_email(
    to_email: str,
    subject: str,
    body: str,
    smtp_server: str,
    smtp_port: int,
    sender_email: str,
    attachment_path: Optional[str] = None
) -> None:
    """
    Send an email with optional attachment.
    Args:
        to_email: Recipient email address.
        subject: Email subject.
        body: HTML email body.
        smtp_server: SMTP server address.
        smtp_port: SMTP server port.
        sender_email: Sender's email address.
        attachment_path: Path to the attachment file.
    """
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))
    
    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(attachment_path)}"
            )
            msg.attach(part)
    
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.sendmail(sender_email, to_email, msg.as_string())