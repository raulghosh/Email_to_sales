import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from typing import Optional, Dict, Any

def send_email(
    to_email: str,
    subject: str,
    body: str,
    attachment_path: Optional[str] = None,
    email_config: Dict[str, Any] = None
) -> None:
    """
    Send an email with optional attachment.
    Args:
        to_email: Recipient email address.
        subject: Email subject.
        body: HTML email body.
        attachment_path: Path to the attachment file.
        email_config: Dictionary containing SMTP server details.
    """
    smtp_server = email_config["smtp_server"]
    smtp_port = email_config["smtp_port"]
    sender_email = email_config["sender_email"]

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