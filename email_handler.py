import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from typing import Optional
from utils.logger import setup_logger
from config import EmailConfig

logger = setup_logger(__name__)

class EmailError(Exception):
    """Custom exception for email-related errors."""
    pass

def send_email(
    to_email: str,
    subject: str,
    body: str,
    attachment_path: Optional[Path] = None,
    email_config: EmailConfig = None
) -> None:
    """
    Send an email with optional attachment.
    Args:
        to_email: Recipient email address.
        subject: Email subject.
        body: HTML email body.
        attachment_path: Path to the attachment file.
        email_config: EmailConfig object containing SMTP server details.
    Raises:
        EmailError: If there's an error sending the email
    """
    try:
        smtp_server = email_config.smtp_server
        smtp_port = email_config.smtp_port
        sender_email = email_config.sender_email
        username = email_config.sender_email
        password = os.getenv("EMAIL_PASSWORD")

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))
        
        if attachment_path:
            if not attachment_path.exists():
                raise EmailError(f"Attachment not found: {attachment_path}")
                
            with open(attachment_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=attachment_path.name)
                part["Content-Disposition"] = f'attachment; filename="{attachment_path.name}"'
                msg.attach(part)
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            # Skip authentication if the server does not support it
            if username and password:
                try:
                    server.login(username, password)
                except smtplib.SMTPNotSupportedError:
                    logger.warning("SMTP AUTH extension not supported by server, skipping authentication")
            server.sendmail(sender_email, to_email, msg.as_string())
            logger.info(f"Email sent successfully to {to_email}")
            
    except Exception as e:
        logger.error(f"Failed to send email to {to_email}: {str(e)}")
        raise EmailError(f"Failed to send email: {str(e)}")