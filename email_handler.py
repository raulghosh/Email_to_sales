# email_handler.py
import smtplib
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
    """Send an email with optional attachment."""
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

def create_email_body(
    name: str,
    month_year: str,
    basement_count: int,
    attic_count: int,
    basement_sales: float,
    attic_sales: float,
    opp_to_floor: float,
    summary_html: str,
    power_bi_link: str
) -> str:
    """Generate the HTML email body."""
    return f"""
    <div style="text-align: left;">
        <p>Hi {name},</p>
        <p>Hope you are doing good. Attached is the Attic and Basement Report for {month_year}.</p>
        <p>This month you have {basement_count} action items in 'Basement' corresponding to ${basement_sales:,.0f} of gross sales and you have {attic_count} action items in 'Attic' corresponding to {attic_sales:,.0f} of gross sales.</p>
        <p>Raising the items in basement to the recommended margin will result in ${opp_to_floor:,.0f} of commission profit gain.</p>
        <p>Summary by Region:</p>
        {summary_html}
        <p>Access the live Power BI Dashboard: <a href="{power_bi_link}">Attic and Basement Report</a></p>
        <p>Thanks,<br>Pricing Team</p>
    </div>
    """