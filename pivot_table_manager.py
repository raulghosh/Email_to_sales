from email_composer import create_email_body
from config import POWER_BI_LINK, EMAIL_CONFIG
import pandas as pd
from email_handler import send_email

def send_manager_email(manager_name: str, manager_email: str, output_file: str, month_year: str) -> None:
    """Send an email to the sales manager with the pivot table report."""
    # Generate pivot table HTML (from the Excel file)
    pivot_df = pd.read_excel(output_file, sheet_name="Pivot Table")
    pivot_html = pivot_df.to_html(index=False, classes="pivot-table")
    
    # Add CSS styling
    pivot_html = f"""
    <style>
        .pivot-table {{
            border-collapse: collapse;
            width: 100%;
        }}
        .pivot-table th, .pivot-table td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: right;
        }}
        .pivot-table th {{
            background-color: #006400;
            color: white;
            text-align: left;
        }}
    </style>
    {pivot_html}
    """
    
    # Create email body
    email_body = create_email_body(
        recipient_type="manager",
        name=manager_name,
        month_year=month_year,
        power_bi_link=POWER_BI_LINK,
        pivot_html=pivot_html
    )
    
    # Send email
    send_email(
        to_email=manager_email,
        subject=f"{manager_name}: Sales Manager Report {month_year}",
        body=email_body,
        smtp_server=EMAIL_CONFIG["smtp_server"],
        smtp_port=EMAIL_CONFIG["smtp_port"],
        sender_email=EMAIL_CONFIG["sender_email"],
        attachment_path=output_file
    )