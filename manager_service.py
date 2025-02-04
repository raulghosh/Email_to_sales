import pandas as pd
from typing import Dict, Any
from excel_formatter import format_excel_sheet
from email_handler import send_email
from email_composer import create_email_body
from config import EMAIL_CONFIG, POWER_BI_LINK
import os

def generate_manager_report(
    data: pd.DataFrame,
    manager_name: str,
    output_folder: str,
    month_year: str
) -> str:
    """Generate a pivot table report for a manager and return the output file path."""
    # Create pivot table
    pivot_df = data.pivot_table(
        index=["Rep Name", "Region"],
        values=["LTM Gross Sales", "Opp to Floor"],
        aggfunc={"LTM Gross Sales": "sum", "Opp to Floor": "sum", "Region": "count"}
    ).rename(columns={"Region": "Item Count"}).reset_index()
    
    # Save to Excel
    output_file = os.path.join(output_folder, f"{manager_name}_Manager_Report.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        pivot_df.to_excel(writer, index=False, sheet_name="Pivot Table")
        worksheet = writer.sheets["Pivot Table"]
        format_excel_sheet(worksheet, pivot_df)
    
    return output_file

def send_manager_email(
    data: pd.DataFrame,
    manager_email: str,
    manager_name: str,
    output_folder: str,
    month_year: str
) -> None:
    """Process and send an email to a manager."""
    # Generate report
    output_file = generate_manager_report(data, manager_name, output_folder, month_year)
    
    # Generate pivot table HTML
    pivot_df = pd.read_excel(output_file, sheet_name="Pivot Table")
    pivot_html = pivot_df.to_html(index=False, classes="pivot-table")
    pivot_html = f"""
    <style>
        .pivot-table th, .pivot-table td {{ text-align: right; }}
        .pivot-table th:first-child, .pivot-table td:first-child {{ text-align: left; }}
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
        to_email=EMAIL_CONFIG["test_email"],# replace by manager_email when productionizing
        subject=f"{manager_name}: Manager Report {month_year}",
        body=email_body,
        smtp_server=EMAIL_CONFIG["smtp_server"],
        smtp_port=EMAIL_CONFIG["smtp_port"],
        sender_email=EMAIL_CONFIG["sender_email"],
        attachment_path=output_file
    )
    print(f"Email sent to manager {manager_name} at {EMAIL_CONFIG["test_email"]}")