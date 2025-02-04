import pandas as pd
from typing import Dict, Any
from excel_formatter import format_excel_sheet
from email_handler import send_email
from email_composer import create_email_body
from config import EMAIL_CONFIG, POWER_BI_LINK
import os

def generate_sales_rep_report(
    data: pd.DataFrame,
    email: str,
    name: str,
    output_folder: str,
    month_year: str
) -> str:
    """Generate a sales rep report and return the output file path."""
    # Filter data for the rep
    filtered_raw = data[data["Rep Email"] == email]
    filtered_formatted = data[data["Rep Email"] == email].drop(columns=["Rep Email", "Rep Name", "Manager Email", "Manager Name"])
    
    # Reorder columns (Opp to Floor next to LTM Gross Sales)
    cols = list(filtered_formatted.columns)
    ltm_index = cols.index("LTM Gross Sales")
    cols.insert(ltm_index + 1, cols.pop(cols.index("Opp to Floor")))
    filtered_formatted = filtered_formatted[cols]
    
    # Save to Excel
    output_file = os.path.join(output_folder, f"{name}_Report.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        filtered_formatted.to_excel(writer, index=False, sheet_name="Sales Report")
        worksheet = writer.sheets["Sales Report"]
        format_excel_sheet(worksheet, filtered_formatted)
    
    return output_file

def send_sales_rep_email(
    data: pd.DataFrame,
    email: str,
    name: str,
    output_folder: str,
    month_year: str
) -> None:
    """Process and send an email to a sales rep."""
    # Generate report
    output_file = generate_sales_rep_report(data, email, name, output_folder, month_year)
    
    # Calculate metrics
    filtered_raw = data[data["Rep Email"] == email]
    basement_count = filtered_raw[filtered_raw["Region"] == "Basement"].shape[0]
    attic_count = filtered_raw[filtered_raw["Region"] == "Attic"].shape[0]
    basement_sales = filtered_raw[filtered_raw["Region"] == "Basement"]["LTM Gross Sales"].sum()
    attic_sales = filtered_raw[filtered_raw["Region"] == "Attic"]["LTM Gross Sales"].sum()
    opp_to_floor = filtered_raw["Opp to Floor"].sum()
    
    # Generate summary table HTML
    summary_table = filtered_raw.groupby("Region").agg({
        "LTM Gross Sales": "sum",
        "Opp to Floor": "sum",
        "Region": "count",
        "KVI Type": lambda x: ((x == "2: KVI") | (x == "3: Super KVI")).sum()
    }).rename(columns={"Region": "Item Count", "KVI Type": "KVI Items"}).reset_index()
    
    summary_table["LTM Gross Sales"] = summary_table["LTM Gross Sales"].apply(lambda x: f"{x:,.0f}")
    summary_table["Opp to Floor"] = summary_table["Opp to Floor"].apply(lambda x: f"{x:,.0f}")
    summary_table["Item Count"] = summary_table["Item Count"].apply(lambda x: f"{x:,.0f}")
    summary_table["KVI Items"] = summary_table["KVI Items"].apply(lambda x: f"{x:,.0f}")
    
    summary_html = summary_table.to_html(index=False, classes="summary-table")
    summary_html = f"""
    <style>
        .summary-table th, .summary-table td {{ text-align: right; }}
        .summary-table th:first-child, .summary-table td:first-child {{ text-align: left; }}
    </style>
    {summary_html}
    """
    
    # Create email body
    email_body = create_email_body(
        recipient_type="sales_rep",
        name=name,
        month_year=month_year,
        power_bi_link=POWER_BI_LINK,
        sales_rep_data={
            "basement_count": basement_count,
            "attic_count": attic_count,
            "basement_sales": basement_sales,
            "attic_sales": attic_sales,
            "opp_to_floor": opp_to_floor,
            "summary_html": summary_html
        }
    )
    
    # Send email
    send_email(
        to_email=EMAIL_CONFIG["test_email"],# replace by email when productionizing
        subject=f"{name}: Attic and Basement Report {month_year}",
        body=email_body,
        smtp_server=EMAIL_CONFIG["smtp_server"],
        smtp_port=EMAIL_CONFIG["smtp_port"],
        sender_email=EMAIL_CONFIG["sender_email"],
        attachment_path=output_file
    )
    print(f"Email sent to {name} at {EMAIL_CONFIG["test_email"]}")