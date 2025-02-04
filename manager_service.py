import pandas as pd
from typing import Dict, Any
from email_handler import send_email
from email_composer import create_email_body
from config import EMAIL_CONFIG, POWER_BI_LINK
import os
from pivot_table_generator import generate_manager_report
import matplotlib.pyplot as plt
import io
import base64

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
    pivot_df = pd.read_excel(output_file, sheet_name="Summary")
    
    # Format numerical columns with commas
    for col in ["LTM_Gross_Sales", "Opp_to_Floor", "KVI_Count", "Row_Count"]:
        pivot_df[col] = pivot_df[col].apply(lambda x: f"{x:,.0f}")
    
    pivot_html = pivot_df.to_html(index=False, classes="pivot-table")
    pivot_html = f"""
    <style>
        .pivot-table th, .pivot-table td {{ text-align: right; }}
        .pivot-table th:first-child, .pivot-table td:first-child {{ text-align: left; }}
        .pivot-table th:nth-child(2), .pivot-table td:nth-child(2) {{ text-align: left; }}
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
    
    # Generate pivot table screenshot
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.axis('tight')
    ax.axis('off')
    ax.table(cellText=pivot_df.values, colLabels=pivot_df.columns, cellLoc='center', loc='center')
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    img_str = base64.b64encode(buf.read()).decode('utf-8')
    buf.close()
    
    # Embed the image in the email body
    email_body += f'<img src="data:image/png;base64,{img_str}" alt="Pivot Table"/>'
    
    # Send email to test email
    send_email(
        to_email=EMAIL_CONFIG["test_email"],  # Use test email
        subject=f"{manager_name}: Manager Report {month_year}",
        body=email_body,
        attachment_path=output_file,
        email_config=EMAIL_CONFIG
    )