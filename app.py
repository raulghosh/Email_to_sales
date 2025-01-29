import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import datetime as dt
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Configuration
main_folder = r'''C:\Users\rghosh\OneDrive - Veritiv Corp\AI Lab\Attic and Basement\Temp_Files'''
input_file_name = "Sales_Report_Temp.xlsx"  # exported file name
output_folder = os.path.join(main_folder, "Filtered_Reports")  # Folder to save filtered files
input_file = os.path.join(main_folder, input_file_name)
email_config = {
    "smtp_server": "relay.int.distco.com",
    "smtp_port": 25,  # Port 25 is commonly used for relay servers
    "sender_email": os.getenv("EMAIL_USER"),
    "test_email": "rghosh@veritivcorp.com"
}

# Get the current month and year
current_date = dt.datetime.now()
month_year = current_date.strftime("%b, %Y")

# Load the exported file
data = pd.read_excel(input_file)

# Filter out rows where Sales Rep Email or Manager Email is NaN
data = data.dropna(subset=["Rep Email", "Manager Email"])

# Get unique sales rep emails and names, limited to the first 3
sales_reps = data[["Rep Email", "Rep Name"]].drop_duplicates().head(3).set_index("Rep Email")["Rep Name"].to_dict()

# Get unique managers emails and names, limited to the first 3
managers = data[["Manager Email", "Manager Name"]].drop_duplicates().head(3).set_index("Manager Email")["Manager Name"].to_dict()

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Function to send email
def send_email(to_email, subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg["From"] = email_config["sender_email"]
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "html"))

    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
            msg.attach(part)

    with smtplib.SMTP(email_config["smtp_server"], email_config["smtp_port"]) as server:
        server.sendmail(email_config["sender_email"], to_email, msg.as_string())

# Process and send emails to sales reps
for email, name in sales_reps.items():
    # Filter data for the current sales rep
    filtered_data = data[data["Rep Email"] == email]

    # Remove the SalesRepEmail column
    filtered_data = filtered_data.drop(columns=["Rep Email", "Rep Name", "Manager Email", "Manager Name"])

    # Save the filtered data to a new file
    output_file = os.path.join(output_folder, f"{name}_Report.xlsx")
    filtered_data.to_excel(output_file, index=False)

    # Generate custom email body
    basement_count = filtered_data[filtered_data["Region"] == "Basement"].shape[0]
    attic_count = filtered_data[filtered_data["Region"] == "Attic"].shape[0]
    KVI_count = filtered_data[(filtered_data["KVI Type"] == "2: KVI") | (filtered_data["KVI Type"] == "3: Super KVI")].shape[0]
    basement_sales = round(filtered_data[filtered_data["Region"] == "Basement"]["LTM Gross Sales"].sum(), 0)
    attic_sales = round(filtered_data[filtered_data["Region"] == "Attic"]["LTM Gross Sales"].sum(), 0)
    opp_to_floor = round(filtered_data["Opp to Floor"].sum(), 0)

    # Format numbers with comma separation and round to zero decimals
    basement_sales_formatted = f"{basement_sales:,.0f}"
    attic_sales_formatted = f"{attic_sales:,.0f}"
    opp_to_floor_formatted = f"{opp_to_floor:,.0f}"

    # Create summary table by Region
    summary_table = filtered_data.groupby("Region").agg({
        "LTM Gross Sales": "sum",
        "Region": "count",
        "Opp to Floor": "sum"
    }).rename(columns={"Region": "#Rows count"}).reset_index()

    # Format the numeric columns with comma separation and round to zero
    summary_table["LTM Gross Sales"] = summary_table["LTM Gross Sales"].apply(lambda x: f"{x:,.0f}")
    summary_table["Opp to Floor"] = summary_table["Opp to Floor"].apply(lambda x: f"{x:,.0f}")

    # Sort the summary table by Region in descending order
    summary_table = summary_table.sort_values(by="Region", ascending=False)

    # Convert summary table to HTML with right-aligned numbers
    summary_html = summary_table.to_html(index=False, classes='summary-table')

    # Add CSS for right-aligning numbers in the table
    summary_html = f"""
    <style>
        .summary-table th, .summary-table td {{
            text-align: right;
        }}
    </style>
    {summary_html}
    """

    body = f"""
    <div style="text-align: left;">
        <p>Hi {name},</p>
        <p>Hope you are doing good. Attached is the Attic and Basement Report for {month_year}.</p>
        <p>This month you have {basement_count} action items in 'Basement' corresponding to ${basement_sales_formatted} of gross sales and you have {attic_count} action items in 'Attic' corresponding to {attic_sales_formatted} of gross sales.</p>
        <p>Raising the items in basement to the recommended margin will result in ${opp_to_floor_formatted} of commission profit gain.</p>
        <p>Summary by Region:</p>
        {summary_html}
        <p>Thanks,<br>Pricing Team</p>
    </div>
    """

    # Send email with attachment
    send_email(
        to_email=email_config["test_email"],  # test email
        subject=f"{name}: Attic and Basement Report {month_year}",
        body=body,
        attachment_path=output_file,
    )

    # Delete the temporary filtered file
    os.remove(output_file)

# Generate and send summary email to managers
for manager_email, manager_name in managers.items():
    manager_data = data[data["Manager Email"] == manager_email]
    summary_by_rep = manager_data.groupby(["Rep Name", "Region"]).agg({
    "LTM Gross Sales": "sum",
    "Opp to Floor": "sum",
    "Rep Name": "count",
    "KVI Type": lambda x: (x == "2: KVI").sum() + (x == "3: Super KVI").sum()
}).rename(columns={
    "Rep Name": "#Rows",
    "KVI Type": "#Key Item"
}).reset_index()

# Format the numeric columns with comma separation and round to zero
summary_by_rep["LTM Gross Sales"] = summary_by_rep["LTM Gross Sales"].apply(lambda x: f"{x:,.0f}")
summary_by_rep["Opp to Floor"] = summary_by_rep["Opp to Floor"].apply(lambda x: f"{x:,.0f}")

# Sort the summary table by Rep Name (ascending), Region (descending), and LTM Gross Sales (descending)
summary_by_rep = summary_by_rep.sort_values(by=["Rep Name", "Region", "LTM Gross Sales"], ascending=[True, False, False])

# Convert summary table to HTML with right-aligned numbers
summary_by_rep_html = summary_by_rep.to_html(index=False, classes='summary-table')

# Add CSS for right-aligning numbers in the table
summary_by_rep_html = f"""
<style>
    .summary-table th, .summary-table td {{
        text-align: right;
    }}
</style>
{summary_by_rep_html}
"""

manager_body = f"""
<div style="text-align: left;">
    <p>Hi {manager_name},</p>
    <p>Attached is the summary report for your sales reps for {month_year}.</p>
    <p>Summary by Sales Rep:</p>
    {summary_by_rep_html}
    <p>Thanks,<br>Pricing Team</p>
</div>
"""

# Send summary email to manager
send_email(
    to_email=email_config["test_email"],  # test email
    subject=f"Summary Report for Your Sales Reps - {month_year}",
    body=manager_body
)