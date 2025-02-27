from pathlib import Path
import pandas as pd
from typing import Dict, Any
from email_handler import send_email, EmailError
from email_composer import create_email_body
from pivot_table_generator import generate_manager_report, PivotTableError, generate_html_table, _create_summary_table
from sales_rep_service import generate_sales_rep_report
from config import CONFIG
from utils.logger import setup_logger

logger = setup_logger(__name__)

class ManagerServiceError(Exception):
    """Custom exception for manager service errors."""
    pass

def generate_manager_pivot_html(data: pd.DataFrame) -> str:
    """
    Generate two HTML pivot tables for the manager email: 
    one for 'Basement' (sorted by 'Opp to Floor'), 
    and one for 'Attic' (sorted by 'LTM Gross Sales').
    """
    try:
        # Define aggregation
        agg_funcs = {
            "LTM Gross Sales": "sum",
            "Opp to Floor": "sum",
            "Item Visibility": lambda x: ((x == "Medium") | (x == "High")).sum(),
            "Sales Rep Name": "count"
        }

        # Process "Basement" table
        basement_df = (
            data[data["Category"] == "Basement"]
            .groupby(["Sales Rep Name"])
            .agg(agg_funcs)
            .rename(columns={"Sales Rep Name": "# Rows", "Item Visibility": "# Visible Items"})
            .reset_index()
            .sort_values(by="Opp to Floor", ascending=False)  # Sort by 'Opp to Floor'
        )

        # Process "Attic" table (Remove "Opp to Floor" column)
        attic_df = (
            data[data["Category"] == "Attic"]
            .groupby(["Sales Rep Name"])
            .agg(agg_funcs)
            .rename(columns={"Sales Rep Name": "# Rows", "Item Visibility": "# Visible Items"})
            .reset_index()
            .drop(columns=["Opp to Floor"])  # Remove Opp to Floor
            .sort_values(by="LTM Gross Sales", ascending=False)  # Sort by 'LTM Gross Sales'
        )

        # Format numerical columns
        for df in [basement_df, attic_df]:
            for col in ["LTM Gross Sales", "# Visible Items"]:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: f"{x:,.0f}")

        # Convert to HTML (Align Sales Rep Name & Category to left)
        def df_to_html(df):
            html = df.to_html(index=False, classes="pivot-table", escape=False)
            html = html.replace("<th>Sales Rep Name</th>", '<th style="text-align: left;">Sales Rep Name</th>')

            # Apply right alignment to all columns except the first
            html = html.replace("<td>", '<td style="text-align: right;">')

            # Left align the first column
            first_col_start = html.find('<td style="text-align: right;">')
            if first_col_start != -1:
                html = html[:first_col_start] + html[first_col_start:].replace('<td style="text-align: right;">', '<td style="text-align: left;">', 1)

            return html

        basement_html = df_to_html(basement_df)
        attic_html = df_to_html(attic_df)

        # Combine with styling
        pivot_html = f"""
        <style>
            .pivot-table th, .pivot-table td {{ text-align: right; padding: 5px; }}
            .pivot-table th:first-child, .pivot-table td:first-child {{ text-align: left; }}
        </style>
        {basement_html}
        <br>
        {attic_html}
        """
        return pivot_html

    except Exception as e:
        logger.error(f"Failed to generate pivot tables: {str(e)}")
        raise ManagerServiceError(f"Failed to generate pivot tables: {str(e)}")

def send_manager_email(
    data: pd.DataFrame,
    manager_email: str,
    manager_name: str,
    output_folder: Path,
    month_year: str
) -> None:
    try:
        if data.empty:
            raise ManagerServiceError(f"No data provided for manager: {manager_name}")

        logger.info(f"Processing manager email for: {manager_name}")

        # Generate and save report
        output_file = generate_manager_report(data, manager_name, output_folder, month_year)

        # Use aggregated summaries instead of raw data
        attic_summary = _create_summary_table(data[data["Category"] == "Attic"], category="Attic")
        basement_summary = _create_summary_table(data[data["Category"] == "Basement"], category="Basement")

        # Generate HTML tables using the aggregated data
        basement_html = generate_html_table(basement_summary, title="Basement Summary:")
        attic_html = generate_html_table(attic_summary, title="Attic Summary:")

        # Create email body
        email_body = create_email_body(
            recipient_type="manager",
            name=manager_name,
            month_year=month_year,
            power_bi_link=CONFIG.power_bi_link,
            basement_html=basement_html,
            attic_html=attic_html
        )

        # Send email
        send_email(
            to_email=CONFIG.email_config.test_email,
            subject=f"{manager_name}: Manager Report {month_year}",
            body=email_body,
            attachment_path=output_file,
            email_config=CONFIG.email_config
        )

        logger.info(f"Successfully sent email to manager: {manager_name}")

    except (PivotTableError, EmailError) as e:
        logger.error(f"Failed to process manager {manager_name}: {str(e)}")
        raise ManagerServiceError(f"Failed to process manager: {str(e)}")
    except Exception as e:
        logger.error(f"Unexpected error processing manager {manager_name}: {str(e)}")
        raise ManagerServiceError(f"Unexpected error: {str(e)}")
