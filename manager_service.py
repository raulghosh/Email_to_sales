from pathlib import Path
import pandas as pd
from typing import Dict, Any
from email_handler import send_email, EmailError
from email_composer import create_email_body
from pivot_table_generator import generate_manager_report, PivotTableError
from config import CONFIG
from utils.logger import setup_logger

logger = setup_logger(__name__)

class ManagerServiceError(Exception):
    """Custom exception for manager service errors."""
    pass

def generate_manager_pivot_html(data: pd.DataFrame) -> str:
    """
    Generate two HTML pivot tables for manager email: 
    one for 'Basement' sorted by 'Opp to Floor', 
    and one for 'Attic' sorted by 'LTM Gross Sales'.
    """
    try:
        # Define aggregation
        agg_funcs = {
            "LTM Gross Sales": "sum",
            "Opp to Floor": "sum",
            "KVI Type": lambda x: ((x == "2: KVI") | (x == "3: Super KVI")).sum(),
            "Rep Name": "count"
        }

        # Process "Basement" table
        basement_df = (
            data[data["Region"] == "Basement"]
            .groupby(["Rep Name", "Region"])
            .agg(agg_funcs)
            .rename(columns={"Rep Name": "Row Count", "KVI Type": "Item Visibility"})
            .reset_index()
            .sort_values(by="Opp to Floor", ascending=False)  # Sort by 'Opp to Floor'
        )

        # Process "Attic" table
        attic_df = (
            data[data["Region"] == "Attic"]
            .groupby(["Rep Name", "Region"])
            .agg(agg_funcs)
            .rename(columns={"Rep Name": "Row Count", "KVI Type": "Item Visibility"})
            .reset_index()
            .sort_values(by="LTM Gross Sales", ascending=False)  # Sort by 'LTM Gross Sales'
        )

        # Format numerical columns
        for df in [basement_df, attic_df]:
            for col in ["LTM Gross Sales", "Opp to Floor", "Row Count", "Item Visibility"]:
                df[col] = df[col].apply(lambda x: f"{x:,.0f}")

        # Convert to HTML
        basement_html = basement_df.to_html(index=False, classes="pivot-table")
        attic_html = attic_df.to_html(index=False, classes="pivot-table")

        # Combine with styling
        pivot_html = f"""
        <style>
            .pivot-table th, .pivot-table td {{ text-align: right; padding: 5px; }}
            .pivot-table th:first-child, .pivot-table td:first-child,
            .pivot-table th:nth-child(2), .pivot-table td:nth-child(2) {{ text-align: left; }}
        </style>
        <h3>Basement Summary (Sorted by 'Opp to Floor')</h3>
        {basement_html}
        <br>
        <h3>Attic Summary (Sorted by 'LTM Gross Sales')</h3>
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
    output_folder: Path | str,
    month_year: str
) -> None:
    """
    Process and send an email to a manager.
    
    Args:
        data: Input DataFrame
        manager_email: Manager's email
        manager_name: Manager's name
        output_folder: Output directory path
        month_year: Month and year for the report
        
    Raises:
        ManagerServiceError: If there's an error in the process
    """
    try:
        # Validate input data
        if data.empty:
            raise ManagerServiceError(f"No data provided for manager: {manager_name}")
            
        logger.info(f"Processing manager email for: {manager_name}")
        
        # Generate manager report
        output_file = generate_manager_report(data, manager_name, output_folder, month_year)
        
        # Generate pivot tables for both Basement & Attic
        pivot_html = generate_manager_pivot_html(data)

        # Create email body with the new tables
        email_body = create_email_body(
            recipient_type="manager",
            name=manager_name,
            month_year=month_year,
            power_bi_link=CONFIG.power_bi_link,
            pivot_html=pivot_html  # Ensure the email template accepts this
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