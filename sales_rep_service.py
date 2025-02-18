from pathlib import Path
import pandas as pd
from typing import Dict, Any
from excel_formatter import format_excel_sheet
from email_handler import send_email, EmailError
from email_composer import create_email_body
from config import CONFIG
from utils.logger import setup_logger
import os

logger = setup_logger(__name__)

class SalesRepServiceError(Exception):
    """Custom exception for sales rep service errors."""
    pass

def generate_sales_rep_report(
    data: pd.DataFrame,
    email: str,
    name: str,
    output_folder: Path | str,
    month_year: str
) -> Path:
    """
    Generate a sales rep report and save it to Excel.
    
    Args:
        data: Input DataFrame
        email: Sales rep email
        name: Sales rep name
        output_folder: Output directory path
        month_year: Month and year for the report
        
    Returns:
        Path to the generated report
        
    Raises:
        SalesRepServiceError: If there's an error generating the report
    """
    try:
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        
        # Filter and format data
        filtered_raw = data[data["Rep Email"] == email]
        if filtered_raw.empty:
            raise SalesRepServiceError(f"No data found for sales rep: {name}")
            
        filtered_formatted = _prepare_report_data(filtered_raw)
        
        # Save to Excel
        output_file = output_folder / f"{name}_Report.xlsx"
        logger.info(f"Generating sales rep report: {output_file}")
        
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            filtered_formatted.to_excel(writer, index=False, sheet_name="Sales Report")
            worksheet = writer.sheets["Sales Report"]
            format_excel_sheet(worksheet, filtered_formatted)
            
        return output_file
        
    except Exception as e:
        logger.error(f"Failed to generate sales rep report for {name}: {str(e)}")
        raise SalesRepServiceError(f"Failed to generate report: {str(e)}")

def _prepare_report_data(data: pd.DataFrame) -> pd.DataFrame:
    """
    Prepare and format data for the sales rep report.
    
    Args:
        data: Input DataFrame
        
    Returns:
        Formatted DataFrame
    """
    # Drop unnecessary columns and reorder
    formatted = data.drop(columns=["Rep Email", "Rep Name", "Manager Email", "Manager Name"])
    
    # Reorder columns (Opp to Floor next to LTM Gross Sales)
    cols = list(formatted.columns)
    ltm_index = cols.index("LTM Gross Sales")
    cols.insert(ltm_index + 1, cols.pop(cols.index("Opp to Floor")))
    
    return formatted[cols]

def calculate_metrics(data: pd.DataFrame) -> Dict[str, Any]:
    """
    Calculate metrics for the sales rep report.
    
    Args:
        data: Input DataFrame
        
    Returns:
        Dictionary containing calculated metrics
    """
    metrics = {
        "basement_count": int(data[data["Region"] == "Basement"].shape[0]),
        "attic_count": int(data[data["Region"] == "Attic"].shape[0]),
        "basement_sales": float(data[data["Region"] == "Basement"]["LTM Gross Sales"].sum()),
        "attic_sales": float(data[data["Region"] == "Attic"]["LTM Gross Sales"].sum()),
        "opp_to_floor": float(data["Opp to Floor"].sum())
    }
    
    # Generate summary table
    summary_table = data.groupby("Region").agg({
        "LTM Gross Sales": "sum",
        "Opp to Floor": "sum",
        "Region": "count",
        "KVI Type": lambda x: ((x == "2: KVI") | (x == "3: Super KVI")).sum()
    }).rename(columns={
        "Region": "Item Count",
        "KVI Type": "KVI Items"
    }).reset_index()
    
    # Format numerical values
    for col in ["LTM Gross Sales", "Opp to Floor", "Item Count", "KVI Items"]:
        summary_table[col] = summary_table[col].apply(lambda x: f"{x:,.0f}")
    
    metrics["summary_html"] = _format_summary_table_html(summary_table)
    return metrics

def _format_summary_table_html(summary_table: pd.DataFrame) -> str:
    """Format summary table as HTML with styling."""
    return f"""
    <style>
        .summary-table th, .summary-table td {{ text-align: right; }}
        .summary-table th:first-child, .summary-table td:first-child {{ text-align: left; }}
    </style>
    {summary_table.to_html(index=False, classes="summary-table")}
    """

def send_sales_rep_email(
    data: pd.DataFrame,
    email: str,
    name: str,
    output_folder: Path | str,
    month_year: str
) -> None:
    """
    Process and send an email to a sales rep.
    
    Args:
        data: Input DataFrame
        email: Sales rep email
        name: Sales rep name
        output_folder: Output directory path
        month_year: Month and year for the report
        
    Raises:
        SalesRepServiceError: If there's an error in the process
    """
    try:
        # Filter data for the rep
        rep_data = data[data["Rep Email"] == email]
        if rep_data.empty:
            raise SalesRepServiceError(f"No data found for sales rep: {name}")
        
        # Generate report
        output_file = generate_sales_rep_report(data, email, name, output_folder, month_year)
        
        # Calculate metrics
        metrics = calculate_metrics(rep_data)
        
        # Create email body
        email_body = create_email_body(
            recipient_type="sales_rep",
            name=name,
            month_year=month_year,
            power_bi_link=CONFIG.power_bi_link,
            sales_rep_data=metrics
        )
        
        # Send email
        send_email(
            to_email=CONFIG.email_config.test_email,
            subject=f"{name}: Sales Report {month_year}",
            body=email_body,
            attachment_path=output_file,
            email_config={
                "smtp_server": CONFIG.email_config.smtp_server,
                "smtp_port": CONFIG.email_config.smtp_port,
                "sender_email": CONFIG.email_config.sender_email,
                "username": CONFIG.email_config.sender_email,
                "password": os.getenv("EMAIL_PASSWORD")
            }
        )
        
        logger.info(f"Successfully sent email to {name}")
        
    except (SalesRepServiceError, EmailError) as e:
        logger.error(f"Failed to process sales rep {name}: {str(e)}")
        raise SalesRepServiceError(f"Failed to process sales rep: {str(e)}")
    except Exception as e:
        logger.error(f"Unexpected error processing sales rep {name}: {str(e)}")
        raise SalesRepServiceError(f"Unexpected error: {str(e)}")