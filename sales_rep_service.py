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
    Generate a sales rep report and save it to Excel with separate sheets for Attic and Basement.
    
    Args:
        data: Input DataFrame
        email: Sales Rep Email
        name: Sales Rep Name
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
        
        # Filter data for the sales rep
        filtered_raw = data[data["Sales Rep Email"] == email]
        if filtered_raw.empty:
            raise SalesRepServiceError(f"No data found for sales rep: {name}")
            
        # Split data into Attic and Basement
        attic_data = filtered_raw[filtered_raw["Category"] == "Attic"]
        basement_data = filtered_raw[filtered_raw["Category"] == "Basement"]
        
        # Format the data
        attic_formatted = _prepare_report_data(attic_data,category="Attic")
        basement_formatted = _prepare_report_data(basement_data,category="Basement")
        
        # Save to Excel
        output_file = output_folder / f"{name}_Report.xlsx"
        logger.info(f"Generating sales rep report: {output_file}")
        
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            
            # Write Basement data to its own sheet
            basement_formatted.to_excel(writer, index=False, sheet_name="Basement")
            basement_worksheet = writer.sheets["Basement"]
            format_excel_sheet(basement_worksheet, basement_formatted)

            # Write Attic data to its own sheet
            attic_formatted.to_excel(writer, index=False, sheet_name="Attic")
            attic_worksheet = writer.sheets["Attic"]
            format_excel_sheet(attic_worksheet, attic_formatted)
        
        return output_file
        
    except Exception as e:
        logger.error(f"Failed to generate sales rep report for {name}: {str(e)}")
        raise SalesRepServiceError(f"Failed to generate report: {str(e)}")

def _prepare_report_data(data: pd.DataFrame, category: str) -> pd.DataFrame:
    """
    Prepare and format data for the sales rep report.
    
    Args:
        data: Input DataFrame
        
    Returns:
        Formatted DataFrame
    """
    # Drop unnecessary columns and reorder
    formatted = data.drop(columns=["Sales Rep Email", "Sales Rep Name", "Manager Email", "Manager Name", "RVP Name", "RVP Email", "VP Name", "VP Email"])
    
    # Convert LTM Gross Sales and Opp to Floor to strings, right-aligned without decimals
    formatted["LTM Gross Sales1"] = pd.to_numeric(formatted["LTM Gross Sales"],errors="coerce")
    formatted["Opp to Floor1"] = pd.to_numeric(formatted["Opp to Floor"],errors="coerce")
    
    # Sort the data
    if category == "Attic":
        formatted = formatted.sort_values(by=["LTM Gross Sales1"], ascending=False)
    else:
        formatted = formatted.sort_values(by=["Opp to Floor1"], ascending=False)
        
    # Drop the temporary columns after sorting
    formatted = formatted.drop(columns=["LTM Gross Sales1", "Opp to Floor1"])
    
    formatted["LTM Gross Sales"] = formatted["LTM Gross Sales"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    formatted["Opp to Floor"] = formatted["Opp to Floor"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    formatted["Opp to Target"] = formatted["Opp to Target"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")

    # Convert Margin columns to strings with one decimal place and percentage sign
    margin_columns = [col for col in formatted.columns if 'margin' in col.lower()]  # Replace with actual margin column names
    for col in margin_columns:
        if col in formatted.columns:
            formatted[col] = formatted[col].apply(lambda x: f"{100*x:.1f}%" if pd.notna(x) else "")
    return formatted

def calculate_metrics(data: pd.DataFrame) -> Dict[str, Any]:
    """
    Calculate metrics for the sales rep report.
    
    Args:
        data: Input DataFrame
        
    Returns:
        Dictionary containing calculated metrics
    """
    metrics = {
        "basement_count": int(data[data["Category"] == "Basement"].shape[0]),
        "attic_count": int(data[data["Category"] == "Attic"].shape[0]),
        "basement_sales": float(data[data["Category"] == "Basement"]["LTM Gross Sales"].sum()),
        "attic_sales": float(data[data["Category"] == "Attic"]["LTM Gross Sales"].sum()),
        "opp_to_floor": float(data["Opp to Floor"].sum())
    }
    
    # Generate summary table
    summary_table = data.groupby("Category").agg({
        "LTM Gross Sales": "sum",
        "Opp to Floor": "sum",
        "Category": "count",
        "Item Visibility": lambda x: ((x == "Medium") | (x == "High")).sum()
    }).rename(columns={
        "Category": "# Rows",
        "Item Visibility": "# Visible Items"
    }).reset_index()
    
    # Format numerical values
    for col in ["LTM Gross Sales", "Opp to Floor", "# Rows", "# Visible Items"]:
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
        email: Sales Rep Email
        name: Sales Rep Name
        output_folder: Output directory path
        month_year: Month and year for the report
        
    Raises:
        SalesRepServiceError: If there's an error in the process
    """
    try:
        # Filter data for the rep
        rep_data = data[data["Sales Rep Email"] == email]
        if rep_data.empty:
            logger.debug(f"Data for sales rep {name} ({email}):\n{data.head()}")
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
            email_config=CONFIG.email_config
        )
        
        logger.info(f"Successfully sent email to {name}")
        
    except (SalesRepServiceError, EmailError) as e:
        logger.error(f"Failed to process sales rep {name}: {str(e)}")
        raise SalesRepServiceError(f"Failed to process sales rep: {str(e)}")
    except Exception as e:
        logger.error(f"Unexpected error processing sales rep {name}: {str(e)}")
        raise SalesRepServiceError(f"Unexpected error: {str(e)}")

def clean_data(data: pd.DataFrame) -> pd.DataFrame:
    """
    Remove rows with missing Rep/Manager emails.
    
    Args:
        data: Input DataFrame
        
    Returns:
        Cleaned DataFrame
    """
    if data.empty:
        logger.warning("Empty DataFrame received for cleaning")
        return data
        
    initial_rows = len(data)
    cleaned_data = data.dropna(subset=["Sales Rep Email", "Manager Email"])
    dropped_rows = initial_rows - len(cleaned_data)
    
    if dropped_rows > 0:
        logger.warning(f"Dropped {dropped_rows} rows with missing emails")
    
    logger.debug(f"Data after cleaning:\n{cleaned_data.head()}")
    
    return cleaned_data