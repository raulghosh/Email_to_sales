from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from utils.logger import setup_logger
from excel_formatter import format_excel_sheet

logger = setup_logger(__name__)

class PivotTableError(Exception):
    """Custom exception for pivot table generation errors."""
    pass

def generate_manager_report(
    data: pd.DataFrame,
    manager_name: str,
    output_folder: Path,
    month_year: str
) -> Path:
    """
    Generate a manager report and save it to an Excel file.

    Args:
        data: Input DataFrame containing sales data.
        manager_name: Name of the manager.
        output_folder: Directory where the report will be saved.
        month_year: Month and year string for naming the report.

    Returns:
        Path to the generated Excel file if successful, otherwise None.
    """
    try:
        # Ensure output_folder is a Path object
        output_folder = Path(output_folder)  
        output_folder.mkdir(parents=True, exist_ok=True)
        file_path = output_folder / f"{manager_name}_Manager_Report_{month_year}.xlsx"

        # Filter data for the given manager
        manager_data = data[data["Manager Name"] == manager_name]

        # Generate summary tables
        attic_summary = _create_summary_table(manager_data[manager_data["Region"] == "Attic"])
        basement_summary = _create_summary_table(manager_data[manager_data["Region"] == "Basement"])

        # Write to Excel
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            _write_summary_sheet(basement_summary, writer, "Basement Summary")
            _write_summary_sheet(attic_summary, writer, "Attic Summary")
            _write_all_data_sheet(manager_data, writer)

        return file_path

    except ValueError as ve:
        logger.error(f"Data validation error: {ve}")
    except PermissionError:
        logger.error(f"Permission denied: Unable to write file {file_path}. Check if the file is open.")
    except Exception as e:
        logger.error(f"Unexpected error in generate_manager_report: {str(e)}")

    return None

def _create_summary_table(data: pd.DataFrame) -> pd.DataFrame:
    """Create an aggregated summary table from input data."""
    if data.empty:
        return pd.DataFrame()
    
    summary_table = data.groupby("Rep Name").agg({
        "LTM Gross Sales": "sum",
        "Opp to Floor": "sum",
        "KVI Type": lambda x: ((x == "2: KVI") | (x == "3: Super KVI")).sum(),
        "Rep Name": "count"
    }).rename(columns={"Rep Name": "# Rows", "KVI Type": "# Visible Items"}).reset_index()
    
    return summary_table.sort_values(by="Rep Name")

def _write_summary_sheet(summary_table: pd.DataFrame, writer: pd.ExcelWriter, sheet_name: str) -> None:
    """Write an aggregated summary sheet to Excel."""
    summary_table.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]
    format_excel_sheet(worksheet, summary_table)

def _write_all_data_sheet(data: pd.DataFrame, writer: pd.ExcelWriter) -> None:
    """Write all filtered data to an Excel sheet."""
    all_data = data.drop(columns=["Rep Email", "Manager Email"])
    all_data.to_excel(writer, index=False, sheet_name='All Data')
    worksheet = writer.sheets['All Data']
    format_excel_sheet(worksheet, all_data)

def generate_html_table(data: pd.DataFrame, title: str) -> str:
    """Convert a Pandas DataFrame into an HTML table with a title and format numerical columns."""
    try:
        if data.empty:
            return f"<h3>{title}</h3><p>No data available.</p>"

        # Format numerical columns
        for col in data.columns:
            if any(keyword in col.lower() for keyword in ['sales', 'opp', 'count']):  # Check for relevant keywords
                data[col] = data[col].apply(lambda x: f"{x:,.0f}")

        return f"<h3>{title}</h3>" + data.to_html(index=False, escape=False)
    except Exception as e:
        logger.error(f"Error generating HTML table: {e}")
        return f"<h3>{title}</h3><p>Unable to display data.</p>"