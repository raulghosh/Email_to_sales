from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from utils.logger import setup_logger
from excel_formatter import format_excel_sheet
from sales_rep_service import _prepare_report_data

logger = setup_logger(__name__)

class PivotTableError(Exception):
    """Custom exception for pivot table generation errors."""
    pass

class ManagerServiceError(Exception):
    """Custom exception for manager service errors."""
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

        # Split data into Attic and Basement
        attic_data = manager_data[manager_data["Category"] == "Attic"]
        basement_data = manager_data[manager_data["Category"] == "Basement"]

        # Format the data
        attic_formatted = _prepare_report_data(attic_data, category="Attic", include_sales_rep_name=True)
        basement_formatted = _prepare_report_data(basement_data, category="Basement", include_sales_rep_name=True)

        # Generate summary tables
        attic_summary = _create_summary_table(attic_data, category="Attic")
        basement_summary = _create_summary_table(basement_data, category="Basement")

        # Write to Excel
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            _write_summary_sheet(basement_summary, writer, "Basement Summary", category="Basement")
            _write_summary_sheet(attic_summary, writer, "Attic Summary", category="Attic")
            _write_data_sheet(basement_formatted, writer, "Basement", category="Basement")
            _write_data_sheet(attic_formatted, writer, "Attic", category="Attic")

        return file_path

    except ValueError as ve:
        logger.error(f"Data validation error: {ve}")
    except PermissionError:
        logger.error(f"Permission denied: Unable to write file {file_path}. Check if the file is open.")
    except Exception as e:
        logger.error(f"Unexpected error in generate_manager_report: {str(e)}")

    return None

def _create_summary_table(data: pd.DataFrame, category: str) -> pd.DataFrame:
    """Create an aggregated summary table from input data with a totals row."""
    if data.empty:
        return pd.DataFrame()
    
    # Group and aggregate data
    summary_table = data.groupby("Sales Rep Name").agg({
        "$ Gross Sales (TTM)": "sum",
        "$ Opp to Floor": "sum",
        # "Manager Name": "count",  # Count of lines for each sales rep
        # "Item Visibility": lambda x: ((x == "Medium") | (x == "High")).sum(),
    }).reset_index()
    
    # Create two columns for sorting
    summary_table["$ Gross Sales LTM1"] = summary_table["$ Gross Sales (TTM)"]
    summary_table["$ Opp to Floor1"] = summary_table["$ Opp to Floor"]
    
    # Format numerical columns
    summary_table["$ Gross Sales (TTM)"] = summary_table["$ Gross Sales (TTM)"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    summary_table["$ Opp to Floor"] = summary_table["$ Opp to Floor"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    
    # Sort the tables
    if category == "Attic":
        summary_table = summary_table.sort_values(by="$ Gross Sales LTM1", ascending=False).drop(columns=["$ Gross Sales LTM1", "$ Opp to Floor1"])
    else:
        summary_table = summary_table.sort_values(by="$ Opp to Floor1", ascending=False).drop(columns=["$ Gross Sales LTM1", "$ Opp to Floor1"])
    
    # Add totals row
    totals = {
        "Sales Rep Name": "Total",
        "$ Gross Sales (TTM)": summary_table["$ Gross Sales (TTM)"].replace(",", "", regex=True).astype(float).sum(),
        "$ Opp to Floor": summary_table["$ Opp to Floor"].replace(",", "", regex=True).astype(float).sum(),
        # "# Lines": summary_table["# Lines"].sum(),  # Count of lines for each sales rep
        # "# Visible Items": summary_table["# Visible Items"].sum(),
    }
    totals["$ Gross Sales (TTM)"] = f"{int(totals['$ Gross Sales (TTM)']):,}"
    totals["$ Opp to Floor"] = f"{int(totals['$ Opp to Floor']):,}"
    summary_table = pd.concat([summary_table, pd.DataFrame([totals])], ignore_index=True)
    
    return summary_table

def _write_summary_sheet(summary_table: pd.DataFrame, writer: pd.ExcelWriter, sheet_name: str, category: str) -> None:
    """Write an aggregated summary sheet to Excel."""
    if category == "Attic":
        summary_table.drop(columns=["$ Opp to Floor"], inplace=True)
    summary_table.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]
    format_excel_sheet(worksheet, summary_table, sales_rep=False,sheet_name=sheet_name)

def _write_data_sheet(data: pd.DataFrame, writer: pd.ExcelWriter, sheet_name: str, category: str) -> None:
    """Write formatted data to an Excel sheet."""
    if category == "Attic":
        data.drop(columns=["$ Opp to Floor", "$ Opp to Target"], inplace=True)
    data=data.drop(columns=["Category"])   
    data.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]
    format_excel_sheet(worksheet, data, sales_rep=False,sheet_name = sheet_name)

def generate_manager_pivot_html(data: pd.DataFrame, manager_name: str) -> str:
    """
    Generate two HTML pivot tables for the manager email: 
    one for 'Basement' (sorted by '$ Opp to Floor'), 
    and one for 'Attic' (sorted by '$ Gross Sales (TTM)').
    """
    try:
        data = data[data["Manager Name"] == manager_name]
        # Define aggregation
        agg_funcs = {
            "$ Gross Sales (TTM)": "sum",
            "$ Opp to Floor": "sum",
            "$ Opp to Target": "sum",
            # "Manager Name": "count",
            # "Item Visibility": lambda x: ((x == "Medium") | (x == "High")).sum(),
        }

        # Process "Basement" table
        basement_df = (
            data[data["Category"] == "Basement"]
            .groupby(["Sales Rep Name"])
            .agg(agg_funcs)
            # .rename(columns={"Manager Name": "# Lines"})  # Rename "Manager Name" to "# Lines"
            # .rename(columns={"Item Visibility": "# Visible Items"})
            .reset_index()
            .sort_values(by="$ Opp to Floor", ascending=False)  # Sort by '$ Opp to Floor'
        )

        # Process "Attic" table (Remove "$ Opp to Floor" column)
        attic_df = (
            data[data["Category"] == "Attic"]
            .groupby(["Sales Rep Name"])
            .agg(agg_funcs)
            # .rename(columns={"Manager Name": "# Lines"})  # Rename "Manager Name" to "# Lines"
            # .rename(columns={ "Item Visibility": "# Visible Items"})
            .reset_index()
            .drop(columns=["$ Opp to Floor","$ Opp to Target"])  # Remove $ Opp to Floor
            .sort_values(by="$ Gross Sales (TTM)", ascending=False)  # Sort by '$ Gross Sales (TTM)'
        )

        # Format numerical columns
        for df in [basement_df, attic_df]:
            for col in ["$ Gross Sales (TTM)"]:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: f"{x:,.0f}")
        # Format "$ Opp to Floor" column in basement_df
        if "$ Opp to Floor" in basement_df.columns:
            basement_df["$ Opp to Floor"] = basement_df["$ Opp to Floor"].apply(lambda x: f"{x:,.0f}")
        if "$ Opp to Floor" in basement_df.columns:
            basement_df["$ Opp to Target"] = basement_df["$ Opp to Target"].apply(lambda x: f"{x:,.0f}")

        # Convert to HTML (Align Sales Rep Name & Category to left)
        def _format_summary_table_html(summary_table: pd.DataFrame) -> str:
            """Format summary table as HTML with styling."""
            return f"""
            <style>
                .summary-table th, .summary-table td {{ text-align: right; }}
                .summary-table th:first-child, .summary-table td:first-child {{ text-align: left; }}
            </style>
            {summary_table.to_html(index=False, classes="summary-table")}
            """

        basement_html = _format_summary_table_html(basement_df)
        attic_html = _format_summary_table_html(attic_df)

        # Combine with styling
        pivot_html = f"""
        <h3>Basement Summary</h3>
        {basement_html}
        <br>
        <h3>Attic Summary</h3>
        {attic_html}
        """
        raise Exception(f"Failed to generate pivot tables: {str(e)}")

    except Exception as e:
        logger.error(f"Failed to generate pivot tables: {str(e)}")
        raise ManagerServiceError(f"Failed to generate pivot tables: {str(e)}")