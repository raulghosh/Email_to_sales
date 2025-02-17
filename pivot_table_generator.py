from pathlib import Path
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from typing import Optional
from excel_formatter import format_excel_sheet
from utils.logger import setup_logger

logger = setup_logger(__name__)

class PivotTableError(Exception):
    """Custom exception for pivot table generation errors."""
    pass

def generate_manager_report(
    data: pd.DataFrame,
    manager_name: str,
    output_folder: Path | str,
    month_year: str
) -> Path:
    """
    Generate a manager report and save it to an Excel file.
    
    Args:
        data: Input DataFrame
        manager_name: Name of the manager
        output_folder: Output directory path
        month_year: Month and year for the report
        
    Returns:
        Path to the generated report
        
    Raises:
        PivotTableError: If there's an error generating the report
    """
    try:
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        
        output_file = output_folder / f"{manager_name}_Manager_Report_{month_year}.xlsx"
        logger.info(f"Generating manager report: {output_file}")
        
        # Validate required columns
        required_columns = ['Opp to Floor', 'Rep Name', 'Region', 'LTM Gross Sales', 'KVI Type']
        missing_columns = [col for col in required_columns if col not in data.columns]
        if missing_columns:
            raise PivotTableError(f"Missing required columns: {', '.join(missing_columns)}")
            
        # Generate summary table
        summary_table = _create_summary_table(data)
        
        # Save to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            _write_summary_sheet(summary_table, writer)
            _write_all_data_sheet(data, writer)
            
        logger.info(f"Successfully generated report: {output_file}")
        return output_file
        
    except Exception as e:
        logger.error(f"Failed to generate manager report: {str(e)}")
        raise PivotTableError(f"Failed to generate manager report: {str(e)}")

def _create_summary_table(data: pd.DataFrame) -> pd.DataFrame:
    """Create summary table from input data."""
    summary_table = data.groupby(["Rep Name", "Region"]).agg({
        "LTM Gross Sales": "sum",
        "Opp to Floor": "sum",
        "KVI Type": lambda x: ((x == "2: KVI") | (x == "3: Super KVI")).sum(),
        "Rep Name": "count"
    }).rename(columns={"Rep Name": "Row_Count"}).reset_index()
    
    # Sort and round values
    summary_table = summary_table.sort_values(by=["Rep Name", "Region"])
    summary_table['LTM Gross Sales'] = summary_table['LTM Gross Sales'].round(0).astype(int)
    summary_table['Opp to Floor'] = summary_table['Opp to Floor'].round(0).astype(int)
    
    return summary_table

def _write_summary_sheet(summary_table: pd.DataFrame, writer: pd.ExcelWriter) -> None:
    """Write summary sheet to Excel."""
    summary_table.to_excel(writer, index=False, sheet_name='Summary')
    summary_sheet = writer.sheets['Summary']
    format_excel_sheet(summary_sheet, summary_table)

def _write_all_data_sheet(data: pd.DataFrame, writer: pd.ExcelWriter) -> None:
    """Write all data sheet to Excel."""
    all_data = data.drop(columns=["Rep Email", "Manager Email", 'Manager Name'])
    all_data = all_data.rename(columns={"Opp to Floor_x": "Opp to Floor"})
    all_data['LTM Gross Sales'] = all_data['LTM Gross Sales'].round(0).astype(int)
    all_data['Opp to Floor'] = all_data['Opp to Floor'].round(0).astype(int)
    
    all_data.to_excel(writer, index=False, sheet_name='All Data')
    all_data_sheet = writer.sheets['All Data']
    format_excel_sheet(all_data_sheet, all_data)