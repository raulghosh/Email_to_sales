import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
from excel_formatter import format_excel_sheet

def generate_manager_report(data: pd.DataFrame, manager_name: str, output_folder: str, month_year: str) -> str:
    """Generate a manager report and save it to an Excel file."""
    output_file = os.path.join(output_folder, f"{manager_name}_Manager_Report_{month_year}.xlsx")
    
    # Ensure the column 'Opp to Floor' exists
    if 'Opp to Floor' not in data.columns:
        raise KeyError("Column 'Opp to Floor' does not exist in the data")

    # Group by 'Rep Name' and 'Region' and aggregate the required columns
    summary_table = data.groupby(["Rep Name", "Region"]).agg(
        LTM_Gross_Sales=("LTM Gross Sales", "sum"),
        Opp_to_Floor=("Opp to Floor", "sum"),
        KVI_Count=("KVI Type", lambda x: ((x == "2: KVI") | (x == "3: Super KVI")).sum()),
        Row_Count=("Rep Name", "count")
    ).reset_index()

    # Sort by 'Rep Name' and 'Region'
    sorted_summary_table = summary_table.sort_values(by=["Rep Name", "Region"])
    
    # Round off to 0 decimal places
    sorted_summary_table['LTM_Gross_Sales'] = sorted_summary_table['LTM_Gross_Sales'].round(0).astype(int)
    sorted_summary_table['Opp_to_Floor'] = sorted_summary_table['Opp_to_Floor'].round(0).astype(int)

    # Save the sorted summary table to an Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        sorted_summary_table.to_excel(writer, index=False, sheet_name='Summary')
        summary_sheet = writer.sheets['Summary']

        # Format the worksheet using the existing function
        format_excel_sheet(summary_sheet, sorted_summary_table)

        # Prepare the data for the 'All Data' sheet
        all_data = data.drop(columns=["Rep Email", "Manager Email",'Manager Name'])
        all_data = all_data.rename(columns={"Opp to Floor_x": "Opp to Floor"})
        all_data['LTM Gross Sales'] = all_data['LTM Gross Sales'].round(0).astype(int)
        all_data['Opp to Floor'] = all_data['Opp to Floor'].round(0).astype(int)
        
        # Save the 'All Data' sheet
        all_data.to_excel(writer, index=False, sheet_name='All Data')
        all_data_sheet = writer.sheets['All Data']
        format_excel_sheet(all_data_sheet, all_data)

        # Adjust column widths
        for sheet in [summary_sheet, all_data_sheet]:
            for col_num, column in enumerate(all_data.columns, 1):
                max_length = max(all_data[column].astype(str).apply(len).max(), len(column)) + 2
                sheet.column_dimensions[get_column_letter(col_num)].width = min(max_length, 30)

    return output_file