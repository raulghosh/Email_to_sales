# excel_formatter.py
from openpyxl.styles import Alignment, PatternFill, Font, numbers
from openpyxl.utils import get_column_letter
from typing import List

def format_excel_sheet(worksheet, df: pd.DataFrame) -> None:
    """Apply formatting to an Excel worksheet."""
    # Style the header row
    header_fill = PatternFill(start_color="006400", end_color="006400", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left")
    
    # Right-align numerical columns
    numerical_columns = [
        col for col in df.columns 
        if any(keyword in col.lower() for keyword in ['sales', 'opp', 'margin'])
    ]
    
    for col_name in numerical_columns:
        col_idx = df.columns.get_loc(col_name)
        col_letter = get_column_letter(col_idx + 1)
        
        # Apply number formatting
        for row in worksheet.iter_rows(min_row=2, min_col=col_idx+1, max_col=col_idx+1):
            for cell in row:
                cell.alignment = Alignment(horizontal="right")
                if "margin" in col_name.lower():
                    cell.number_format = numbers.FORMAT_PERCENTAGE_00
                else:
                    cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
        
        # Auto-adjust column width
        max_length = max(
            len(str(cell.value)) for cell in worksheet[col_letter]
        )
        adjusted_width = min(max_length + 2, 30)
        worksheet.column_dimensions[col_letter].width = adjusted_width
    
    # Add filters to the header
    worksheet.auto_filter.ref = worksheet.dimensions