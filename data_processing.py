# data_processing.py
import pandas as pd
from typing import Tuple, Dict

def load_data(input_file: str) -> pd.DataFrame:
    """Load data from an Excel file."""
    return pd.read_excel(input_file)

def clean_data(data: pd.DataFrame) -> pd.DataFrame:
    """Remove rows with missing Rep/Manager emails."""
    return data.dropna(subset=["Rep Email", "Manager Email"])

def format_columns(data: pd.DataFrame) -> pd.DataFrame:
    """Format numerical columns (sales, opp, margin)."""
    data = data.copy()
    
    # Sales and Opp columns: round to integers
    sales_columns = [col for col in data.columns if 'sales' in col.lower()]
    opp_columns = [col for col in data.columns if 'opp' in col.lower()]
    for col in sales_columns + opp_columns:
        data[col] = data[col].fillna(0).round(0).astype(int)
    
    # Margin columns: convert to percentages
    margin_columns = [col for col in data.columns if 'margin' in col.lower()]
    for col in margin_columns:
        data[col] = (data[col].fillna(0)).round(3)
    
    return data

def get_sales_reps(data: pd.DataFrame, limit: int = 3) -> Dict[str, str]:
    """Get a dictionary of sales reps (email: name)."""
    return data[["Rep Email", "Rep Name"]].drop_duplicates().head(limit).set_index("Rep Email")["Rep Name"].to_dict()

def get_managers(data: pd.DataFrame, limit: int = 3) -> Dict[str, str]:
    """Get a dictionary of sales reps (email: name)."""
    return data[["Manager Email", "Manager Name"]].drop_duplicates().head(limit).set_index("Manager Email")["Manager Name"].to_dict()