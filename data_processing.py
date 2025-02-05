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

def get_sales_reps(data: pd.DataFrame, limit: int = 1) -> Dict[str, str]:
    """Get a dictionary of sales reps (email: name)."""
    return data[["Rep Email", "Rep Name"]].drop_duplicates().head(limit).set_index("Rep Email")["Rep Name"].to_dict()

def get_managers(data: pd.DataFrame, start: int = 5, end: int = 7) -> Dict[str, str]:
    """Get a dictionary of managers (email: name) from the specified range."""
    return data[["Manager Email", "Manager Name"]].drop_duplicates().iloc[start:end].set_index("Manager Email")["Manager Name"].to_dict()

def get_sorted_manager_data(data: pd.DataFrame) -> pd.DataFrame:
    """Sort the manager data by Opp to Floor."""
    # Group by Rep Name and sum Opp to Floor
    grouped_data = data.groupby("Rep Name")["Opp to Floor"].sum().reset_index()
    
    # Sort by Opp to Floor in descending order
    sorted_grouped_data = grouped_data.sort_values(by="Opp to Floor", ascending=False)
    
    # Merge sorted grouped data with the original data to maintain the sorting order
    sorted_data = pd.merge(sorted_grouped_data, data, on="Rep Name", how="left")
    
    # Rename 'Opp to Floor_y' to 'Opp to Floor'
    sorted_data.rename(columns={"Opp to Floor_y": "Opp to Floor"}, inplace=True)
    
    # Drop 'Opp to Floor_x' column
    sorted_data.drop(columns="Opp to Floor_x", inplace=True)
    
    return sorted_data