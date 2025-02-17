from pathlib import Path
import pandas as pd
from typing import Dict, Optional
from utils.logger import setup_logger

logger = setup_logger(__name__)

class DataProcessingError(Exception):
    """Custom exception for data processing errors."""
    pass

def load_data(input_file: Path | str) -> pd.DataFrame:
    """
    Load data from an Excel file.
    
    Args:
        input_file: Path to the input Excel file
        
    Raises:
        DataProcessingError: If file doesn't exist or can't be read
    """
    try:
        input_path = Path(input_file)
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")
            
        logger.info(f"Loading data from {input_path}")
        return pd.read_excel(input_path)
    except Exception as e:
        logger.error(f"Failed to load data: {str(e)}")
        raise DataProcessingError(f"Failed to load data: {str(e)}")

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
    cleaned_data = data.dropna(subset=["Rep Email", "Manager Email"])
    dropped_rows = initial_rows - len(cleaned_data)
    
    if dropped_rows > 0:
        logger.warning(f"Dropped {dropped_rows} rows with missing emails")
    
    return cleaned_data

def format_columns(data: pd.DataFrame) -> pd.DataFrame:
    """
    Format numerical columns (sales, opp, margin).
    
    Args:
        data: Input DataFrame
        
    Returns:
        Formatted DataFrame
        
    Raises:
        DataProcessingError: If required columns are missing
    """
    try:
        data = data.copy()
        
        # Identify columns by type
        sales_columns = [col for col in data.columns if 'sales' in col.lower()]
        opp_columns = [col for col in data.columns if 'opp' in col.lower()]
        margin_columns = [col for col in data.columns if 'margin' in col.lower()]
        
        if not any([sales_columns, opp_columns, margin_columns]):
            raise DataProcessingError("No sales, opp, or margin columns found")
        
        # Format sales and opp columns
        for col in sales_columns + opp_columns:
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0).round(0).astype(int)
            logger.debug(f"Formatted column: {col}")
        
        # Format margin columns
        for col in margin_columns:
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0).round(3)
            logger.debug(f"Formatted column: {col}")
        
        return data
        
    except Exception as e:
        logger.error(f"Error formatting columns: {str(e)}")
        raise DataProcessingError(f"Failed to format columns: {str(e)}")

def get_sales_reps(data: pd.DataFrame, limit: Optional[int] = None) -> Dict[str, str]:
    """
    Get a dictionary of sales reps (email: name).
    
    Args:
        data: Input DataFrame
        limit: Optional limit on number of reps to return
        
    Returns:
        Dictionary mapping email to name
        
    Raises:
        DataProcessingError: If required columns are missing
    """
    try:
        if "Rep Email" not in data.columns or "Rep Name" not in data.columns:
            raise DataProcessingError("Required columns 'Rep Email' or 'Rep Name' missing")
            
        reps_df = data[["Rep Email", "Rep Name"]].drop_duplicates()
        if limit:
            reps_df = reps_df.head(limit)
            
        return reps_df.set_index("Rep Email")["Rep Name"].to_dict()
        
    except Exception as e:
        logger.error(f"Error getting sales reps: {str(e)}")
        raise DataProcessingError(f"Failed to get sales reps: {str(e)}")

def get_managers(data: pd.DataFrame, start: int = 5, end: int = 7) -> Dict[str, str]:
    """
    Get a dictionary of managers (email: name) from the specified range.
    
    Args:
        data: Input DataFrame
        start: Start index for slicing
        end: End index for slicing
        
    Returns:
        Dictionary mapping email to name
        
    Raises:
        DataProcessingError: If required columns are missing
    """
    try:
        if "Manager Email" not in data.columns or "Manager Name" not in data.columns:
            raise DataProcessingError("Required columns 'Manager Email' or 'Manager Name' missing")
            
        managers_df = data[["Manager Email", "Manager Name"]].drop_duplicates()
        return managers_df.iloc[start:end].set_index("Manager Email")["Manager Name"].to_dict()
        
    except Exception as e:
        logger.error(f"Error getting managers: {str(e)}")
        raise DataProcessingError(f"Failed to get managers: {str(e)}")

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