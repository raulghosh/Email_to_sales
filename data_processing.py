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
    cleaned_data = data.dropna(subset=["Sales Rep Email", "Manager Email"])
    dropped_rows = initial_rows - len(cleaned_data)
    
    if dropped_rows > 0:
        logger.warning(f"Dropped {dropped_rows} rows with missing emails")
    
    cleaned_data=cleaned_data[cleaned_data["Category"]!='Market']
    
    dropped_rows = initial_rows - len(cleaned_data)
    
    if dropped_rows > 0:
        logger.warning(f"Dropped {dropped_rows} rows with missing emails")
        
    logger.debug(f"Data after cleaning:\n{cleaned_data.head()}")
    
    
    
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
        # columns = Category | Customer Name | Bill_to # | Legacy Item # | Item Desc | Channel | Gross Sales (TTM) | Comm. Margin (TTM) | Last Comm. Margin |\
        # Last Trans. Date | Floor Margin | Target Margin | Start Margin | Opp to Floor | Opp to Target | Item Visibility |Vendor Name | Cat1 |\
        # Sales Rep Name | Sales Rep Email | Manager Name | Manager Email | RVP Name | RVP Email | VP Name | VP Email


        sales_columns = [col for col in data.columns if 'sales' in col.lower() and 'rep' not in col.lower() and 'margin' not in col.lower()]
        sales_columns = [col for col in data.columns if 'sales' in col.lower() and 'rep' not in col.lower()]
        opp_columns = [col for col in data.columns if 'opp' in col.lower()]
        margin_columns = [col for col in data.columns if 'margin' in col.lower()]
        
        if not any([sales_columns, opp_columns, margin_columns]):
            raise DataProcessingError("No sales, opp, or margin columns found")
        
        logger.debug(f"Sales columns: {sales_columns}")
        logger.debug(f"Opp columns: {opp_columns}")
        logger.debug(f"Margin columns: {margin_columns}")
        
        # Force Legacy Item # to string
        if "Legacy Item #" in data.columns:
            data["Legacy Item #"] = data["Legacy Item #"].astype(str)
            logger.debug("Converted 'Legacy Item #' to string")
        
        # Format sales and opp columns
        for col in sales_columns + opp_columns:
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0).round(0).astype(int)
            logger.debug(f"Formatted column: {col}")
            logger.debug(f"Data after formatting {col}:\n{data[[col]].head()}")
        
        # Format margin columns
        for col in margin_columns:
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0).round(3)
            logger.debug(f"Formatted column: {col}")
            logger.debug(f"Data after formatting {col}:\n{data[[col]].head()}")
        
        for col in opp_columns:
            data.rename(columns={col: col.replace("Opp", "$ Opp")}, inplace=True)
        
        logger.debug(f"Data after formatting all columns:\n{data.head()}")        
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
        if "Sales Rep Email" not in data.columns or "Sales Rep Name" not in data.columns:
            raise DataProcessingError("Required columns 'Sales Rep Email' or 'Sales Rep Name' missing")
            
        reps_df = data[["Sales Rep Email", "Sales Rep Name"]].drop_duplicates()
        if limit:
            reps_df = reps_df.head(limit)
            
        return reps_df.set_index("Sales Rep Email")["Sales Rep Name"].to_dict()
        
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