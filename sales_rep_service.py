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
        attic_formatted = _prepare_report_data(attic_data, category="Attic", include_sales_rep_name=False)
        basement_formatted = _prepare_report_data(basement_data, category="Basement", include_sales_rep_name=False)
        
        # Save to Excel
        output_file = output_folder / f"{name}_Report.xlsx"
        logger.info(f"Generating sales rep report: {output_file}")
        
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            
            # Write Basement data to its own sheet
            basement_formatted.to_excel(writer, index=False, sheet_name="Basement")
            basement_worksheet = writer.sheets["Basement"]
            format_excel_sheet(basement_worksheet, basement_formatted, sales_rep=True, sheet_name="Basement")

            # Write Attic data to its own sheet
            attic_formatted=attic_formatted.drop(columns=["$ Opp to Floor","$ Opp to Target"])
            attic_formatted.to_excel(writer, index=False, sheet_name="Attic")
            attic_worksheet = writer.sheets["Attic"]
            format_excel_sheet(attic_worksheet, attic_formatted, sales_rep=True, sheet_name="Attic")
        
        return output_file
        
    except Exception as e:
        logger.error(f"Failed to generate sales rep report for {name}: {str(e)}")
        raise SalesRepServiceError(f"Failed to generate report: {str(e)}")

def _prepare_report_data(data: pd.DataFrame, category: str, include_sales_rep_name: bool = False) -> pd.DataFrame:
    """
    Prepare and format data for the sales rep report.
    
    Args:
        data: Input DataFrame
        category: Category of the data ("Attic" or "Basement")
        include_sales_rep_name: Whether to include the "Sales Rep Name" column
        
    Returns:
        Formatted DataFrame
    """
    # Drop unnecessary columns and reorder
    columns_to_drop = ["Sales Rep Email", "Manager Email", "Manager Name", "RVP Name", "RVP Email", "VP Name", "VP Email"]
    if not include_sales_rep_name:
        columns_to_drop.append("Sales Rep Name")
    
    formatted = data.drop(columns=columns_to_drop, errors='ignore')
    
    formatted["$ Gross Sales (TTM)1"] = formatted["$ Gross Sales (TTM)"]
    formatted["$ Opp to Floor1"] = formatted["$ Opp to Floor"]
    
    # Sort the data
    if include_sales_rep_name:
        cols=["Sales Rep Name"] + [col for col in formatted.columns if col != "Sales Rep Name"]
        formatted = formatted[cols]
        if category == "Attic":
            formatted = formatted.sort_values(by=["Sales Rep Name","$ Gross Sales (TTM)1"], ascending=[True,False])
        else:
            formatted = formatted.sort_values(by=["Sales Rep Name","$ Opp to Floor1"], ascending=[True, False])

        # formatted=formatted.style.set_property(subset=["Sales Rep Name"], **{'text-align', 'left'})        
    else:    
        if category == "Attic":
            formatted = formatted.sort_values(by=["$ Gross Sales (TTM)1"], ascending=False)
        else:
            formatted = formatted.sort_values(by=["$ Opp to Floor1"], ascending=False)
        
    # Drop the temporary columns after sorting
    formatted = formatted.drop(columns=["$ Gross Sales (TTM)1", "$ Opp to Floor1"])
    
    # Convert $ Gross Sales (TTM) and $ Opp to Floor to strings, right-aligned without decimals
    formatted["Item #"] = formatted["Item #"].astype(float).astype(int).astype(str)
    formatted["$ Gross Sales (TTM)"] = formatted["$ Gross Sales (TTM)"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    formatted["$ Opp to Floor"] = formatted["$ Opp to Floor"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    formatted["$ Opp to Target"] = formatted["$ Opp to Target"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    
    # Convert Margin columns to strings with one decimal place and percentage sign
    margin_columns = [col for col in formatted.columns if 'margin' in col.lower()]  # Replace with actual margin column names
    for col in margin_columns:
        if col in formatted.columns:
            formatted[col] = formatted[col].apply(lambda x: f"{100*x:.1f}%" if pd.notna(x) else "")
    
    # Retain the Sales Rep Name column if required
    if include_sales_rep_name:
        formatted["Sales Rep Name"] = data["Sales Rep Name"]
    
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
        "basement_sales": float(data[data["Category"] == "Basement"]["$ Gross Sales (TTM)"].sum()),
        "attic_sales": float(data[data["Category"] == "Attic"]["$ Gross Sales (TTM)"].sum()),
        "$ Opp_to_floor": float(data["$ Opp to Floor"].sum())
    }
    
    # Generate summary table
    summary_table = data.groupby("Category").agg({
        "$ Gross Sales (TTM)": "sum",
        "$ Opp to Floor": "sum",
        "$ Opp to Target": "sum",
        # "Sales Rep Name": "count",  # Count of lines (items) for each category
        # "Item Visibility": lambda x: ((x == "Medium") | (x == "High")).sum()
    # }).rename(columns={
        # "Sales Rep Name": "# Lines",
        # "Item Visibility": "# Visible Items"
    }).reset_index()
    
    # Format numerical values
    for col in ["$ Gross Sales (TTM)", "$ Opp to Floor", "$ Opp to Target"]:
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
            subject=f"{name}: Attic and Basement Report {month_year}",
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
        logger.warning(f"Dr$ Opped {dropped_rows} rows with missing emails")
    
    logger.debug(f"Data after cleaning:\n{cleaned_data.head()}")
    
    return cleaned_data

def generate_sales_rep_pivot_html(data: pd.DataFrame, sales_rep_name: str) -> str:
    """
    Generate two HTML pivot tables for the sales rep email:
    one for 'Basement' (sorted by '$ Opp to Floor'),
    and one for 'Attic' (sorted by '$ Gross Sales (TTM)').
    """
    try:
        data = data[data["Sales Rep Name"] == sales_rep_name]
        # Define aggregation
        agg_funcs = {
            "$ Gross Sales (TTM)": "sum",
            "$ Opp to Floor": "sum",
            "$ Opp to Target": "sum",
            # "Sales Rep Name": "count",
            # "Item Visibility": lambda x: ((x == "Medium") | (x == "High")).sum(),
        }

        # Process "Basement" table
        basement_df = (
            data[data["Category"] == "Basement"]
            .groupby(["Item Name"])
            .agg(agg_funcs)
            # .rename(columns={"Sales Rep Name": "# Lines"})  # Change to "# Lines" for clarity
            # .rename(columns={"Item Visibility": "# Visible Items"})
            .reset_index()
            .sort_values(by="$ Opp to Floor", ascending=False)  # Sort by '$ Opp to Floor'
        )

        # Process "Attic" table (Remove "$ Opp to Floor" column)
        attic_df = (
            data[data["Category"] == "Attic"]
            .groupby(["Item Name"])
            .agg(agg_funcs)
            # .rename(columns={"Sales Rep Name": "# Lines"})  # Change to "# Lines" for clarity
            # .rename(columns={"Item Visibility": "# Visible Items"})
            .reset_index()
            .drop(columns=["$ Opp to Floor", "$ Opp to Target"])  # Remove $ Opp to Floor
            .sort_values(by="$ Gross Sales (TTM)", ascending=False)  # Sort by '$ Gross Sales (TTM)'
        )

        # Add totals row to both tables
        def add_totals_row(df, columns_to_format):
            totals = {col: df[col].sum() if col in df.columns else None for col in df.columns}
            totals["Item Name"] = "Total"
            df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)
            for col in columns_to_format:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
            return df

        basement_df = add_totals_row(basement_df, ["$ Gross Sales (TTM)", "$ Opp to Floor"])
        attic_df = add_totals_row(attic_df, ["$ Gross Sales (TTM)"])

        # Convert to HTML (Align Item Name & Category to left)
        def df_to_html(df, title):
            """Format summary table as HTML with styling and title."""
            # Apply bold styling to the totals row
            def highlight_totals(row):
                if row["Item Name"] == "Total":
                    return ["font-weight: bold;" for _ in row]
                return [""] * len(row)

            styled_df = df.style.apply(highlight_totals, axis=1)

            html = f"""
            <h3>{title}</h3>
            <style>
                .summary-table th, .summary-table td {{ text-align: right; }}
                .summary-table th:first-child, .summary-table td:first-child {{ text-align: left; }}
            </style>
            {styled_df.to_html(index=False, classes="summary-table", escape=False)}
            """
            return html

        basement_html = df_to_html(basement_df, "Basement Summary")
        attic_html = df_to_html(attic_df, "Attic Summary")

        # Combine with styling
        pivot_html = f"""
        {basement_html}
        <br>
        {attic_html}
        """
        return pivot_html

    except Exception as e:
        logger.error(f"Failed to generate pivot tables for sales rep {sales_rep_name}: {str(e)}")
        raise SalesRepServiceError(f"Failed to generate pivot tables for sales rep: {str(e)}")