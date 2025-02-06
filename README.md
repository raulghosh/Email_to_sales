# Attic and Basement Report Automation

This project automates the generation and distribution of Attic and Basement reports for sales reps and managers. The reports are generated from an Excel file, formatted, and sent via email.

## Table of Contents

- Project Structure
- Features
- Prerequisites
- Installation
- Usage
- File Descriptions
  - app.py
  - data_processing.py
  - pivot_table_generator.py
  - manager_service.py
  - sales_rep_service.py
  - excel_formatter.py
  - config.py
- Example

## Project Structure

```markdown
├── app.py                     # Main script to load data, process it, generate reports, and send emails.
├── config.py                  # Configuration file to load environment variables and set file paths and email settings.
├── data_processing.py         # Functions to load, clean, and format data.
├── pivot_table_generator.py   # Functions to generate pivot tables and manager reports.
├── manager_service.py         # Functions to process and send emails to managers.
├── sales_rep_service.py       # Functions to process and send emails to sales reps.
├── excel_formatter.py         # Functions to format Excel sheets.
├── .env                       # Environment variables file.
└── README.md                  # Project documentation.
```

## Features

- Loads sales data from an Excel file.
- Filters out rows with missing email addresses.
- Sends emails with the filtered sales reports as attachments.

## Prerequisites

- Python 3.x
- `pandas` library
- `openpyxl` library
- `python-dotenv` library

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/your-username/Email_to_sales.git
   ```

2. Install the required Python packages:
   ```
   pip install pandas openpyxl python-dotenv
   ```

3. Create a .env file in the root directory and add the following environment variables:
   ```
   MAIN_FOLDER=<path_to_main_folder>
   EMAIL_USER=<your_email>
   ```

## Usage

1. Run the main script to generate and send the reports:
```
python app.py
```

## File Descriptions

### `app.py`

This is the main script that orchestrates the entire process:
- Loads and cleans the data.
- Formats the data.
- Generates reports for sales reps and managers.
- Sends the reports via email.

### `data_processing.py`

Contains functions to process the data:
- `load_data(input_file: str) -> pd.DataFrame`: Loads data from an Excel file.
- `clean_data(data: pd.DataFrame) -> pd.DataFrame`: Cleans the data by removing rows with missing emails.
- `format_columns(data: pd.DataFrame) -> pd.DataFrame`: Formats numerical columns.
- `get_sales_reps(data: pd.DataFrame, limit: int = 1) -> Dict[str, str]`: Retrieves a dictionary of sales reps.
- `get_managers(data: pd.DataFrame, start: int = 5, end: int = 7) -> Dict[str, str]`: Retrieves a dictionary of managers.
- `get_sorted_manager_data(data: pd.DataFrame) -> pd.DataFrame`: Sorts the manager data by "Opp to Floor".

### `pivot_table_generator.py`

Contains functions to generate pivot tables and manager reports:
- `generate_manager_report(data: pd.DataFrame, manager_name: str, output_folder: str, month_year: str) -> str`: Generates a manager report and saves it to an Excel file.

### `manager_service.py`

Contains functions to process and send emails to managers:
- `send_manager_email(data: pd.DataFrame, manager_email: str, manager_name: str, output_folder: str, month_year: str) -> None`: Processes and sends an email to a manager.

### `sales_rep_service.py`

Contains functions to process and send emails to sales reps:
- `send_sales_rep_email(data: pd.DataFrame, email: str, name: str, output_folder: str, month_year: str) -> None`: Processes and sends an email to a sales rep.

### `excel_formatter.py`

Contains functions to format Excel sheets:
- `format_excel_sheet(worksheet, df: pd.DataFrame) -> None`: Applies formatting to an Excel worksheet.

### `config.py`

Configuration file to load environment variables and set file paths and email settings:
- Loads environment variables from a `.env` file.
- Sets file paths and email settings.

### Example

Here is an example of how to use the main script:

```
import datetime as dt
import pandas as pd
from config import MAIN_FOLDER, INPUT_FILE_NAME, OUTPUT_FOLDER, EMAIL_CONFIG
from data_processing import load_data, clean_data, format_columns, get_sales_reps, get_managers, get_sorted_manager_data
from sales_rep_service import send_sales_rep_email
from manager_service import send_manager_email
import os

def main():
    # Load and clean data
    input_file = os.path.join(MAIN_FOLDER, INPUT_FILE_NAME)
    data = load_data(input_file)
    data = clean_data(data)
    formatted_data = format_columns(data)
    
    # Get sales reps and managers
    sales_reps = get_sales_reps(data, limit=1)
    managers = get_managers(data, start=5, end=10)
    
    # Create output folder
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Process sales reps
    current_date = dt.datetime.now()
    month_year = current_date.strftime("%b, %Y")
    
    for email, name in sales_reps.items():
        send_sales_rep_email(formatted_data, email, name, OUTPUT_FOLDER, month_year)
    
    # Process managers
    sorted_manager_data = get_sorted_manager_data(formatted_data)
    for manager_email, manager_name in managers.items():
        manager_data = sorted_manager_data[sorted_manager_data["Manager Email"] == manager_email]
        send_manager_email(manager_data, manager_email, manager_name, OUTPUT_FOLDER, month_year)

if __name__ == "__main__":
    main()
```
