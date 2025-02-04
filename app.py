import datetime as dt
import pandas as pd
from config import MAIN_FOLDER, INPUT_FILE_NAME, OUTPUT_FOLDER
from data_processing import load_data, clean_data, format_columns, get_sales_reps, get_managers
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
    managers = get_managers(data, limit=1)
    
    # Create output folder
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Process sales reps
    current_date = dt.datetime.now()
    month_year = current_date.strftime("%b, %Y")
    
    for email, name in sales_reps.items():
        send_sales_rep_email(formatted_data, email, name, OUTPUT_FOLDER, month_year)
    
    # Process managers
    for manager_email, manager_name in managers.items():
        manager_data = formatted_data[formatted_data["Manager Email"] == manager_email]
        send_manager_email(manager_data, manager_email, manager_name, OUTPUT_FOLDER, month_year)

if __name__ == "__main__":
    main()