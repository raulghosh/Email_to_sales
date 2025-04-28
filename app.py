import datetime as dt
import pandas as pd
from config import CONFIG
from data_processing import load_data, clean_data, format_columns, get_sales_reps, get_managers
from sales_rep_service import send_sales_rep_email
from manager_service import send_manager_email
import os
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def main():
    # Load and clean data
    input_file = os.path.join(CONFIG.main_folder, CONFIG.input_file_name)
    data = load_data(input_file)
    data = clean_data(data)
    formatted_data = format_columns(data)
    
    # Log the formatted data for verification
    logger.debug(f"Formatted data:\n{formatted_data.head()}")
    
    # Get sales reps and managers
    sales_reps = get_sales_reps(data,limit =2)  # For testing, limit to 2 sales reps
    print(sales_reps)  # For debugging purposes, to see the sales reps dictionary
    # sales_reps={
    # 'Ellis, Aimee': 'aellis@veritivcorp.com',
    # 'Geiger, Kimiko': 'kgeiger@veritivcorp.com',
    # 'Alford, Dustin T': 'dalford@veritivcorp.com',
    # 'Bell, Nicole': 'nbell@veritivcorp.com',
    # 'Montibon, Donna': 'dmontib@veritivcorp.com',
    # 'Reisinger, Jeremy': 'jreisin@veritivcorp.com'}
    managers = get_managers(data)
    print(managers)  # For debugging purposes, to see the managers dictionary
    # managers={'Reisinger, Jeremy': 'jreisin@veritivcorp.com'}
    # Create output folder
    os.makedirs(CONFIG.output_folder, exist_ok=True)
    
    # Process sales reps
    current_date = dt.datetime.now()
    month_year = current_date.strftime("%b, %Y")
    
    for email, name in sales_reps.items():
        if name:  # Check if the Sales Rep Name is non-null
            send_sales_rep_email(formatted_data, email, name, CONFIG.output_folder, month_year)
    
    # Process managers
    for manager_email, manager_name in managers.items():
        if manager_name:  # Check if the manager name is non-null
            send_manager_email(formatted_data, manager_email, manager_name, CONFIG.output_folder, month_year)

if __name__ == "__main__":
    main()