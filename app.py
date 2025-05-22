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
    
    managers = get_managers(data)
    print(managers)
    os.makedirs(CONFIG.output_folder, exist_ok=True)
    
    # Process sales reps
    current_date = dt.datetime.now()
    month_year = current_date.strftime("%b, %Y")
    
    # activate when ready to send emails to reps
    
    # for email, name in sales_reps.items():
    #     if name:  # Check if the Sales Rep Name is non-null
    #         try:
    #             send_sales_rep_email(formatted_data, email, name, CONFIG.output_folder, month_year)
    #         except Exception as e:
    #             logger.error(f"Error processing sales rep {name} ({email}): {e}")



    # Define exclusion lists
    EXCLUDED_MANAGER_NAMES = ['macpherson, scott', 'mayerle, stephanie', 'moy, stephanie', 'stonebrook, ryan e']
    EXCLUDED_MANAGER_EMAILS = ['smacphe@veritivcorp.com', 'stoner03@veritivcorp.com', 'smoy@veritivcorp.com', 'smayerl@veritivcorp.com']

    # Process managers
    for manager_email, manager_name in managers.items():
        if (
            manager_name
            and manager_name.lower() not in EXCLUDED_MANAGER_NAMES
            and manager_email.lower() not in EXCLUDED_MANAGER_EMAILS
        ):
            try:
                send_manager_email(formatted_data, manager_email, manager_name, CONFIG.output_folder, month_year)
            except Exception as e:
                logger.error(f"Error processing manager {manager_name} ({manager_email}): {e}")

if __name__ == "__main__":
    main()