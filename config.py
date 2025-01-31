# config.py
import os
from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env

# File paths
MAIN_FOLDER = os.getenv("MAIN_FOLDER")
INPUT_FILE_NAME = "Sales_Report_Temp.xlsx"
OUTPUT_FOLDER = os.path.join(MAIN_FOLDER, "Filtered_Reports")

# Email settings
EMAIL_CONFIG = {
    "smtp_server": "relay.int.distco.com",
    "smtp_port": 25,
    "sender_email": os.getenv("EMAIL_USER"),
    "test_email": "rghosh@veritivcorp.com"
}

# Power BI link
POWER_BI_LINK = "https://your-powerbi-link.com"