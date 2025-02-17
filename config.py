from dataclasses import dataclass
from typing import Optional
import os
from dotenv import load_dotenv
import logging

@dataclass
class EmailConfig:
    smtp_server: str
    smtp_port: int
    sender_email: str
    test_email: str

@dataclass
class AppConfig:
    main_folder: str
    input_file_name: str
    output_folder: str
    email_config: EmailConfig
    power_bi_link: str

def load_config() -> AppConfig:
    """Load and validate configuration settings."""
    load_dotenv()
    
    main_folder = os.getenv("MAIN_FOLDER")
    if not main_folder:
        raise ValueError("MAIN_FOLDER environment variable is not set")
    
    email_user = os.getenv("EMAIL_USER")
    if not email_user:
        raise ValueError("EMAIL_USER environment variable is not set")

    email_config = EmailConfig(
        smtp_server="relay.int.distco.com",
        smtp_port=25,
        sender_email=email_user,
        test_email="rghosh@veritivcorp.com"
    )

    return AppConfig(
        main_folder=main_folder,
        input_file_name="Sales_Report_Temp.xlsx",
        output_folder=os.path.join(main_folder, "Filtered_Reports"),
        email_config=email_config,
        power_bi_link="https://app.powerbi.com/links/la6Wz4H0aX?ctid=ab15d0ad-ff4d-4eb2-b09b-e0743223e142&pbi_source=linkShare"
    )

# Global config instance
CONFIG = load_config()