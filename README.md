# Attic and Basement Report Automation

<<<<<<< HEAD
An automated system for generating and distributing sales performance reports to sales representatives and managers. The system processes Excel-based sales data, creates customized reports, and distributes them via email.
=======
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
>>>>>>> a2f48ef27241dbdc01381ec8302cfa52ceb5914e

## Features

- **Data Processing**
  - Automated loading and cleaning of sales data
  - Intelligent handling of missing data
  - Custom formatting for sales and margin data

- **Report Generation**
  - Personalized reports for sales representatives
  - Comprehensive manager summaries
  - KVI (Key Value Item) analysis
  - Performance metrics calculation

- **Email Distribution**
  - Automated email distribution
  - Customized email content for different roles
  - Embedded performance summaries
  - Excel report attachments

- **Error Handling & Logging**
  - Comprehensive error tracking
  - Detailed logging system
  - Robust exception handling

## Project Structure

```
Attic-and-Basement-Report-Automation/
├── app.py                     # Main application entry point
├── config.py                  # Configuration management
├── data_processing.py         # Data loading and processing
├── pivot_table_generator.py   # Report generation utilities
├── manager_service.py         # Manager report handling
├── sales_rep_service.py       # Sales rep report handling
├── email_handler.py           # Email distribution system
├── email_composer.py          # Email content generation
├── excel_formatter.py         # Excel formatting utilities
├── utils/
│   └── logger.py             # Logging configuration
├── logs/                      # Log files directory
├── .env                      # Environment variables
├── requirements.txt          # Project dependencies
└── README.md                 # Project documentation
```

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/attic-and-basement-automation.git
cd attic-and-basement-automation
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file with required configuration:
```env
EMAIL_USER=your.email@company.com
MAIN_FOLDER=/path/to/data/folder
```

## Configuration

The system uses a configuration system based on environment variables and the `config.py` file:

- **Email Settings**
  - SMTP server configuration
  - Sender email settings
  - Test mode options

- **File Paths**
  - Input file location
  - Output directory
  - Log file location

- **Report Settings**
  - Power BI dashboard links
  - Report formatting options

## Usage

1. Ensure your input Excel file is in the configured location

2. Run the main script:
```bash
python app.py
```

3. Monitor the logs directory for execution details

## Development

### Code Style

This project follows PEP 8 guidelines and uses:
- Black for code formatting
- Pylint for code quality
- Mypy for type checking

### Testing

Run tests using:
```bash
python -m pytest tests/
```

### Logging

<<<<<<< HEAD
Logs are written to the `logs/` directory with different files for:
- Application events
- Email operations
- Data processing
- Error tracking

## Error Handling

The system implements a comprehensive error handling system:
- Custom exceptions for different components
- Detailed error logging
- Graceful failure handling
- Email notifications for critical errors

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support and questions, please contact:
- Email: your.email@company.com
- Internal Ticket System: Link to ticket system

## Acknowledgments

- Sales Operations Team
- IT Infrastructure Team
- Data Analytics Team
=======
if __name__ == "__main__":
    main()
```
>>>>>>> a2f48ef27241dbdc01381ec8302cfa52ceb5914e
