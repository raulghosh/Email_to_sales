from typing import Optional, Dict, Any

def create_email_body(
    recipient_type: str,
    name: str,
    month_year: str,
    power_bi_link: str,
    sales_rep_data: Dict[str, Any] = None,
    pivot_html: str = ""
) -> str:
    """
    Create the email body for the recipient.

    Args:
        recipient_type: Type of the recipient ("manager" or "sales_rep")
        name: Name of the recipient
        month_year: Month and year for the report
        power_bi_link: Link to the Power BI report
        sales_rep_data: Data for the sales rep (optional)
        pivot_html: HTML content for the pivot tables (optional)

    Returns:
        Email body as a string
    """
    if recipient_type == "manager":
        body = f"""
        <p>Dear {parse_first_name(name)},</p>
        <p>Please find attached the sales report for {month_year}.</p>
        {pivot_html}
        <p>You can also view the detailed report on Power BI: <a href="{power_bi_link}">Power BI Report</a></p>
        <p>Best regards,</p>
        <p>Pricing Team</p>
        """
    else:
        summary_html = sales_rep_data.get("summary_html", "") if sales_rep_data else ""
        body = f"""
        <p>Dear {parse_first_name(name)},</p>
        <p>Please find attached the sales report for {month_year}.</p>
        {summary_html}
        <p>You can also view the detailed report on Power BI: <a href="{power_bi_link}">Power BI Report</a></p>
        <p>Best regards,</p>
        <p>Pricing Team</p>
        """
    return body

def parse_first_name(full_name):
    """
    Parse the first name from a full name string.
    
    Args:
        full_name (str): Full name string in the format "Last Name, First Name Middle Name"
        
    Returns:
        str: First name extracted from the full name
    """
    # Split the name by comma to separate last name and first name
    parts = full_name.split(',')
    
    # Check if the name has a comma
    if len(parts) > 1:
        # Split the first name part by space to separate first name and middle name (if any)
        first_name_parts = parts[1].strip().split()
        
        # Return the first name
        return first_name_parts[0]
    else:
        # If no comma, assume the full name is just the first name
        return full_name.strip()


