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
        <p>Dear {name},</p>
        <p>Please find attached the sales report for {month_year}.</p>
        <p>You can also view the detailed report on Power BI: <a href="{power_bi_link}">Power BI Report</a></p>
        {pivot_html}
        <p>Best regards,<br>Your Sales Team</p>
        """
    else:
        summary_html = sales_rep_data.get("summary_html", "") if sales_rep_data else ""
        body = f"""
        <p>Dear {name},</p>
        <p>Please find attached the sales report for {month_year}.</p>
        <p>You can also view the detailed report on Power BI: <a href="{power_bi_link}">Power BI Report</a></p>
        {summary_html}
        <p>Best regards,<br>Your Sales Team</p>
        """
    return body


