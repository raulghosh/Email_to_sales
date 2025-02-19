from typing import Optional, Dict, Any

def create_email_body(
    recipient_type: str,
    name: str,
    month_year: str,
    power_bi_link: str,
    sales_rep_data: Optional[Dict[str, Any]] = None,
    pivot_html: Optional[str] = None
) -> str:
    """
    Create an HTML email body tailored to sales reps or managers.
    
    Args:
        recipient_type: "sales_rep" or "manager".
        name: Recipient's name.
        month_year: Month/year of the report (e.g., "Oct 2023").
        power_bi_link: URL to the Power BI dashboard.
        sales_rep_data: Required for "sales_rep" (keys: basement_count, attic_count, basement_sales, attic_sales, opp_to_floor, summary_html).
        pivot_html: Required for "manager" (HTML table of pivot data).
    """
    if recipient_type == "sales_rep" and sales_rep_data:
        return _sales_rep_email_body(name, month_year, power_bi_link, **sales_rep_data)
    elif recipient_type == "manager" and pivot_html:
        return _manager_email_body(name, month_year, power_bi_link, pivot_html)
    else:
        raise ValueError("Invalid recipient type or missing data!")

def _sales_rep_email_body(
    name: str,
    month_year: str,
    power_bi_link: str,
    basement_count: int,
    attic_count: int,
    basement_sales: float,
    attic_sales: float,
    opp_to_floor: float,
    summary_html: str
) -> str:
    """Generate HTML body for sales reps."""
    return f"""
    <div style="text-align: left;">
        <p>Hi {name},</p>
        <p>Hope you are doing good.</p> 
        <p>Attached is the Attic and Basement Report for {month_year}.</p>
        <p>This month you have {basement_count} action items in 'Basement' corresponding to ${basement_sales:,.0f} of gross sales and you have {attic_count} action items in 'Attic' corresponding to {attic_sales:,.0f} of gross sales.</p>
        <p>Raising the items in basement to the recommended margin will result in ${opp_to_floor:,.0f} of commission profit gain.</p>
        <p>Summary by Region:</p>
        {summary_html}
        <p>Access the live Power BI Dashboard: <a href="{power_bi_link}">Attic and Basement Report</a></p>
        <p>Thanks,<br>Pricing Team</p>
    </div>
    """

def _manager_email_body(
    name: str,
    month_year: str,
    power_bi_link: str,
    pivot_html: str  # Updated argument
) -> str:
    """Generate HTML email body for managers with Basement and Attic tables."""
    return f"""
    <div style="text-align: left;">
        <p>Hi {name},</p>
        <p>Attached is the Manager Report for {month_year}.</p>
        <p>Below are summaries for the Basement and Attic items:</p>
        {pivot_html}
        <p>Access the live Power BI Dashboard: <a href="{power_bi_link}">Manager Report</a></p>
        <p>Thanks,<br>Pricing Team</p>
    </div>
    """

