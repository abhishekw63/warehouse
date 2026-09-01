import smtplib
from email.message import EmailMessage
from datetime import datetime
from config import Config
from utils import format_indian

class EmailService:
    @staticmethod
    def send_email_summary(marketplace, summary_data, tracker_df=None, sku_df=None):
        """
        Send email summary with PO report details.

        Args:
            marketplace: Name of the marketplace
            summary_data: Dictionary containing summary statistics
            tracker_df: DataFrame with tracker data (optional)
            sku_df: DataFrame with SKU demand data (optional)
        """
        try:
            # Create email message
            message = EmailMessage()
            message["From"] = Config.EMAIL_SENDER
            message["To"] = Config.DEFAULT_RECIPIENT

            # Add CC recipients if any are configured
            if Config.CC_RECIPIENTS:
                message["Cc"] = ", ".join(Config.CC_RECIPIENTS)

            message["Subject"] = f"📊 Purchase Order Summary: {marketplace} - {datetime.now().strftime('%d-%m-%Y')}"

            # Build HTML email body
            html_body = EmailService._build_email_html(marketplace, summary_data, tracker_df, sku_df)

            # Set email content
            message.set_content("Please view this email in an HTML-compatible email client.")
            message.add_alternative(html_body, subtype="html")

            # Send email
            server = smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT)
            server.starttls()
            server.login(Config.EMAIL_SENDER, Config.EMAIL_PASSWORD)

            # Prepare all recipients (TO + CC)
            all_recipients = [Config.DEFAULT_RECIPIENT]
            if Config.CC_RECIPIENTS:
                all_recipients.extend(Config.CC_RECIPIENTS)

            server.send_message(message, to_addrs=all_recipients)
            server.quit()

            return True, ""

        except Exception as e:
            return False, str(e)

    @staticmethod
    def _build_email_html(marketplace, summary_data, tracker_df, sku_df=None):
        """Build HTML email body with summary, tracker, and SKU data."""

        # Extract summary data
        total_pos = summary_data.get('total_pos', 0)
        total_units = summary_data.get('total_units', 0)
        total_value = summary_data.get('total_value', 0)
        min_date = summary_data.get('min_date', 'N/A')
        max_date = summary_data.get('max_date', 'N/A')

        # Check if SKU data exists
        has_sku = sku_df is not None and not sku_df.empty

        html = f"""
<html>
<head>
<style>
body {{
    font-family: Arial, sans-serif;
    margin: 0;
    padding: 0;
}}
.header {{
    background-color: #4472C4;
    color: white;
    padding: 15px;
    text-align: center;
}}
.summary-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 20px 0;
}}
.summary-table td {{
    padding: 10px;
    border: 1px solid #ddd;
}}
.summary-table td:first-child {{
    background-color: #f2f2f2;
    font-weight: bold;
    width: 200px;
}}
.data-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 20px 0;
}}
.data-table th {{
    background-color: #4472C4;
    color: white;
    padding: 10px;
    text-align: center;
    border: 1px solid #ddd;
}}
.data-table td {{
    padding: 8px;
    text-align: center;
    border: 1px solid #ddd;
}}
.data-table tr:nth-child(even) {{
    background-color: #f9f9f9;
}}
.sku-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 20px 0;
}}
.sku-table th {{
    background-color: #70AD47;
    color: white;
    padding: 10px;
    text-align: center;
    border: 1px solid #ddd;
}}
.sku-table td {{
    padding: 8px;
    text-align: center;
    border: 1px solid #ddd;
}}
.sku-table tr:nth-child(even) {{
    background-color: #f9f9f9;
}}
.footer {{
    text-align: center;
    color: #666;
    font-size: 11px;
    margin-top: 20px;
    padding-top: 10px;
    border-top: 1px solid #ddd;
}}
</style>
</head>
<body>
<div class="header">
<h2 style="margin: 0;">📊 {marketplace} PO Report - {datetime.now().strftime('%d-%m-%Y %H:%M')}</h2>
</div>

<h3 style="margin: 20px 0 10px 0;">Summary Overview</h3>
<table class="summary-table">
<tr>
    <td>Total POs:</td>
    <td>{total_pos}</td>
</tr>
<tr>
    <td>Order Date Range:</td>
    <td>{min_date} to {max_date}</td>
</tr>
<tr>
    <td>Total Units Ordered:</td>
    <td>{format_indian(total_units)}</td>
</tr>
<tr>
    <td>Total Order Value:</td>
    <td>₹ {format_indian(total_value)}</td>
</tr>
<tr>
    <td>SKU Data Available:</td>
    <td>{'✅ YES - ' + str(len(sku_df)) + ' SKUs' if has_sku else '❌ NO (' + marketplace + ' format)'}</td>
</tr>
</table>
"""

        # Add PO Details table if tracker_df is provided
        if tracker_df is not None and not tracker_df.empty:
            html += '<h3 style="margin: 20px 0 10px 0;">PO Details</h3>\n<table class="data-table">\n<tr>'

            # Add table headers
            for col in tracker_df.columns:
                display_col = str(col).replace('_', ' ').title().replace('Po ', 'PO ')
                html += f"<th>{display_col}</th>"
            html += "</tr>\n"

            # Add table rows
            for row in tracker_df.itertuples(index=False):
                html += "<tr>"
                for col_idx, col in enumerate(tracker_df.columns):
                    cell_value = row[col_idx]

                    # Format dates to DD-MM-YYYY
                    if 'Date' in col and hasattr(cell_value, 'strftime'):
                        cell_value = cell_value.strftime('%d-%m-%Y')
                    # Format PO Value with Indian currency formatting
                    elif col == 'PO Value' or 'total_amount' in col:
                        if isinstance(cell_value, str):
                            clean_value = cell_value.replace('₹', '').replace(',', '').strip()
                            try:
                                numeric_value = float(clean_value)
                                cell_value = f"₹ {format_indian(numeric_value)}"
                            except:
                                pass
                        elif isinstance(cell_value, (int, float)):
                            cell_value = f"₹ {format_indian(cell_value)}"

                    html += f"<td>{cell_value}</td>"
                html += "</tr>\n"


            html += "</table>"

        # Add SKU Demand table if sku_df is provided
        if has_sku:
            html += '<h3 style="margin: 20px 0 10px 0;">SKU Demand</h3>\n<table class="sku-table">\n<tr>'

            # Add SKU table headers
            for col in sku_df.columns:
                display_col = str(col).replace('_', ' ').title()
                html += f"<th>{display_col}</th>"
            html += "</tr>\n"

            # Add SKU table rows (limit to first 50 rows)
            for row in sku_df.itertuples(index=False):
                html += "<tr>"
                for col_idx, col in enumerate(sku_df.columns):
                    cell_value = row[col_idx]
                    # Format numbers with Indian formatting if it's the units column
                    if 'unit' in str(col).lower() or 'qty' in str(col).lower():
                        try:
                            cell_value = format_indian(int(cell_value))
                        except:
                            pass
                    html += f"<td>{cell_value}</td>"
                html += "</tr>\n"


            html += "</table>"

        # Add footer
        html += """
<div class="footer" style="background: linear-gradient(to right, #f8f9fa, #e9ecef); padding: 25px; border-radius: 10px; margin-top: 30px; border-left: 5px solid #4472C4;">
    <div style="text-align: center;">
        <div style="display: inline-block; background: #4472C4; color: white; padding: 8px 20px; border-radius: 20px; margin-bottom: 15px;">
            <span style="font-size: 13px; font-weight: bold;">📊 PO Report Generator v1.0</span>
        </div>
        <p style="margin: 10px 0; font-size: 11px; color: #666; font-style: italic;">
            Automated Purchase Order Intelligence System for E-commerce Marketplaces
        </p>
        <div style="background: white; padding: 15px; border-radius: 8px; margin: 15px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
            <p style="margin: 5px 0; font-size: 11px; color: #333;">
                <span style="background: #70AD47; color: white; padding: 4px 12px; border-radius: 12px; font-size: 10px; margin-right: 8px;">
                    👨‍💻 DEVELOPER
                </span>
                <strong>Abhishek Wagh</strong>
            </p>
            <p style="margin: 5px 0; font-size: 10px; color: #666;">
                🆔 Owner ID: RENEE-723 &nbsp;•&nbsp; 📧 abhishek.wagh@reneecosmetics.in
            </p>
        </div>
        <p style="margin: 5px 0; font-size: 9px; color: #999;">
            © 2026 RENEE Cosmetics Pvt. Ltd. | All Rights Reserved | Confidential
        </p>
    </div>
</div>
</body>
</html>
"""

        return html
