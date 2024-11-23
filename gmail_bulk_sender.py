"""
===========================================
Excel Gmail Bulk Sender with Python 📧
===========================================

Easily send personalized bulk emails directly from Excel using Python and Gmail.
This tool simplifies email campaigns with features like placeholders, attachments, 
and real-time delivery status.

Looking for a VBA-based tool with more advanced features? 
Try the Gmail Excel Blaster: https://pythonandvba.com/gmail-excel-blaster

Created by Sven Bosau | PythonAndVBA.com
Explore more solutions: https://pythonandvba.com/solutions
YouTube: https://youtube.com/@codingisfun
"""

import xlwings as xw  # pip install xlwings
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
import re
from datetime import datetime
import webbrowser

# Configuration
PLACEHOLDER_COUNT = 7  # Adjust the number of placeholders here
MAX_ATTACHMENT_SIZE_MB = 25  # Maximum attachment size in MB


def validate_email(email):
    """Validate an email address using regex."""
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email.strip()) is not None


def send_email(sender, password, recipients, subject, body, cc=None, attachments=None):
    """Send an email with optional CC and attachments."""
    try:
        message = MIMEMultipart()
        message["From"] = sender
        message["To"] = ", ".join(recipients)
        if cc:
            message["Cc"] = ", ".join(cc)
        message["Subject"] = subject

        # Attach HTML body
        message.attach(MIMEText(body, "html"))

        # Attach files
        if attachments:
            for file_path in attachments:
                path = Path(file_path.strip())
                if path.exists():
                    if path.stat().st_size > MAX_ATTACHMENT_SIZE_MB * 1024 * 1024:
                        raise ValueError(f"Attachment too large: {path.name}")
                    with path.open("rb") as file:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(file.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename={path.name}",
                        )
                        message.attach(part)
                else:
                    raise FileNotFoundError(f"Attachment not found: {path}")

        # Send email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(
                sender, recipients + (cc if cc else []), message.as_string()
            )

        return True
    except Exception as e:
        raise e


def main():
    # Open workbook and sheets
    wb = xw.Book.caller()
    send_list_sheet = wb.sheets["SEND_LIST"]

    # Get table object
    table = send_list_sheet.tables["tblSendList"]

    # Get data and headers
    headers = table.header_row_range.value
    rows = table.data_body_range.value

    # Get sender credentials
    sender_email = send_list_sheet["SenderEmail"].value
    sender_password = send_list_sheet["SenderPassword"].value

    # Ensure credentials are not None and strip whitespace
    sender_email = sender_email.strip() if sender_email else None
    sender_password = sender_password.strip() if sender_password else None

    # Check sender email and password
    if not sender_email or not sender_password:
        wb.app.alert(
            prompt=(
                "Sender email or app password is missing.\n"
                "To set up a Gmail app password, please follow this guide:\n"
                "https://docs.pythonandvba.com/gmail-blaster/guides/setting-up-an-app-password-in-gmail"
            ),
            title="Missing Credentials",
            mode="critical",
        )
        return

    # Get email body
    email_body_sheet = wb.sheets["EMAIL_BODY"]
    email_body = email_body_sheet.range("EmailBody").value

    # Ensure email body is not empty
    if not email_body:
        wb.app.alert(
            prompt="The Email Body is missing. Please fill in the Email Body before sending.",
            title="Missing Email Body",
            mode="critical",
        )
        return

    # Clear the Status column
    status_col_index = headers.index("Status") + 1  # +1 as Excel index starts from 1
    for row_index, row in enumerate(rows):
        if row[headers.index("Receiver")]:
            send_list_sheet.range(
                (row_index + table.data_body_range.row, status_col_index + 1)
            ).value = None

    sent_count = 0  # Track number of successfully sent emails

    # Iterate over rows with non-empty Receiver
    for row_index, row in enumerate(rows):
        if not row[headers.index("Receiver")]:
            continue

        row_data = dict(zip(headers, row))  # Create a dictionary of header-value pairs

        try:
            # Parse row data
            recipients = [
                email.strip()
                for email in row_data["Receiver"].split(",")
                if email.strip()
            ]
            cc_list = [
                email.strip()
                for email in (row_data.get("CC", "").split(","))
                if email.strip()
            ]
            attachment_list = [
                file.strip()
                for file in (row_data.get("Attachment(s)", "").split(","))
                if file.strip()
            ]
            subject = row_data["Subject"] or ""  # Subject can be empty
            placeholders = {
                f"{{{{Placeholder{i+1}}}}}": row_data.get(f"Placeholder{i+1}", "")
                for i in range(PLACEHOLDER_COUNT)
            }

            # Validate recipients
            if not all(validate_email(email) for email in recipients):
                raise ValueError("Invalid recipient email address(es).")

            # Check attachments
            invalid_attachments = [
                file for file in attachment_list if not Path(file).exists()
            ]
            if invalid_attachments:
                raise FileNotFoundError(
                    f"Attachments not found: {', '.join(invalid_attachments)}"
                )

            # Replace placeholders in the email body
            email_content = email_body
            for placeholder, value in placeholders.items():
                email_content = email_content.replace(placeholder, value or "")

            # Send email
            success = send_email(
                sender_email,
                sender_password,
                recipients,
                subject,
                email_content,
                cc_list,
                attachment_list,
            )

            # Update status with timestamp
            if success:
                sent_count += 1
                status_value = f"Sent - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

        except Exception as e:
            # Log failure with reason
            status_value = f"Failed - {str(e)}"

        # Update the status column
        send_list_sheet.range(
            (row_index + table.data_body_range.row, status_col_index + 1)
        ).value = status_value

    # Inform user via MsgBox
    wb.app.alert(
        prompt=f"Task completed.\n\nEmails sent successfully: {sent_count}\n\nCheck the Status column for details.",
        title="Status",
        mode="info",
    )

    # Offer to explore the advanced Gmail Bulk Sender
    button_value = wb.app.alert(
        prompt="Task completed successfully. If you're looking for even more advanced features, consider exploring the advanced Gmail Bulk Sender.\nWould you like to learn more?",
        title="Advanced Gmail Bulk Sender",
        buttons="yes_no",
        mode="info",
    )
    if button_value == "yes":
        webbrowser.open("https://pythonandvba.com/gmail-excel-blaster")


if __name__ == "__main__":
    xw.Book("gmail_bulk_sender.xlsm").set_mock_caller()
    main()
