import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# ---------------------------------------
# --- CHANGE THE FOLLOWING SETTINGS ---
# ---------------------------------------
sender_email = "XXXX@gmail.com"
sender_password = "XXXX XXXX XXXX XXXX"
recipients = ["XXXX@XXX.com", "XXXX@XXX.com"]
cc = ["XXXX@XXX.com"]
subject = "Hello from Python"
body = """
<html>
  <body>
    <p>This is an <b>HTML</b> email sent from Python using the Gmail SMTP server.</p>
  </body>
</html>
"""
attachments = ["attachment.txt"]
# ---------------------------------------


# Create email with multiple parts (text, attachments, etc.)
message = MIMEMultipart()
message["Subject"] = subject
message["From"] = sender_email
message["To"] = ", ".join(recipients)
if cc:
    message["Cc"] = ", ".join(cc)

# Add HTML content for better formatting in email
message.attach(MIMEText(body, "html"))

# Attach files, ensuring the paths are valid
for attachment_path in attachments:
    file_path = Path(attachment_path)
    if file_path.is_file():  # Skip if the file doesn't exist
        with file_path.open("rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())  # Read and prepare file for email
            encoders.encode_base64(part)  # Encode file for email transmission
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={file_path.name}",  # Use the correct file name
            )
            message.attach(part)
    else:
        print(f"Attachment not found: {file_path}")  # Notify about missing files

# Combine all recipients (To + CC) for sending
all_recipients = recipients + cc

# Connect to Gmail SMTP server and send the email
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(sender_email, sender_password)  # Authenticate with Gmail
    server.sendmail(sender_email, all_recipients, message.as_string())  # Send email
