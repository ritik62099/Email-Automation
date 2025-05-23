import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load Excel file
file_path = "email.xlsx"  # Aapka Excel file ka path
df = pd.read_excel(file_path)

# Email credentials
EMAIL_ADDRESS = "ritik841207@gmail.com"  # Apna email daalein
EMAIL_PASSWORD = "ixhu osjt icng emid"  # Gmail app password use karein

# SMTP Server setup
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Function to send email
def send_email(to_email, subject, message):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))

        # Connect to SMTP server
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()

        print(f"Email sent successfully to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

# Iterate over Excel emails and send messages
subject = "Your Subject Here"
message_body = "Hello,\n\nThis is a test email sent via Python.\n\nRegards,\nYour Name"

for index, row in df.iterrows():
    email = row['Email']  # Ensure 'Email' column exists in Excel
    send_email(email, subject, message_body)
