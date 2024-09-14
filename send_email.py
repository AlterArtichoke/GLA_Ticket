import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Load the Excel file containing email addresses and registration numbers
excel_file = 'email_details.xlsx'  # Ensure this file has columns 'Roll Number' and 'Email ID'
df = pd.read_excel(excel_file)

# Gmail account credentials
gmail_user = 'krish.tejwani2022@vitstudent.ac.in'
gmail_password = 'nmqo ghup fgol zdcr'  # If 2-factor auth is enabled, use app-specific password

# Email settings
subject = "Welcome to the Event!"
body = """Dear Participant,

Thank you for registering for the event! Please find attached your ticket.

We look forward to seeing you!

Best regards,
Event Team
"""

# Create the "Tickets" folder if it doesn't exist
output_folder = 'Tickets'

# Function to send an email with an image attachment
def send_email_with_ticket(roll_number, email, ticket_image):
    # Set up the MIME message
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = email
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Attach the ticket image
    attachment = open(ticket_image, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(ticket_image)}')
    msg.attach(part)
    attachment.close()

    # Send the email via Gmail's SMTP server
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_user, gmail_password)
        server.send_message(msg)
        server.quit()
        print(f"Email sent successfully to {email}")
    except Exception as e:
        print(f"Failed to send email to {email}. Error: {str(e)}")

# Send emails to each participant
for _, row in df.iterrows():
    roll_number = row['Roll Number']
    email = row['Email ID']
    
    # Find the corresponding ticket image for the roll number
    ticket_image = os.path.join(output_folder, f"{roll_number}_ticket.png")

    if os.path.exists(ticket_image):
        send_email_with_ticket(roll_number, email, ticket_image)
    else:
        print(f"Ticket image for Roll Number {roll_number} not found.")

print("All emails sent.")
