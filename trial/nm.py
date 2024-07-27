import os
import base64
from google.oauth2 import credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import qrcode
from PIL import Image

# Replace with your own values
CLIENT_SECRET_FILE = r"your file directory"
FROM_EMAIL = 'your mail'

# Load credentials from file
flow = InstalledAppFlow.from_client_secrets_file(
    CLIENT_SECRET_FILE, scopes=['https://mail.google.com/']
)
creds = flow.run_local_server(port=0)

# Create the Gmail API client
service = build('gmail', 'v1', credentials=creds)

def generate_qr_code(data, filename):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    img.save(filename)

def send_email(to_email, subject, body, qr_filename):
    # Create a message
    msg = MIMEMultipart()
    msg['From'] = FROM_EMAIL
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach the QR code
    with open(qr_filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {qr_filename}")
        msg.attach(part)

    # Encode the message
    raw_message = base64.urlsafe_b64encode(msg.as_string().encode('utf-8'))

    # Send the message using the Gmail API
    service.users().messages().send(userId='me', body={'raw': raw_message.decode('utf-8')}).execute()

# Load participant data from Excel file
df = pd.read_excel('./participants.xlsx')

for index, row in df.iterrows():
    email = row['Email']
    name = row['Name']
    reg_no = row['RegNo']
    qr_filename = f"{name}_{reg_no}.png"
    qr_data = f"Name: {name}, RegNo: {reg_no}"
    generate_qr_code(qr_data, qr_filename)
    subject = f"Your QR Code for the  Event, {name}"
    body = f"Dear {name},\n\nPlease find your QR code attached. Use it for attendance during the event. See ya there.\n\nBest regards,\n \nEvent Organiser"
    send_email(email, subject, body, qr_filename)
