import os
from dotenv import load_dotenv
import imaplib
import email
from email.header import decode_header
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import io

load_dotenv()

username = os.getenv('EMAIL_USERNAME')
password = os.getenv('EMAIL_PASSWORD')
imap_server = os.getenv('IMAP_SERVER')


# Connect to the server
mail = imaplib.IMAP4_SSL(imap_server)
mail.login(username, password)


mail.select('inbox')


subject = "Here is the attachment"
sender_email = "devil10bro@gmail.com"
result, data = mail.search(None, f'(FROM "{sender_email}" SUBJECT "{subject}")')

# Fetch the email
email_ids = data[0].split()
if email_ids:
    for email_id in email_ids:
        result, message_data = mail.fetch(email_id, '(RFC822)')
        raw_email = message_data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Check if the email has an attachment
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                
                # Check for the attachment and its name
                filename = part.get_filename()
                if filename:
                    # Decode the filename if necessary
                    filename = decode_header(filename)[0][0]
                    if isinstance(filename, bytes):
                        filename = filename.decode()

                    # Check if the file is an Excel file
                    if filename.endswith(('.xls', '.xlsx')):
                        # Read the attachment into a pandas DataFrame
                        attachment = part.get_payload(decode=True)
                        excel_data = pd.read_excel(io.BytesIO(attachment),skiprows=2)
                        
                        # Generate a heatmap of the DataFrame
                        plt.figure(figsize=(8, 6))
                        sns.heatmap(excel_data.corr(), annot=True, cmap='coolwarm')
                        plt.title(f'Heatmap of {filename}')
                        plt.show()
                    else:
                        print("Attachment is there but it's not excel file.")
else:
    print("No emails found with the specified subject and sender.")

# Logout and close the connection
mail.logout()
