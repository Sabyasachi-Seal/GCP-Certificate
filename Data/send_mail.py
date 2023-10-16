import openpyxl
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Replace with your email configuration
email_address = input("Enter your gmail id: ")
email_password = input("Enter your password: ")
smtp_server = 'smtp.gmail.com'
smtp_port = 587

# Load the Excel file
file_path = 'Data/Mail.xlsm'
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook.active

# Find the column indices
to_column = 'A'
cc_column = 'B'
subject_column = 'C'
body_column = 'D'
attachment_column = 'E'

is_test = int(input("\nIs this for testing?\n1. Yes  2. No\n=> "))

print("\nStarting to send the emails...\n")

# Iterate through the rows in the Excel file
for row in worksheet.iter_rows(min_row=2, values_only=True):  # Assuming data starts from the second row
    if is_test == 1:
        to_address = input("Enter to email: ")
        cc_address = input("Enter cc emails: ")
        subject = "[TEST] Google Cloud Certificate"
        body = row[3]
        attachment_path = row[4]
    elif is_test == 2:
        to_address = row[0]
        cc_address = row[1]
        subject = row[2]
        body = row[3]
        attachment_path = row[4]
    else:
        print("\nWrong input, enter either 1 or 2\n")
        break
    # Create an email message
    message = MIMEMultipart()
    message['From'] = email_address
    message['To'] = to_address
    message['Cc'] = cc_address
    message['Subject'] = subject
    message.attach(MIMEText(body, 'html'))

    # Attach the file (if any)
    if attachment_path:
        attachment_path = os.path.normpath(attachment_path)
        attachment = MIMEApplication(open(attachment_path, 'rb').read())
        attachment.add_header('Content-Disposition', 'attachment', filename='GoogleCloudCertificate.pdf')
        message.attach(attachment)

    # Connect to the SMTP server and send the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email_address, email_password)
        recipients = [to_address]
        server.sendmail(email_address, recipients, message.as_string())
        server.quit()
        print(f"Email sent to {to_address}")
    except Exception as e:
        print(f"Failed to send email to {to_address}: {str(e)}")
    finally:
        if is_test == 1:
            break

# Close the Excel file
workbook.close()
