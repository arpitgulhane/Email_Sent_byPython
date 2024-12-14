# import zipfile
#
# import pandas as p
#
# # data = p.read_excel("students.xlsx")
# # print(type(data))
#
# # email_col = data.get("email")
# # print(email_col)
#
# # ---------------------------------------------------------------------------------------
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
# # -------------------------------------------------------------------
# import smtplib
# import schedule
# import time
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# import os
#
#
#
# def compress_file(file_path):
#     """
#     Compresses the file at `file_path` into a ZIP archive.
#     Returns the path to the compressed file.
#     """
#     compressed_file_path = file_path + ".zip"
#     with zipfile.ZipFile(compressed_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
#         zipf.write(file_path, os.path.basename(file_path))
#     print(f"File compressed to: {compressed_file_path}")
#     return compressed_file_path
# # ----------------------------------------------------
# def send_email_with_attachment():
#     # Email configuration
#     print("  Started ")
#     sender_email = "arpitgulhane99@gmail.com"  # Replace with your email
#     # sender_password = " "  # Replace with your app-specific password
#     sender_password = "apon ahzp skrl ujpu"  # Replace with your app-specific password
#     recipient_email = "arpitgulhane99@gmail.com"  # Replace with recipient email
#     subject = "Hourly Call Count Report"
#     body = "Please find the attached report for the hourly call counts."
#
#     # File path to the report
#     # file_path = "C:\\Users\\YourUsername\\Desktop\\CallCountReport.pdf"  # Update with your file path
#     file_path = "C:\\Users\\USER\\Desktop\\TestEmail.pdf"  # Update with your file path
#     if not os.path.exists(file_path):
#         print(f"Error: File not found at {file_path}")
#         return
#
#     # Create email message
#     message = MIMEMultipart()
#     message['From'] = sender_email
#     message['To'] = recipient_email
#     message['Subject'] = subject
#     message.attach(MIMEText(body, 'plain'))
#
#     # # Attach the report
#     # with open(file_path, "rb") as attachment:
#     #     part = MIMEBase('application', 'octet-stream')
#     #     part.set_payload(attachment.read())
#     #     encoders.encode_base64(part)
#     #     part.add_header(
#     #         'Content-Disposition',
#     #         f'attachment; filename={os.path.basename(file_path)}'
#     #     )
#     #     message.attach(part)
#     # ---------------------------------------------------------------
#     compressed_file_path = compress_file(file_path)   # call Zip
#     # -----------------------------------------------------------
#         # Attach the compressed file
#     with open(compressed_file_path, "rb") as attachment:
#         part = MIMEBase('application', 'octet-stream')
#         part.set_payload(attachment.read())
#         encoders.encode_base64(part)
#         part.add_header(
#             'Content-Disposition',
#             f'attachment; filename={os.path.basename(compressed_file_path)}'
#         )
#         message.attach(part)
#
#
#
#     # Send the email
#     try:
#         server = smtplib.SMTP('smtp.gmail.com', 587)
#         server.starttls()
#         server.login(sender_email, sender_password)
#         server.sendmail(sender_email, recipient_email, message.as_string())
#         server.quit()
#         print("Email sent successfully.")
#     except Exception as e:
#         print(f"Error: {e}")
#
#
# # Schedule the job to run every hour
# schedule.every().hour.do(send_email_with_attachment)
# schedule.every(1).minutes.do(send_email_with_attachment)
#
#
#
# print("Scheduler started. Press Ctrl+C to stop.")
# while True:
#     # schedule.run_pending()
#     send_email_with_attachment()   # just to check
#     time.sleep(30)                 # just to check
#

# -------------------------- send by excel ----------------------------------------------
import smtplib
import schedule
import time
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import zipfile

def compress_file(file_path):
    """
    Compresses the file at `file_path` into a ZIP archive.
    Returns the path to the compressed file.
    """
    compressed_file_path = file_path + ".zip"
    with zipfile.ZipFile(compressed_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(file_path, os.path.basename(file_path))
    print(f"File compressed to: {compressed_file_path}")
    return compressed_file_path


def send_email_with_attachment():
    # Email configuration
    print("Started")
    sender_email = "arpitgulhane99@gmail.com"  # Replace with your email
    sender_password = "apon ahzp skrl ujpu"  # Replace with your app-specific password
    subject = "Hourly Call Count Report"
    body = "Please find the attached report for the hourly call counts."

    # File path to the report
    file_path = "C:\\Users\\USER\\Desktop\\TestEmail.pdf"  # Update with your file path
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    # Read recipient emails from Excel file
    excel_path = "C:\\Users\\USER\\Desktop\\recipients.xlsx"  # Update with your Excel file path
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found at {excel_path}")
        return

    try:
        recipients_df = pd.read_excel(excel_path)
        recipient_emails = recipients_df['Email'].dropna().tolist()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Compress the file
    compressed_file_path = compress_file(file_path)

    # Create email message
    for recipient_email in recipient_emails:
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        # Attach the compressed file
        with open(compressed_file_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename={os.path.basename(compressed_file_path)}'
            )
            message.attach(part)

        # Send the email
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
            server.quit()
            print(f"Email sent successfully to {recipient_email}.")
        except Exception as e:
            print(f"Error sending email to {recipient_email}: {e}")


# Schedule the job to run every hour
schedule.every().hour.do(send_email_with_attachment)
schedule.every(1).minutes.do(send_email_with_attachment)

print("Scheduler started. Press Ctrl+C to stop.")
while True:
    # schedule.run_pending()
    send_email_with_attachment()  # Just to check
    time.sleep(30)  # Just to check




