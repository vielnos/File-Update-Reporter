#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import getpass  # for secure password input


def scan_directory(path):
    """Recursively scan the directory and return a list of files with their full paths and creation times."""
    files_info = []
    for root, dirs, files in os.walk(path):
        for file in files:
            full_path = os.path.join(root, file)
            creation_time = datetime.fromtimestamp(os.path.getctime(full_path)).strftime('%Y-%m-%d %H:%M:%S')
            files_info.append((full_path, creation_time))
    return files_info


def update_excel(path, data, report_data):
    """Update or create an Excel file with current and new file data."""
    if not os.path.exists(path):
        # Create new Excel file if it doesn't exist
        with pd.ExcelWriter(path, engine='openpyxl', mode='w') as writer:
            pd.DataFrame(data, columns=['Filename', 'Creation Time']).to_excel(writer, sheet_name='Current Data', index=False)
            pd.DataFrame(report_data, columns=['Filename', 'Creation Time']).to_excel(writer, sheet_name='Report Data', index=False)
    else:
        # Update existing file, replacing sheets
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            pd.DataFrame(data, columns=['Filename', 'Creation Time']).to_excel(writer, sheet_name='Current Data', index=False)
            pd.DataFrame(report_data, columns=['Filename', 'Creation Time']).to_excel(writer, sheet_name='Report Data', index=False)


def send_email(filepath):
    """Send an email with the Excel file as an attachment."""
    sender_email = input("Enter the sender's email address (e.g., your Gmail): ").strip()
    receiver_email = input("Enter the receiver's email address: ").strip()
    password = getpass.getpass("Enter the sender's email password or app password (input hidden): ")

    # Compose the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "Updated File Report"

    # Attach the Excel file
    with open(filepath, "rb") as file:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(filepath)}"')
    msg.attach(part)

    # Send email using Gmail's SMTP server
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        print("\n✅ Email sent successfully.")
    except Exception as e:
        print(f"\n❌ Failed to send email: {e}")


def main():
    print("\n=== 📁 File Updation and Auto Email Script ===\n")
    directory_path = input("Enter the directory path to scan: ").strip()

    if not os.path.exists(directory_path):
        print(f"\n❌ The directory '{directory_path}' does not exist. Please check the path.")
        return

    excel_path = os.path.join(directory_path, 'File Updation.xlsx')

    # Step 1: Scan directory for current file info
    current_files = scan_directory(directory_path)

    # Step 2: Check previous Excel data to find new files
    try:
        old_data = pd.read_excel(excel_path, sheet_name='Current Data')
        old_files = set((row['Filename'], row['Creation Time']) for _, row in old_data.iterrows())
        new_files = [file for file in current_files if file not in old_files]
    except FileNotFoundError:
        # First run — consider all files new
        new_files = current_files

    # Step 3: Update Excel with both current and new files
    update_excel(excel_path, current_files, new_files)
    print(f"\n✅ Excel file updated successfully at: {excel_path}")

    # Step 4: Send updated Excel report by email
    send_email(excel_path)


if __name__ == "__main__":
    main()

