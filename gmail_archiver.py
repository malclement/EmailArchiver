#!/usr/bin/env python3
"""
Gmail Archiver

This script downloads all emails from a Gmail account and saves them in an organized
local archive structure. Each email is saved in .eml format, preserving all metadata,
attachments, and content.

Requirements:
- Python 3.6+
- imaplib, email, tqdm libraries

Usage:
1. Enable IMAP in your Gmail settings
2. Allow less secure apps OR set up an app password (recommended)
(https://myaccount.google.com/apppasswords)
3. Run the script and provide your Gmail credentials
"""

import os
import imaplib
import email
import email.policy
import time
import getpass
from datetime import datetime
from tqdm import tqdm
import argparse
import re
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("gmail_archiver.log"),
        logging.StreamHandler()
    ]
)

def sanitize_filename(filename):
    """Clean up filename to be safe for all operating systems."""
    # Replace invalid characters with underscore
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

def get_email_date(email_message):
    """Extract and format the date from an email."""
    date_tuple = email.utils.parsedate_tz(email_message.get('Date'))
    if date_tuple:
        local_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
        return local_date.strftime("%Y-%m-%d_%H-%M-%S")
    return "unknown_date"

def get_email_folder_path(email_message, base_path):
    """Create a folder path based on email date."""
    date_tuple = email.utils.parsedate_tz(email_message.get('Date'))
    if date_tuple:
        local_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
        year_month = local_date.strftime("%Y-%m")
        return os.path.join(base_path, year_month)
    return os.path.join(base_path, "unknown_date")

def get_unique_filename(email_message, folder_path):
    """Generate a unique filename for the email."""
    date = get_email_date(email_message)
    subject = email_message.get('Subject', 'No Subject')
    subject = sanitize_filename(subject)
    from_addr = email_message.get('From', 'unknown@sender.com')
    from_addr = sanitize_filename(from_addr.split('<')[-1].split('>')[0] if '<' in from_addr else from_addr)

    # Create base filename
    base_filename = f"{date}_{from_addr}_{subject[:50]}"
    if len(base_filename) > 120:  # Keep filename length reasonable
        base_filename = base_filename[:120]

    # Ensure filename is unique
    filename = f"{base_filename}.eml"
    counter = 1
    while os.path.exists(os.path.join(folder_path, filename)):
        filename = f"{base_filename}_{counter}.eml"
        counter += 1

    return filename

def save_attachments(email_message, base_path):
    """Save email attachments to a subfolder."""
    attachments_path = None

    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue

        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        if not filename:
            continue

        # Create attachments folder if needed
        if not attachments_path:
            date = get_email_date(email_message)
            subject = sanitize_filename(email_message.get('Subject', 'No Subject'))[:30]
            attach_folder_name = f"{date}_{subject}_attachments"
            attachments_path = os.path.join(base_path, 'attachments', attach_folder_name)
            os.makedirs(attachments_path, exist_ok=True)

        # Save attachment
        filename = sanitize_filename(filename)
        filepath = os.path.join(attachments_path, filename)

        # Handle duplicate filenames
        counter = 1
        while os.path.exists(filepath):
            name, ext = os.path.splitext(filename)
            filepath = os.path.join(attachments_path, f"{name}_{counter}{ext}")
            counter += 1

        with open(filepath, 'wb') as f:
            f.write(part.get_payload(decode=True))

        logging.info(f"Saved attachment: {filepath}")

def archive_gmail(username, password, output_dir, batch_size=100, include_folders=None):
    """
Download and archive all emails from Gmail.

Args:
username: Gmail username
password: Gmail password or app password
output_dir: Directory to save emails
batch_size: Number of emails to process in each batch
include_folders: List of folders to include (default: all folders)
    """
    try:
        # Connect to Gmail using IMAP
        logging.info("Connecting to Gmail...")
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(username, password)

        # Create main output directory
        os.makedirs(output_dir, exist_ok=True)

        # Get list of all folders (labels in Gmail)
        status, folder_list = mail.list()

        if status != 'OK':
            logging.error("Failed to retrieve folder list")
            return

        # Process each folder
        for folder_info in folder_list:
            folder_name = folder_info.decode().split('"/"')[-1].strip().strip('"')

            # Skip folders if include_folders is specified
            if include_folders and folder_name not in include_folders:
                continue

            # Some folder names might contain special characters, sanitize for file system
            safe_folder_name = sanitize_filename(folder_name)
            folder_path = os.path.join(output_dir, safe_folder_name)
            os.makedirs(folder_path, exist_ok=True)

            try:
                # Select the folder
                logging.info(f"Processing folder: {folder_name}")
                status, select_data = mail.select(f'"{folder_name}"', readonly=True)

                if status != 'OK':
                    logging.warning(f"Could not select folder: {folder_name}, skipping...")
                    continue

                # Get total number of emails in the folder
                status, message_count_data = mail.search(None, 'ALL')
                if status != 'OK':
                    logging.warning(f"Could not search folder: {folder_name}, skipping...")
                    continue

                message_numbers = message_count_data[0].split()
                total_messages = len(message_numbers)
                logging.info(f"Found {total_messages} messages in {folder_name}")

                # Process emails in batches to avoid memory issues
                for i in range(0, total_messages, batch_size):
                    batch_end = min(i + batch_size, total_messages)
                    batch_numbers = message_numbers[i:batch_end]

                    logging.info(f"Processing batch {i//batch_size + 1} ({i+1}-{batch_end} of {total_messages})")

                    # Process each email in the batch
                    for msg_index, message_number in enumerate(tqdm(batch_numbers, desc=f"Folder: {safe_folder_name}")):
                        try:
                            # Fetch the email
                            status, msg_data = mail.fetch(message_number, '(RFC822)')

                            if status != 'OK':
                                logging.warning(f"Failed to fetch message {message_number}, skipping...")
                                continue

                            # Parse the email data
                            raw_email = msg_data[0][1]
                            email_message = email.message_from_bytes(raw_email, policy=email.policy.default)

                            # Create year-month subfolder based on email date
                            date_folder = get_email_folder_path(email_message, folder_path)
                            os.makedirs(date_folder, exist_ok=True)

                            # Generate unique filename
                            filename = get_unique_filename(email_message, date_folder)
                            filepath = os.path.join(date_folder, filename)

                            # Save email as .eml file
                            with open(filepath, 'wb') as f:
                                f.write(raw_email)

                            # Save attachments separately if configured
                            save_attachments(email_message, date_folder)

                            # Add a small delay to avoid overwhelming Gmail servers
                            if msg_index % 10 == 0 and msg_index > 0:
                                time.sleep(0.1)

                        except Exception as e:
                            logging.error(f"Error processing message {message_number}: {str(e)}")

                    # Add a delay between batches
                    if batch_end < total_messages:
                        time.sleep(1)

            except Exception as e:
                logging.error(f"Error processing folder {folder_name}: {str(e)}")

            # Close the selected folder before moving to the next one
            mail.close()

        # Logout from Gmail
        mail.logout()
        logging.info("Gmail archiving completed successfully")

    except Exception as e:
        logging.error(f"Error in archiving process: {str(e)}")

def main():
    """Main function to parse arguments and start archiving."""
    parser = argparse.ArgumentParser(description='Archive Gmail emails to local storage.')
    parser.add_argument('--username', help='Gmail username')
    parser.add_argument('--output', default='gmail_archive', help='Output directory (default: gmail_archive)')
    parser.add_argument('--batch-size', type=int, default=100, help='Number of emails to process in each batch (default: 100)')
    parser.add_argument('--folders', nargs='+', help='Specific folders to archive (default: all folders)')

    args = parser.parse_args()

    # Get username if not provided
    username = args.username
    if not username:
        username = input("Enter Gmail address: ")

    # Get password (securely)
    password = getpass.getpass("Enter Gmail password or app password: ")

    # Start archiving
    archive_gmail(username, password, args.output, args.batch_size, args.folders)

if __name__ == "__main__":
    main()