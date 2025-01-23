#!/usr/bin/env python
import argparse
import os
import shutil
from mail import Mail
from sheetreader import SheetReader
from sheetinserter import SheetInserter

def mkdir(d: str) -> None:
    """
    Create a directory if it does not exist.

    :param d: Directory path to create.
    """
    try:
        os.makedirs(d, exist_ok=True)
        print(f"Directory '{d}' created successfully.")
    except OSError as error:
        print(f"Error creating directory '{d}': {error}")

def prepare() -> None:
    """
    Prepare the necessary directories for the script.
    """
    mkdir(".attachments")
    mkdir("qr_codes")

def clear() -> None:
    """
    Clear the '.attachments' directory and all its contents.
    """
    if os.path.exists(".attachments"):
        shutil.rmtree(".attachments")
        print(f"Directory '.attachments' and all its contents have been deleted.")
    else:
        print(f"Directory '.attachments' does not exist.")

if __name__ == "__main__":
    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="passverify: A tool required to insert data into a remote Google Sheet."
    )

    # Define command-line arguments
    parser.add_argument('imap_server', type=str, help="IMAP server address")
    parser.add_argument('email', type=str, help="Email address")
    parser.add_argument('appcode', type=str, help="App code")
    parser.add_argument('inbox', type=str, help="Inbox folder")
    parser.add_argument('api_file', type=str, help="File with Google API credentials")
    parser.add_argument('sheet_name', type=str, help="Google sheet name")
    parser.add_argument('page_number', type=int, help="Page number. Start counting with 0, so first is 0.")
    parser.add_argument('--verbose', action='store_true', help="Enable verbose output")

    # Parse the command-line arguments
    args = parser.parse_args()

    # Extract arguments
    imap_server = args.imap_server
    email = args.email
    appcode = args.appcode
    inbox = args.inbox
    verbose = args.verbose
    api_file = args.api_file
    sheet_name = args.sheet_name
    page_number = args.page_number

    # Prepare the necessary directories
    prepare()

    # Initialize the Mail class to fetch emails
    mail = Mail(imap_server, email, appcode, inbox, verbose)

    # Fetch and process emails
    if mail():
        # Initialize the SheetReader class to read data from the attachments
        reader = SheetReader()
        data = reader()

        # Initialize the SheetInserter class to insert data into the Google Sheet
        inserter = SheetInserter(api_file, sheet_name, page_number, verbose)
        inserter(data)

    # Clear the '.attachments' directory
    clear()
