import argparse
import os
import shutil
from mail import Mail
from sheetreader import SheetReader
from sheetinserter import SheetInserter

def mkdir(d:str) -> None:
    try:
        os.makedirs(d,exist_ok=True)
        print(f"Directory '{d}' created successfully.")
    except OSError as error:
        print(f"Error creating directory '{d}': {error}")
    
def prepare():
    mkdir(".attachments")
    mkdir("qr_codes")
    
def clear():
    if os.path.exists(".attachments"):
        shutil.rmtree(".attachments")
        print(f"Directory '.attachments' and all its contents have been deleted.")
    else:
        print(f"Directory '.attachments' does not exist.")
        
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="passverify: A tool required to insert data into a remote Google Sheet."
    )

    parser.add_argument('imap_server', type=str, help="IMAP server address")
    parser.add_argument('email', type=str, help="Email address")
    parser.add_argument('appcode', type=str, help="App code")
    parser.add_argument('inbox', type=str, help="Inbox folder")
    parser.add_argument('api_file',type=str,help="file with Google Api fluff")
    parser.add_argument('sheet_name',type=str,help="Google sheet name")
    parser.add_argument('page_number',type=int,help="page number. Start counting with 0, so first is 0.")

    parser.add_argument('--verbose', action='store_true', help="Enable verbose output")

    args = parser.parse_args()

    imap_server = args.imap_server
    email = args.email
    appcode = args.appcode
    inbox = args.inbox
    verbose = args.verbose
    api_file = args.api_file
    sheet_name=args.sheet_name
    
    prepare()
    
    mail = Mail(imap_server,email,appcode,inbox,verbose)
    if mail():
        reader = SheetReader()
        data = reader()

        inserter = SheetInserter(api_file,sheet_name,verbose)
        inserter(data)
        
    clear()
    

