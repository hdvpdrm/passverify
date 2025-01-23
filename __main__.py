import os
import shutil
from mail import Mail

def prepare():
    try:
        os.makedirs(".attachments",exist_ok=True)
        print("Directory '.attachments' created successfully.")
    except OSError as error:
        print(f"Error creating directory '.attachments': {error}")
def clear():
    if os.path.exists(".attachments"):
        shutil.rmtree(".attachments")
        print(f"Directory '.attachments' and all its contents have been deleted.")
    else:
        print(f"Directory '.attachments' does not exist.")
        
if __name__ == "__main__":
    prepare()
    
    mail = Mail("imap.mail.ru","maganer.krik@mail.ru","bYZt59p2YTaxT5YHTPF0","INBOX/ToMyself")

    clear()
    

