import imaplib
import email
from email.policy import default
from datetime import datetime, timedelta
import os


class Mail:
    def __init__(self,server,username,password,inbox="inbox", verbose=True):
        self.output_dir = ".attachments"

        self.mail = imaplib.IMAP4_SSL(server)
        self.mail.login(username, password)

        self.mail.select(inbox)

        self.verbose = verbose
        self._saved_any = False

    def __call__(self):
        status = self._fetch_emails()
        return status and self._saved_any
        
    def _fetch_emails(self) -> bool:
        today = datetime.today().strftime('%d-%b-%Y')

        status, messages = self.mail.search(None, f'(SENTSINCE {today})')
        if status != 'OK':
            print("No messages found!")
            return False

        if self.verbose: print("fetched messages: status: {}".format(status))

        splitted = messages[0].split()
        if not splitted and self.verbose: print("no suitable messages found!")
        
        for num in splitted:            
            status, data = self.mail.fetch(num, '(BODY.PEEK[])')
            if status != 'OK':
                print(f"Failed to fetch email {num}")
                continue

            if self.verbose: print("parsed mail: status: {}".format(status))

            msg = email.message_from_bytes(data[0][1], policy=default)
            if self._save_attachment(msg): self._saved_any = True
        return True


    def _check_fname(self,name:str) -> bool:
        check1 = "разовый пропуск" in name
        check2 = "постоянный пропуск" in name
        return check1 or check2
    
    def _save_attachment(self,msg) -> bool:
        if not msg.is_multipart():
            if self.verbose: print("skip attachment.")
            return None
        
        for part in msg.iter_attachments():
            filename = part.get_filename()
            if not self._check_fname(filename):
                if self.verbose: print("skip '{}' attachment".format(filename))
                return None
                
            print("writing '{}' file...".format(filename))
            with open(os.path.join(self.output_dir, filename), 'wb') as f:
                f.write(part.get_payload(decode=True))
                print(f"Saved attachment: '{filename}'")

        return True

    
        
