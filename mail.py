import imaplib
import email
from email.policy import default
from datetime import datetime
import os

class Mail:
    def __init__(self, server: str, username: str, password: str, inbox: str = "inbox", verbose: bool = True):
        """
        Initialize the Mail class with IMAP server details and login credentials.

        :param server: IMAP server address.
        :param username: Username for the email account.
        :param password: Password for the email account.
        :param inbox: Mailbox to select (default is "inbox").
        :param verbose: Boolean flag to enable verbose output.
        """
        # Directory to save attachments
        self.output_dir = ".attachments"

        # Connect to the IMAP server
        self.mail = imaplib.IMAP4_SSL(server)

        # Login to the email account
        self.mail.login(username, password)

        # Select the specified mailbox
        self.mail.select(inbox)

        # Verbose flag for detailed output
        self.verbose = verbose

        # Flag to track if any attachments were saved
        self._saved_any = False

    def __call__(self) -> bool:
        """
        Fetch and process emails.

        :return: True if emails were fetched and any attachments were saved, otherwise False.
        """
        # Fetch emails and check if any attachments were saved
        status = self._fetch_emails()
        return status and self._saved_any

    def _fetch_emails(self) -> bool:
        """
        Fetch emails sent since today and process them.

        :return: True if emails were fetched successfully, otherwise False.
        """
        # Get today's date in the required format
        today = datetime.today().strftime('%d-%b-%Y')

        # Search for emails sent since today
        status, messages = self.mail.search(None, f'(SENTSINCE {today})')
        if status != 'OK':
            print("No messages found!")
            return False

        if self.verbose:
            print("fetched messages: status: {}".format(status))

        # Split the message IDs
        splitted = messages[0].split()
        if not splitted and self.verbose:
            print("no suitable messages found!")

        # Process each email
        for num in splitted:
            status, data = self.mail.fetch(num, '(BODY.PEEK[])')
            if status != 'OK':
                print(f"Failed to fetch email {num}")
                continue

            if self.verbose:
                print("parsed mail: status: {}".format(status))

            # Parse the email message
            msg = email.message_from_bytes(data[0][1], policy=default)
            if self._save_attachment(msg):
                self._saved_any = True

            # Mark the email as read
            self.mail.store(num, '+FLAGS', '\\Seen')
        return True

    def _check_fname(self, name: str) -> bool:
        """
        Check if the filename contains specific keywords.

        :param name: Filename to check.
        :return: True if the filename contains the keywords, otherwise False.
        """
        check1 = "разовый пропуск" in name
        check2 = "постоянный пропуск" in name
        return check1 or check2

    def _save_attachment(self, msg) -> bool:
        """
        Save attachments from the email message if they meet the criteria.

        :param msg: Email message object.
        :return: True if any attachments were saved, otherwise False.
        """
        # Check if the message is multipart
        if not msg.is_multipart():
            if self.verbose:
                print("skip attachment.")
            return None

        # Iterate over the attachments
        for part in msg.iter_attachments():
            filename = part.get_filename()
            if not self._check_fname(filename):
                if self.verbose:
                    print("skip '{}' attachment".format(filename))
                return None

            print("writing '{}' file...".format(filename))
            # Save the attachment to the output directory
            with open(os.path.join(self.output_dir, filename), 'wb') as f:
                f.write(part.get_payload(decode=True))
                print(f"Saved attachment: '{filename}'")

        return True

# Example usage:
# mail_client = Mail('imap.example.com', 'username', 'password')
# mail_client()
