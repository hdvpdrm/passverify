import gspread
import random
import time
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from qrcodegen import QRcodeGen

class SheetInserter:
    def __init__(self, cred_file: str, sheet_name: str, sheet_page: int, verbose: bool):
        """
        Initialize the SheetInserter class with Google Sheets credentials and sheet details.

        :param cred_file: Path to the credentials JSON file.
        :param sheet_name: Name of the Google Sheet.
        :param sheet_page: Index of the sheet page (worksheet) to use.
        :param verbose: Boolean flag to enable verbose output.
        """
        # Define the scope of the Google Sheets API
        self.scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

        # Authenticate using the service account credentials
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, self.scope)

        # Authorize the client to interact with the Google Sheets API
        self.client = gspread.authorize(self.creds)

        # Open the specified Google Sheet
        self.sheet = self.client.open(sheet_name)

        # Get the specified worksheet within the Google Sheet
        self.sheet_instance = self.sheet.get_worksheet(sheet_page)

        # Verbose flag for detailed output
        self.verbose = verbose

        # Initialize the ID counter
        self.id = 10010

        # Define the column indices for various data fields
        self.cols = {
            "date": 0,
            "lastname": 1,
            "name": 2,
            "secondname": 3,
            "org": 4,
            "id": 9,
            "range": 8
        }

    def __call__(self, data: dict):
        """
        Insert data into the Google Sheet.

        :param data: Dictionary containing data to be inserted.
        """
        # Get the last free row in the sheet
        start = self._get_last_free_row()

        # Iterate over the data dictionary
        for key in data.keys():
            if self.verbose:
                print("starting info from '{}'".format(key))
            for entry in data[key]:
                # Insert each entry into the sheet
                self._insert(entry, start)
                if self.verbose:
                    print("Inserted entry:", entry)
                start += 1

    def rc_to_a1(self, row: int, col: int) -> str:
        """
        Convert row and column indices to A1 notation.

        :param row: Row index.
        :param col: Column index.
        :return: A1 notation string.
        """
        import string
        letters = string.ascii_uppercase
        col_label = ''
        while col >= 0:
            col_label = letters[col % 26] + col_label
            col = col // 26 - 1
        return "{}{}".format(col_label, row)

    def _insert(self, entry: tuple, row: int) -> None:
        """
        Insert a single entry into the Google Sheet.

        :param entry: Tuple containing the data to be inserted.
        :param row: Row index where the data will be inserted.
        """
        # Get today's date in the desired format
        today = datetime.today().strftime('%d-%b-%Y')

        # Prepare the data dictionary for the entry
        data = {
            "date": today,
            "lastname": entry[0],
            "name": entry[1],
            "secondname": entry[2],
            "org": entry[3],
            "id": self.id,
        }

        # If the entry has a range, include it in the data
        if len(entry) == 5:
            data["range"] = entry[4]

        # Insert each field into the corresponding cell
        for key in data.keys():
            col = self.cols[key]
            value = data[key]
            pos = self.rc_to_a1(row, col)
            self._update_with_backoff(pos, [[value]])

            # Generate a QR code for the ID
        qr = QRcodeGen()
        qr(str(self.id), "qr_codes/" + data["lastname"] + ".png")

        if self.verbose:
            print("created qr for '{}' id!".format(self.id))
        self.id += 1

    def _update_with_backoff(self, pos: str, value: list) -> None:
        """
        Update a cell in the Google Sheet with exponential backoff for quota exceeded errors.

        :param pos: Cell position in A1 notation.
        :param value: Value to be inserted into the cell.
        """
        retries = 5
        for attempt in range(retries):
            try:
                # Update the cell with the given value
                self.sheet_instance.update(pos, value)
                break
            except gspread.exceptions.APIError as e:
                if "Quota exceeded" in str(e):
                    # If quota exceeded, wait and retry
                    wait_time = (2 ** attempt) + random.random()
                    print(f"Quota exceeded. Retrying in {wait_time:.2f} seconds...")
                    time.sleep(wait_time)
                else:
                    # Raise any other exceptions
                    raise e

    def _get_last_free_row(self) -> int:
        """
        Get the index of the last free row in the Google Sheet.

        :return: Index of the last free row.
        """
        # Filter out empty rows
        pred = lambda x: not all(el == '' for el in x)
        all_values = list(filter(pred, self.sheet_instance.get_all_values()))
        return len(all_values)

# Example usage:
# inserter = SheetInserter('path/to/credentials.json', 'SheetName', 0, True)
# data = {
#     'file1': [('Lastname1', 'Name1', 'Secondname1', 'Org1'), ('Lastname2', 'Name2', 'Secondname2', 'Org2')],
#     'file2': [('Lastname3', 'Name3', 'Secondname3', 'Org3', 'Range1')]
# }
# inserter(data)
