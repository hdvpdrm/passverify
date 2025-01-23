import gspread
import random
import time
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from qrcodegen import QRcodeGen


class SheetInserter:
    def __init__(self,cred_file:str,sheet_name:str, verbose:bool):
        self.scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, self.scope)
        self.client = gspread.authorize(self.creds)
        self.sheet = self.client.open(sheet_name)
        self.sheet_instance = self.sheet.get_worksheet(1)

        self.verbose = verbose
        
        self.id = 10010
        self.cols = {
            "date":      0,
            "lastname":  1,
            "name":      2,
            "secondname":3,
            "org":       4,
            "id":        9,
            "range":     8
            }
        
    def __call__(self,data:dict):
        start = self._get_last_free_row()

        for key in data.keys():
            if self.verbose: print("starting info from '{}'".format(key))
            for entry in data[key]:
                self._insert(entry,start)
                if self.verbose: print("Inserted entry:",entry)
                start+=1


    def rc_to_a1(self,row, col):
        import string
        letters = string.ascii_uppercase
        col_label = ''
        while col >= 0:
            col_label = letters[col % 26] + col_label
            col = col // 26 - 1
        return "{}{}".format(col_label,row)
                
    def _insert(self,entry:tuple, row:int) -> None:
        today = datetime.today().strftime('%d-%b-%Y')
        data = {
            "date":       today,
            "lastname":   entry[0],
            "name":       entry[1],
            "secondname": entry[2],
            "org":        entry[3],
            "id":         self.id,
            }

        if len(entry) == 5:
            data["range"] = entry[4]

        for key in data.keys():
            col = self.cols[key]
            value = data[key]
            pos = self.rc_to_a1(row,col)
            self._update_with_backoff(pos,[[value]])
            qr = QRcodeGen()
            qr(str(self.id),str(self.id)+".png")
            
            if self.verbose: print("created qr for '{}' id!".format(self.id))
            
        self.id+=1

    def _update_with_backoff(self, pos, value):
        retries = 5
        for attempt in range(retries):
            try:
                self.sheet_instance.update(pos, value)
                break
            except gspread.exceptions.APIError as e:
                if "Quota exceeded" in str(e):
                    wait_time = (2 ** attempt) + random.random()
                    print(f"Quota exceeded. Retrying in {wait_time:.2f} seconds...")
                    time.sleep(wait_time)
                else:
                    raise e

    def _get_last_free_row(self) -> int:
        pred = lambda x: not all(el == '' for el in x)
        all_values = list(filter(pred,self.sheet_instance.get_all_values()))
        return len(all_values)

        
        
