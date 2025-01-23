import pandas as pd
import os

class SheetReader:
    def __init__(self):
        self.files = os.listdir(".attachments")

    def __call__(self):
        entries = {}
        for f in self.files:
            data = self._get_personal_data(self._read_file(f,[0,1,2,4]))

            if "постоянный" in f:
                additional_data = self._get_period(self._read_file(f,[0,1,2,4,9]))
                merged_data = []
                for personal,period in zip(data,additional_data):
                    merged = list(personal)+[period]
                    merged_data.append(tuple(merged))
                entries[f] = merged_data
            else:
                entries[f] = data

        return entries
        
    def _read_file(self,f_path,cols):
        return pd.read_excel(f_path, usecols=cols)
    def _get_personal_data(self,df):
        result = []
        for index, row in df.iterrows():            
            entry = (row["Фамилия"],row["Имя"],row["Отчество"],row["Организация"])
            result.append(entry)

        return result

    def _get_period(self,df):
        result = []
        key = "Период действия пропуска (дата посещения/дата начала и окончания действия пропуска)"
        for index, row in df.iterrows():
            result.append(row[key])
        return result

        
        
