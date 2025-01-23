import pandas as pd
import os

class SheetReader:
    def __init__(self):
        # Initialize the class by listing all files in the ".attachments" directory
        self.files = os.listdir(".attachments")

    def __call__(self):
        """
        Read all Excel files in the ".attachments" directory and extract relevant data.

        :return: A dictionary where keys are file names and values are lists of tuples containing the extracted data.
        """
        entries = {}
        for f in self.files:
            # Read personal data from the file
            data = self._get_personal_data(self._read_file(f, [0, 1, 2, 4]))

            if "постоянный" in f:
                # If the file name contains "постоянный", read additional period data
                additional_data = self._get_period(self._read_file(f, [0, 1, 2, 4, 9]))
                merged_data = []
                for personal, period in zip(data, additional_data):
                    # Merge personal data with period data
                    merged = list(personal) + [period]
                    merged_data.append(tuple(merged))
                entries[f] = merged_data
            else:
                entries[f] = data

        return entries

    def _read_file(self, f_path, cols):
        """
        Read an Excel file and return a DataFrame with specified columns.

        :param f_path: The path to the Excel file.
        :param cols: A list of column indices to read.
        :return: A DataFrame containing the specified columns.
        """
        return pd.read_excel(f_path, usecols=cols)

    def _get_personal_data(self, df):
        """
        Extract personal data from a DataFrame.

        :param df: The DataFrame containing the data.
        :return: A list of tuples containing personal data.
        """
        result = []
        for index, row in df.iterrows():
            entry = (row["Фамилия"], row["Имя"], row["Отчество"], row["Организация"])
            result.append(entry)
        return result

    def _get_period(self, df):
        """
        Extract period data from a DataFrame.

        :param df: The DataFrame containing the data.
        :return: A list of period data.
        """
        result = []
        key = "Период действия пропуска (дата посещения/дата начала и окончания действия пропуска)"
        for index, row in df.iterrows():
            result.append(row[key])
        return result

# Example usage:
# reader = SheetReader()
# data = reader()
# print(data)
