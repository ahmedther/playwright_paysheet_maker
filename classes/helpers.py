import os
import re
import pickle
from datetime import datetime, timedelta
import win32com.client as win32


class Helpers:
    def __init__(self):
        pass

    def extract_client_name(self, reference):
        match = re.search(
            r"(?:(?:Mrs\.|Miss|Mr\.?|Ms\.?)\s)?([A-Za-z]+(?:\s[A-Za-z]+)?)\s([A-Za-z']+)",
            reference,
        )
        if match:
            first_name = match.group(1).split()[
                0
            ]  # Take only the first part of the first name
            last_name = match.group(2)

            # Special handling for names with apostrophes
            if "'" in last_name:
                return f"{first_name[0]} {last_name[:5]}"
            else:
                return f"{first_name[0]} {last_name[:3]}"
        return None

    def save_data(self, data, filename):
        with open(filename, "wb") as file:
            pickle.dump(data, file)

    # Load data from a file
    def load_data(self, filename):
        with open(filename, "rb") as file:
            return pickle.load(file)

    def get_working_dates(self):

        # Get today's date
        today = datetime.now()

        # Calculate the date of this week's Saturday
        to_date = today + timedelta((5 - today.weekday()) % 7)
        to_date_str = to_date.strftime("%d-%b-%Y")

        # Calculate the date of last week's Monday
        from_date = today - timedelta(days=today.weekday() + 7)
        from_date_str = from_date.strftime("%d-%b-%Y")

        return from_date_str, to_date_str

    def get_file_name(self, file_path, folder_name):

        from_date_str, to_date_str = self.get_working_dates()

        folder_path = os.path.join(
            os.path.dirname(file_path), "generated_files", folder_name
        )

        file_name = f"Paysheet {from_date_str} - {to_date_str}"

        return os.path.join(folder_path, file_name)

    def convert_excel_to_pdf(self, excel_path, pdf_path):
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(excel_path)
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close()
        excel.Quit()
