import pandas as pd
import gspread
from google.auth import exceptions
from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSheetData:
    def __init__(self, sheet_url, credentials_file="credentials.json"):
        self.sheet_url = sheet_url
        self.credentials_file = credentials_file
        self.Question_Sheet = None
        self.Chemical_Formula_Sheet = None
        self.Symbols_Sheet = None
        self.Superscript_Sheet = None
        self.Options_Sheet = None

    def _authorize_client(self):
        try:
            # Load credentials from the service account key JSON file
            creds = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_file)

            # Authorize the client
            return gspread.authorize(creds)

        except exceptions.GoogleAuthError as e:
            print(f"Authentication failed: {e}")
            return None

    def _get_sheet_data(self, gc, sheet_name):
        try:
            # Open the Google Sheet using its URL
            sh = gc.open_by_url(self.sheet_url)

            # Access the specified sheet
            worksheet = sh.worksheet(sheet_name)
            return pd.DataFrame(worksheet.get_all_records())

        except gspread.exceptions.WorksheetNotFound:
            print(f"Sheet '{sheet_name}' not found in the document.")
            return pd.DataFrame()

    def connect_to_google_sheets(self):
        # Authorize the client
        gc = self._authorize_client()

        if gc:
            # Access Question Sheet
            self.Question_Sheet = self._get_sheet_data(gc, 'question')

            # Access Chemical Formula Sheet
            self.Chemical_Formula_Sheet = self._get_sheet_data(gc, 'Chemical Formula')

            # Access Symbol Replacement Sheet
            self.Symbols_Sheet = self._get_sheet_data(gc, 'Symbols Replacement')

            # Access Superscript Sheet
            self.Superscript_Sheet = self._get_sheet_data(gc, 'Superscript')

            # Access Options Removal Sheet
            self.Options_Sheet = self._get_sheet_data(gc, 'Options Removal')

    def get_question_sheet(self):
        return self.Question_Sheet

    def get_chemical_formula_sheet(self):
        return self.Chemical_Formula_Sheet

    def get_symbols_sheet(self):
        return self.Symbols_Sheet

    def get_superscript_sheet(self):
        return self.Superscript_Sheet

    def get_options_sheet(self):
        return self.Options_Sheet

    # def print_sheet(self, sheet, sheet_name):
    #     print(f"\n{sheet_name} Sheet:")
    #     print(sheet)

# Example usage:
if __name__ == "__main__":
    # Replace 'your_google_sheets_url' with the actual URL of your Google Sheets document
    google_sheets_url = 'https://docs.google.com/spreadsheets/d/1V_gLg2RipDqcg5yJjUsT7gDqkkOYb7G1PuLLMBGw7bg/edit?usp=sharing'

    # Create an instance of the GoogleSheetData class
    sheet_data = GoogleSheetData(google_sheets_url)

    # Connect to Google Sheets and retrieve data
    sheet_data.connect_to_google_sheets()

    # Get individual sheets
    question_sheet = sheet_data.get_question_sheet()
    chemical_formula_sheet = sheet_data.get_chemical_formula_sheet()
    symbols_sheet = sheet_data.get_symbols_sheet()
    superscript_sheet = sheet_data.get_superscript_sheet()
    option_sheet = sheet_data.get_options_sheet()

    # Print or use the sheets as needed
    sheet_data.print_sheet(option_sheet, "Options")
    # sheet_data.print_sheet(question_sheet, "Question")
    # sheet_data.print_sheet(chemical_formula_sheet, "Chemical Formula")
    # sheet_data.print_sheet(symbols_sheet, "Symbols")
    # sheet_data.print_sheet(superscript_sheet, "Superscript")