import pandas as pd
from GoogleSheetData2 import GoogleSheetData
import re

class FindReplace:
    def __init__(self, sheet_data):
        self.sheet_data = sheet_data

    def _prepare_find_replace_data(self, sheet, find_col, replace_col):
        find_and_replace_data = {
            'Find': sheet[find_col].apply(lambda x: re.escape(str(x)) if pd.notnull(x) else ''),
            'Replace': sheet[replace_col]
        }
        return pd.DataFrame(find_and_replace_data)

    def _perform_find_replace(self, find_pattern, replace_pattern, Question_Sheet):
        for col_name in Question_Sheet.columns:
            try:
                contains_find = Question_Sheet[col_name].str.contains(find_pattern, regex=True, case=False)
                if contains_find.any():
                    print(f"Replacing {find_pattern} with {replace_pattern} in {col_name}")
                    Question_Sheet[col_name] = Question_Sheet[col_name].str.replace(find_pattern, replace_pattern, regex=True, flags=re.IGNORECASE)
            except Exception as e:
                print(f"Error performing find and replace: {e}")

    def replace_in_sheet(self, sheet, find_col, replace_col):
        Question_Sheet = self.sheet_data.get_question_sheet()
        find_df = self._prepare_find_replace_data(sheet, find_col, replace_col)

        # Perform the find and replace operation with case-insensitivity
        for find_pattern, replace_pattern in zip(find_df['Find'], find_df['Replace']):
            try:
                if find_pattern:
                    self._perform_find_replace(find_pattern, replace_pattern, Question_Sheet)
            except Exception as e:
                print(f"Error in replace_in_sheet: {e}")

    def replace_in_question_sheet(self):
        Chemical_Formula_Sheet = self.sheet_data.get_chemical_formula_sheet()
        Symbols_Sheet = self.sheet_data.get_symbols_sheet()
        Superscrip_Sheet = self.sheet_data.get_superscript_sheet()
        Options_Sheet = self.sheet_data.get_options_sheet()

        # Specify the columns for find and replace dynamically
        find_replace_columns = {'Keyword': 'Replacement'}
        for find_col, replace_col in find_replace_columns.items():
            try:
                self.replace_in_sheet(Chemical_Formula_Sheet, find_col, replace_col)
                self.replace_in_sheet(Symbols_Sheet, find_col, replace_col)
                self.replace_in_sheet(Superscrip_Sheet, find_col, replace_col)
                self.replace_in_sheet(Options_Sheet, find_col, replace_col)
            except Exception as e:
                print(f"Error in replace_in_question_sheet: {e}")

    def save_result_to_excel(self, output_path):
        try:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                Question_Sheet = self.sheet_data.get_question_sheet()
                Question_Sheet.to_excel(writer, sheet_name='sheet1', index=False)
        except Exception as e:
            print(f"Error saving to Excel: {e}")

if __name__ == "__main__":
    google_sheets_url = 'https://docs.google.com/spreadsheets/d/1V_gLg2RipDqcg5yJjUsT7gDqkkOYb7G1PuLLMBGw7bg/edit?usp=sharing'
    sheet_data = GoogleSheetData(google_sheets_url)

    try:
        sheet_data.connect_to_google_sheets()

        find_replace_instance = FindReplace(sheet_data)
        find_replace_instance.replace_in_question_sheet()

        output_path = "C:\\Users\\waquar\\Downloads\\FindAndReplace28DecCB12.xlsx"
        find_replace_instance.save_result_to_excel(output_path)
    except Exception as e:
        print(f"An error occurred: {e}")
