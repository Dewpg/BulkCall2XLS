import os
import openpyxl
import pandas as pd
import re
import xlrd

def extract_sheet_name(filename):
    # Use regular expressions to extract the last 8 characters before .txt in the first file
    pattern = r"(.{8})\.txt$"
    match = re.search(pattern, filename)
    if match:
        return match.group(1)
    else:
        return None

def add_txt_to_excel(excel_file):
    script_folder = os.path.dirname(os.path.abspath(__file__))
    txt_files = [f for f in os.listdir(script_folder) if f.endswith('.txt')]

    if not txt_files:
        print("No txt files found in the folder.")
        return

    # Create or load the master.xlsx file
    if os.path.exists(excel_file):
        wb = openpyxl.load_workbook(excel_file)
    else:
        wb = openpyxl.Workbook()

    # Get the correct sheet or create it if not exists
    first_txt_file = None
    for txt_file in txt_files:
        if "BULK" in txt_file.upper():
            first_txt_file = txt_file
            break

    if not first_txt_file:
        print("No file with 'BULK' in its filename found.")
        return

    first_sheet_name = extract_sheet_name(first_txt_file)
    if first_sheet_name not in wb.sheetnames:
        sheet = wb.create_sheet(title=first_sheet_name)
    else:
        sheet = wb[first_sheet_name]

    # Process the first .txt file and write its contents to the sheet in the Excel file
    first_txt_path = os.path.join(script_folder, first_txt_file)
    if first_txt_file.endswith('.xls'):
        xls_data = xlrd.open_workbook(first_txt_path)
        sheet_data = xls_data.sheet_by_index(0)
        headers = sheet_data.row_values(0)
        data = [sheet_data.row_values(i) for i in range(sheet_data.nrows)]
    else:
        df = pd.read_csv(first_txt_path, delimiter='\t', dtype=str, low_memory=False)
        data = df.values.tolist()
        headers = df.columns.tolist()

    # Calculate the starting column based on existing data on the sheet
    start_column = sheet.max_column + 1 if sheet.max_column is not None else 1

    # Set the starting row to 1 for the first .txt file
    start_row = 1

    # Append the contents to the sheet for the first file (include the first column)
    for index, row in enumerate(data):
        for col_index, value in enumerate(row):
            sheet.cell(row=start_row + index, column=start_column + col_index, value=value)

    # Process subsequent .txt files and append their contents to the right of the first file
    for txt_file in txt_files:
        if txt_file != first_txt_file:
            txt_path = os.path.join(script_folder, txt_file)
            if txt_file.endswith('.xls'):
                xls_data = xlrd.open_workbook(txt_path)
                sheet_data = xls_data.sheet_by_index(0)
                data = [sheet_data.row_values(i) for i in range(sheet_data.nrows)]
            else:
                df = pd.read_csv(txt_path, delimiter='\t', dtype=str, low_memory=False)

                # Remove the first column for subsequent files
                df = df.iloc[:, 1:]

                # Convert the DataFrame to a list of lists to match the .xls data structure
                data = df.values.tolist()

            # Append the contents to the sheet for subsequent files (exclude the first column)
            for index, row in enumerate(data):
                for col_index, value in enumerate(row):
                    sheet.cell(row=start_row + index, column=start_column + col_index, value=value)

            # Update the starting column for subsequent files
            start_column = sheet.max_column + 1

    # Save the changes to master.xlsx
    wb.save(excel_file)

# Example usage
if __name__ == "__main__":
    excel_file_path = 'master.xlsx'  # Change the extension to .xlsx
    add_txt_to_excel(excel_file_path)
