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

    # Check if the second sheet exists, if not, create it
    second_sheet_title = extract_sheet_name(txt_files[0])
    if second_sheet_title:
        if second_sheet_title not in wb.sheetnames:
            wb.create_sheet(title=second_sheet_title)

    # Get the correct sheet
    sheet = wb[second_sheet_title]

    # Process each .txt file and write its contents to the second sheet in the Excel file
    for txt_file in txt_files:
        txt_path = os.path.join(script_folder, txt_file)
        if txt_file.endswith('.xls'):
            xls_data = xlrd.open_workbook(txt_path)
            sheet = xls_data.sheet_by_index(0)
            headers = sheet.row_values(0)
            data = [sheet.row_values(i) for i in range(1, sheet.nrows)]
            df = pd.DataFrame(data, columns=headers)
        else:  # Assuming it's a .txt file with tab-delimited data
            # Specify data types and set low_memory=False to prevent DtypeWarning
            df = pd.read_csv(txt_path, delimiter='\t', dtype=str, low_memory=False)

        # Calculate the starting row and column based on existing data on the second sheet
        start_row = sheet.max_row + 1 if sheet.max_row > 0 else 1
        start_column = sheet.max_column + 1 if sheet.max_column > 0 else 1

        # Append the contents to the second sheet
        if txt_file == txt_files[0]:  # For the first file, include the first column
            for index, row in df.iterrows():
                for col_index, value in enumerate(row):
                    sheet.cell(row=start_row + index, column=start_column + col_index, value=value)
        else:  # For subsequent files, exclude the first column and append to the right
            for index, row in df.iloc[:, 1:].iterrows():
                for col_index, value in enumerate(row):
                    sheet.cell(row=start_row + index, column=start_column + col_index, value=value)

    # Save the changes to master.xlsx
    wb.save(excel_file)

# Example usage
if __name__ == "__main__":
    excel_file_path = 'master.xlsx'  # Change the extension to .xlsx
    add_txt_to_excel(excel_file_path)
