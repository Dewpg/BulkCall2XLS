import csv
import openpyxl
import os
from openpyxl.utils import get_column_letter

def clean_sheet_name(sheet_name):
    # Remove characters that cannot be used in a worksheet name
    return "".join(c for c in sheet_name if c.isalnum() or c in [' ', '_', '-'])

def convert_value(value):
    # Convert the value to a string, handling special cases if needed
    if isinstance(value, str):
        return value
    elif value is None:
        return ""
    else:
        return str(value)

def main():
    # Get all filenames that end with ".txt"
    filenames = [f for f in os.listdir('.') if f.endswith('.txt')]

    # Find the filename that contains the word "BULK" and remove it from the list
    bulk_filename = next((f for f in filenames if 'Bulk' in f), None)
    if bulk_filename:
        filenames.remove(bulk_filename)
    else:
        print("No file with 'BULK' in its filename found.")
        return

    # Create or load the master.xlsx file
    if os.path.exists('master.xlsx'):
        wb = openpyxl.load_workbook('master.xlsx')
    else:
        wb = openpyxl.Workbook()

    # Create a worksheet with the name of the last 8 characters before the .txt for the BULK file
    sheet_name = bulk_filename[-12:-4]
    sheet_name = clean_sheet_name(sheet_name)  # Clean the sheet name
    if sheet_name not in wb.sheetnames:
        sheet = wb.create_sheet(title=sheet_name)
    else:
        sheet = wb[sheet_name]

    # Read the data from the BULK file
    with open(bulk_filename, 'r', encoding='latin-1') as f:
        reader = csv.reader(f, delimiter='\t')
        data = list(reader)

    # Write the data to the worksheet
    for row, line in enumerate(data):
        for col, cell in enumerate(line):
            col_letter = get_column_letter(col + 1)  # Convert the column index to letter
            sheet[col_letter + str(row + 1)] = convert_value(cell)
    
    # Insert a blank row after the headers
    sheet.insert_rows(2)
    
    # Get the starting column based on existing data on the sheet
    start_column = sheet.max_column + 1 if sheet.max_column is not None else 1

    # Iterate through the rest of the TXT files and add the data to the right
    for filename in filenames:
        with open(filename, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter='\t')
            data = list(reader)

        # Write the data to the existing sheet
        for row, line in enumerate(data):
            for col, cell in enumerate(line[1:]):  # Skip the first column
                col_letter = get_column_letter(start_column + col)  # Convert the column index to letter
                sheet[col_letter + str(row + 1)] = convert_value(cell)

        # Increment the starting column for the next file
        start_column += len(data[0]) - 1

    # Save the changes to master.xlsx
    wb.save('master.xlsx')

if __name__ == '__main__':
    main()
