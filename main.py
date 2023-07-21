import csv
import openpyxl
import os
from openpyxl.utils import get_column_letter

def clean_sheet_name(sheet_name):
    # Remove characters that cannot be used in a worksheet name
    return "".join(c for c in sheet_name if c.isalnum() or c in [' ', '_', '-'])

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
            
            # Check if the cell value is a number and set the data type accordingly
            if cell.isnumeric():
                sheet[col_letter + str(row + 1)] = int(cell)
            else:
                sheet[col_letter + str(row + 1)] = cell

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

                # Check if the cell value is a number and set the data type accordingly
                if cell.isnumeric():
                    sheet[col_letter + str(row + 2)].data_type = 'n'
                    sheet[col_letter + str(row + 2)].value = int(cell)
                else:
                    sheet[col_letter + str(row + 2)] = cell

        # Increment the starting column for the next file
        start_column += len(data[0]) - 1

    # Delete empty columns
    columns_to_delete = [col for col in range(1, sheet.max_column + 1) if not any(sheet.cell(row=row, column=col).value for row in range(1, sheet.max_row + 1))]
    for col in reversed(columns_to_delete):
        sheet.delete_cols(col)

    # Save the changes to master.xlsx
    wb.save('master.xlsx')

if __name__ == '__main__':
    main()
