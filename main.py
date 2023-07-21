import csv
import openpyxl
import os
from openpyxl.utils import get_column_letter

def clean_sheet_name(sheet_name):
    # Remove characters that cannot be used in a worksheet name
    return "".join(c for c in sheet_name if c.isalnum() or c in [' ', '_', '-'])

def delete_empty_columns(sheet):
    # Find the columns that are empty (all cells in the column are None or empty strings)
    empty_columns = set(col_idx for col_idx in range(1, sheet.max_column + 1) if all(not sheet[get_column_letter(col_idx) + str(row)].value for row in range(1, sheet.max_row + 1)))

    # Delete the empty columns in reverse order to avoid index issues
    for col_idx in sorted(empty_columns, reverse=True):
        sheet.delete_cols(col_idx)

def find_empty_column(sheet):
    # Find the first empty column in the sheet
    for col in range(1, sheet.max_column + 2):
        col_letter = get_column_letter(col)
        if all(not sheet[col_letter + str(row)].value for row in range(1, sheet.max_row + 1)):
            return col
    return sheet.max_column + 2

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

    print("Creating dated sheetname as second sheet in file")
    # Create a worksheet with the name of the last 8 characters before the .txt for the BULK file
    sheet_name = bulk_filename[-12:-4]
    sheet_name = clean_sheet_name(sheet_name)  # Clean the sheet name
    if sheet_name not in wb.sheetnames:
        sheet = wb.create_sheet(title=sheet_name)
    else:
        sheet = wb[sheet_name]
    print("Sheet named " + sheet_name)
    
    # Read the data from the BULK file
    with open(bulk_filename, 'r') as f:
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
    
    # Get the starting column based on existing data on the sheet or the first empty column
    start_column = find_empty_column(sheet)

    # Iterate through the rest of the TXT files and add the data to the right
    for filename in filenames:
        with open(filename, 'r') as f:
            reader = csv.reader(f, delimiter='\t')
            data = list(reader)
        print("Adding " + filename)
        print("Checking to validate value type as number")
        # Write the data to the existing sheet starting from the first empty column
        for row, line in enumerate(data):
            for col, cell in enumerate(line):
                col_letter = get_column_letter(start_column + col)  # Convert the column index to letter
                
                # Check if the cell value is a number and set the data type accordingly
                if cell.isnumeric():
                    sheet[col_letter + str(row + 1)].data_type = 'n'
                    sheet[col_letter + str(row + 1)].value = int(cell)
                else:
                    sheet[col_letter + str(row + 1)] = cell
        print("Finished datatyping a file") 
        # Increment the starting column for the next file
        start_column += len(data[0])
    
    # Save the changes to master.xlsx
    wb.save('master.xlsx')
    
    print("Deleting Blank Columns")
    # Open the xlsx document again
    wb = openpyxl.load_workbook('master.xlsx')

    # Get the sheet
    sheet = wb[sheet_name]

    # Delete empty columns from the sheet
    delete_empty_columns(sheet)

    # Save the workbook again
    wb.save('master.xlsx')
    print("Deletions processed")
    
if __name__ == '__main__':
    main()
