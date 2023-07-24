import pandas as pd
import os
import zipfile

def handle_irregular_data(filename):
    data = []
    with open(filename, 'r') as f:
        lines = f.readlines()
        headers = [header.replace('"', '') for header in lines[0].strip().split('\t')]
        for line in lines[1:]:
            values = [value.replace('"', '') for value in line.strip().split('\t')]
            row = {headers[i]: values[i] for i in range(len(values))}
            data.append(row)
    return data, headers

# Unzip the file
with zipfile.ZipFile('FFIEC CDR Call Bulk All Schedules 033120232.zip', 'r') as zip_ref:
    zip_ref.extractall('extracted')

# Get the list of unzipped files
unzipped_dir = 'extracted'
unzipped_files = os.listdir(unzipped_dir)

# Identify the "Bulk" file and filter out those that have "AL" in their 9th column
bulk_file_name = [file for file in unzipped_files if "Bulk" in file][0]
bulk_file_path = os.path.join(unzipped_dir, bulk_file_name)

# Handle the irregular data in the "Bulk" file
bulk_data, bulk_headers = handle_irregular_data(bulk_file_path)

# Store the IDRSSD values from the "Bulk" file in a set for faster lookup
idrssd_set = set(row['IDRSSD'] for row in bulk_data if row[bulk_headers[8]] == "AL")

# Dictionary to store the headers and corresponding data from each file
data_dict = {}

# Handle the irregular data in the "Bulk" file and store the headers and corresponding data
for row in bulk_data:
    idrssd = row[bulk_headers[0]]
    if idrssd in idrssd_set:
        if idrssd not in data_dict:
            data_dict[idrssd] = {}
        for i in range(1, len(bulk_headers)):
            if bulk_headers[i] in row:
                try:
                    data_dict[idrssd][bulk_headers[i]] = float(row[bulk_headers[i]])
                except ValueError:
                    data_dict[idrssd][bulk_headers[i]] = row[bulk_headers[i]]

# For each file that isn't the README file, the 'Bulk' file, or the "CI" file, store the headers and corresponding data
other_files = [file for file in unzipped_files if "README" not in file and "Bulk" not in file and "CI" not in file]

for file in other_files:
    file_path = os.path.join(unzipped_dir, file)
    
    with open(file_path, 'r') as f:
        headers = [header.replace('"', '') for header in f.readline().strip().split('\t')]

    data = handle_irregular_data(file_path)[0]

    for row in data:
        idrssd = row[headers[0]]
        if idrssd in idrssd_set:
            if idrssd not in data_dict:
                data_dict[idrssd] = {}
            for i in range(1, len(headers)):
                if headers[i] in row:
                    try:
                        data_dict[idrssd][headers[i]] = float(row[headers[i]])
                    except ValueError:
                        data_dict[idrssd][headers[i]] = row[headers[i]]

# ...

# Convert the data dictionary into a pandas DataFrame
df = pd.DataFrame(data_dict).T

# Create a new Excel writer object
with pd.ExcelWriter('master.xlsx') as excel_writer:
    # Write the DataFrame to the Excel file
    df.to_excel(excel_writer, index_label='IDRSSD')

    # Save and close the Excel file
    excel_writer.close()
