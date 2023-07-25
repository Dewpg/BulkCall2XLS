import requests
import json
import pandas as pd
import os

base_url = "https://banks.data.fdic.gov/api"
endpoint = "/financials"

# Define the parameters based on your requirements
params = {
    "filters": "STALP:NY",
    "fields":"NAME,ASSET,NETINC,ROA,IDT1RWAJR,RBCRWAJ,LNLSNET,DEPDOM,ELNLOS,LNATRES,BRO,OTBFH1L,OTBFH1T3,OTBGH3T5,OTBGHOV5,OTHBOT1L,OTBOT1T3,OTBOT3T5,OTBOTOV5,FREPP,DEPLSNB,CHBALNI,CHBALI,FREPO,SCAA,SCHF,SCPLEDGE",
    "limit": 100,
    "sort_order": "DESC",
    "sort_by": "REPDTE",
    "offset": 0,
    "format": "json"
}

# Send the request
response = requests.get(base_url + endpoint, params=params)

# Parse the response
data = response.json()

# Extract the data records
records = [item['data'] for item in data['data']]

# Convert the data into a DataFrame
df = pd.DataFrame(records)

# Add a new column for the last 8 digits of the ID
df['ID_Last8'] = df['ID'].apply(lambda x: str(x)[-8:])

# The name of the output Excel file
output_file = 'fdicdata.xlsx'

# If the file exists, read the existing data
if os.path.exists(output_file):
    existing_data = pd.read_excel(output_file, sheet_name=None)
else:
    existing_data = {}

# Append the new data to each sheet
for id_value in df['ID_Last8'].unique():
    df_subset = df[df['ID_Last8'] == id_value]
    
    # If the sheet exists, append the data. Otherwise, create a new sheet.
    if id_value in existing_data:
        existing_data[id_value] = pd.concat([existing_data[id_value], df_subset])
    else:
        existing_data[id_value] = df_subset

# Write the data back to the Excel file
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, data in existing_data.items():
        data.to_excel(writer, sheet_name=sheet_name, index=False)
