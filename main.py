import requests
import json
import pandas as pd
import os

# Load the abbreviation file
abrev_data = pd.read_excel("fdicabrev.xlsx")

# Create a dictionary from abbreviation file
abrev_dict = abrev_data.set_index('Variable')['Title'].to_dict()

base_url = "https://banks.data.fdic.gov/api"
endpoint = "/financials"

# Define the parameters based on your requirements
params = {
    "filters": "CERT:17975",
    "fields": "CERT,RSSDHCR,NAMEFULL,CITY,STALP,REPDTE,BKCLASS,NAMEHCR,OFFDOM,SUBCHAPS,ESTYMD,EFFDATE,PARCERT,TRUST,REGAGNT,CB,NTINCHPP,INTINCY,INTEXPY,NIMY,NONIIAY,NONIXAY,ELNATRY,NOIJY,ROA,ROAPTX,ROE,ROEINJR,NTLNLSR,NTRER,NTRECOSR,NTRENRSR,NTREMULR,NTRERESR,NTRELOCR,NTREOTHR,IDNTCIR,IDNTCONR,IDNTCRDR,IDNTCOOR,NTAUTOPR,NTCONOTR,NTALLOTHR,NTCOMRER,ELNANTR,IDERNCVR,EEFFR,ASTEMPM,EQCDIVNTINC,ERNASTR,LNATRESR,LNRESNCR,NPERFV,NCLNLS,NCLNLSR,NCRER,NCRECONR,NCRENRER,NCREMULR,NCRERESR,NCRELOCR,NCREREOR,IDNCCIR,IDNCCONR,IDNCCRDR,IDNCCOOR,IDNCATOR,IDNCCOTR,IDNCOTHR,NCCOMRER,IDNCGTPR,LNLSNTV,LNLSDEPR,IDLNCORR,DEPDASTR,EQV,RBC1AAJ,CBLRIND,IDT1CER,IDT1RWAJR,RBCRWAJ",
    "limit": 24,
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

# If the filters field in the params contains only one field and it's a "CERT"...
if len(params['filters'].split(',')) == 1 and 'CERT' in params['filters']:
    # Convert the last 8 digits of the ID field to a date
    df['ID'] = pd.to_datetime(df['ID'].apply(lambda x: str(x)[-8:]), format='%Y%m%d')
    
    # Sort the DataFrame by REPDTE
    df.sort_values(by='REPDTE', inplace=True)
    
    # Make REPDTE the first column
    df = df.set_index('REPDTE').reset_index()
    
    # Get the name for the sheet
    sheet_name = df['NAMEFULL'].iloc[0]
    
    # The name of the output Excel file (first 10 characters of 'NAMEFULL')
    output_file = f'{sheet_name[:10]}.xlsx'
    
    # Write the data to the Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Create a list to keep track of the new column names
        new_cols = []

        # Loop over the columns of the output file
        for col in df.columns:
            # If the column is in the abbreviation dictionary, replace it
            if col in abrev_dict:
                new_cols.append(abrev_dict[col])
            else:
                # If not, keep the original column name
                new_cols.append(col)

        # Rename the columns in the output file
        df.columns = new_cols

        # Write the DataFrame to the sheet
        df.to_excel(writer, sheet_name=sheet_name, index=False)
