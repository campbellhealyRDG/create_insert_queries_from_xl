import pandas as pd

# Load the spreadsheet
# file_path = 'data/Staging_Inventory.xlsx'  # Change this to your file path
file_path = 'data/Live_Inventory.xlsx'  # Change this to your file path
df = pd.read_excel(file_path)

# Define the table name and columns (excluding IdDevice which is auto-incremented)
table_name = 'Devices'
columns = [
    'IdEnvironment', 'IdTOC', 'IdLogicalGroup', 'IdISAMVersion', 'IRN',
    'ISAMID', 'DTS_Commissioned', 'IdFirmware', 'IdSupplier', 'POSTIdentifier',
    'IdISAMUse', 'IdStatus', 'HotlistPOSTSet', 'ActionlistPOSTSet',
    'DTS_Decommissioned', 'ServiceNowReference', 'RDGNotes', 'OtherNotes'
]

# Generate the INSERT statements
insert_statements = []
for index, row in df.iterrows():
    values = [
        f"(SELECT IdEnvironment FROM Environments WHERE EnvironmentName = '{row['IdEnvironment']}')",
        f"(SELECT IdTOC FROM TOCs WHERE TOCName = '{row['IdTOC']}')",
        f"(SELECT IdLogicalGroup FROM LogicalGroups WHERE LogicalGroupName = '{row['IdLogicalGroup']}')",
        f"(SELECT IdISAMVersion FROM ISAMVersions WHERE ISAMVersionName = '{row['IdISAMVersion']}')",
        f"'{row['IRN']}'",
        f"'{row['ISAMID']}'",
        'NULL' if pd.isnull(row['DTS_Commissioned']) else f"'{row['DTS_Commissioned']}'",
        'NULL' if pd.isnull(row['IdFirmware']) else f"(SELECT IdFirmware FROM Firmware WHERE FirmwareName = '{row['IdFirmware']}')",
        'NULL' if pd.isnull(row['IdSupplier']) else f"(SELECT IdSupplier FROM Suppliers WHERE SupplierName = '{row['IdSupplier']}')",
        'NULL' if pd.isnull(row['POSTIdentifier']) else f"'{row['POSTIdentifier']}'",
        f"(SELECT IdISAMUse FROM ISAMUses WHERE ISAMUseName = '{row['IdISAMUse']}')",
        f"(SELECT IdStatus FROM Status WHERE StatusName = '{row['IdStatus']}')",
        'NULL' if pd.isnull(row['HotlistPOSTSet']) else f"'{row['HotlistPOSTSet']}'",
        'NULL' if pd.isnull(row['ActionlistPOSTSet']) else f"'{row['ActionlistPOSTSet']}'",
        'NULL' if pd.isnull(row['DTS_Decommissioned']) else f"'{row['DTS_Decommissioned']}'",
        'NULL' if pd.isnull(row['ServiceNowReference']) else f"'{row['ServiceNowReference']}'",
        'NULL' if pd.isnull(row['RDGNotes']) else f"'{row['RDGNotes']}'",
        'NULL' if pd.isnull(row['OtherNotes']) else f"'{row['OtherNotes']}'"
    ]
    insert_statement = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES\n({', '.join(values)});"
    insert_statements.append(insert_statement)

# Save the INSERT statements to a text file
# output_file_path = 'staging_insert_statements.txt'
output_file_path = 'live_insert_statements.txt'
with open(output_file_path, 'w') as file:
    for statement in insert_statements:
        file.write(statement + '\n')

print(f"INSERT statements saved to {output_file_path}")
