import pandas as pd
import re

# Load data
missing = pd.read_excel("Missing_Buses.xlsx")
withdrawn = pd.read_excel('NYISO-Interconnection-Queue.xlsx', sheet_name='Withdrawn')
inservice = pd.read_excel('NYISO-Interconnection-Queue.xlsx', sheet_name='In Service')
que = pd.read_excel('NYISO-Interconnection-Queue.xlsx', sheet_name='Interconnection Queue')

# Normalize columns
missing['Bus  Name'] = missing['Bus  Name'].astype(str).str.strip().str.upper()

inservice['Queue Pos.'] = inservice['Queue Pos.'].astype(str).str.strip()
inservice['Project Name'] = inservice['Project Name'].astype(str).str.strip()
inservice['Fuel'] = inservice['Fuel'].astype(str).str.strip()
inservice['Interconnection Point'] = inservice['Interconnection Point'].astype(str).str.strip()

withdrawn['Queue Pos.'] = withdrawn['Queue Pos.'].astype(str).str.strip()
que['Queue Pos.'] = que['Queue Pos.'].astype(str).str.strip()


# Ensure output columns exist
for col in ['Reasoning', 'Real Name', 'Type', 'Capacity', 'POI', 'Lat', 'Long']:
    if col not in missing.columns:
        missing[col] = ""

# Process matches
for idx, row in missing.iterrows():
    bus_name = row['Bus  Name']
    match_id = re.match(r'^.(\d{3})', bus_name)  # extract 3 digits after first character

    if not match_id:
        continue

    project_id_prefix = match_id.group(1)

    # Check in-service
    matches = inservice[inservice['Queue Pos.'] == project_id_prefix]
    if not matches.empty:
        missing.at[idx, 'Reasoning'] = 'In Service'
        missing.at[idx, 'Real Name'] = matches.iloc[0]['Project Name']
        missing.at[idx, 'Type'] = matches.iloc[0]['Fuel']
        missing.at[idx, 'Capacity'] = matches.iloc[0]['SP (MW)']
        missing.at[idx, 'POI'] = matches.iloc[0]['Interconnection Point']
        continue

    # Check in queue
    matches = que[que['Queue Pos.'] == project_id_prefix]
    if not matches.empty:
        missing.at[idx, 'Reasoning'] = 'In Queue'
        missing.at[idx, 'Real Name'] = matches.iloc[0]['Project Name']
        missing.at[idx, 'Type'] = matches.iloc[0]['Type/ Fuel']
        missing.at[idx, 'Capacity'] = matches.iloc[0]['SP (MW)']
        continue

    # Check withdrawn
    matches = withdrawn[withdrawn['Queue Pos.'] == project_id_prefix]
    if not matches.empty:
        missing.at[idx, 'Reasoning'] = 'Withdrawn'
        missing.at[idx, 'Lat'] = 'X'
        missing.at[idx, 'Long'] = 'X'

# Save results
missing.to_excel("test.xlsx", index=False)
