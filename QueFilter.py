import pandas as pd
import re

# Load both files
missing = pd.read_excel("Missing_Buses.xlsx")
planning = pd.read_excel("PlanningQueues.xlsx")

# Normalize text
missing['Bus  Name'] = missing['Bus  Name'].astype(str).str.strip().str.upper()
planning['Project ID'] = planning['Project ID'].astype(str).str.strip().str.upper()
planning['Status'] = planning['Status'].astype(str).str.strip().str.upper()
planning['Commercial Name'] = planning['Commercial Name'].astype(str).str.strip()
planning['Fuel'] = planning['Fuel'].astype(str).str.strip()

# Add new columns if not already present
if 'Reasoning' not in missing.columns:
    missing['Reasoning'] = ""
if 'Real Name' not in missing.columns:
    missing['Real Name'] = ""

missing['Type'] = ""
missing['Capacity'] = None


# Match and apply logic
for idx, row in missing.iterrows():
    bus_name = row['Bus  Name']

    # Extract just the project ID portion (e.g. "AD2-210")
    match_id = re.match(r'([A-Z]+\d*-\d+)', bus_name)

    if not match_id:
        continue  # Skip if no match

    project_id_prefix = match_id.group(1)
    matches = planning[planning['Project ID'].str.startswith(project_id_prefix)]

    if not matches.empty:
        status = matches.iloc[0]['Status']
        if status in ['WITHDRAWN', 'RETRACTED','SUSPENDED','CANCELED']:
            continue
        elif status in ['IN SERVICE','UNDER CONSTRUCTION' , 'Partially in Service - Under Construction' , 'Engineering and Procurement' , 'Under Construction']:
            missing.at[idx, 'Reasoning'] = 'In Service'
            missing.at[idx, 'Real Name'] = matches.iloc[0]['Commercial Name']
            missing.at[idx, 'Type'] = matches.iloc[0]['Fuel']
            missing.at[idx, 'Capacity'] = matches.iloc[0]['MW In Service']

# Save the updated file
missing.to_excel("Missing_Buses.xlsx", index=False)
