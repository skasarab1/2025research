import pandas as pd

# Load both files
missing = pd.read_excel("Missing_Buses.xlsx")
planning = pd.read_excel("PlanningQueues.xlsx")

# Normalize columns
missing['Bus  Name'] = missing['Bus  Name'].astype(str).str.strip()
planning['Project ID'] = planning['Project ID'].astype(str).str.strip()
planning['Status'] = planning['Status'].astype(str).str.upper().str.strip()

# Add new columns
missing['Reasoning'] = ""
missing['Real Name'] = ""

# Define areas to strike
target_areas = ['IESO', 'TE', 'NB', 'NS', 'CORNWALL', 'NF', 'SPC', 'MHEB']
missing[' Area Name'] = missing[' Area Name'].astype(str).str.upper().str.strip()

# Strikethrough function
def strikethrough(text):
    return ''.join([c + '\u0336' for c in str(text)])

# Match and apply logic
for idx, row in missing.iterrows():
    prefix = row['Bus  Name'][:6]
    area = row[' Area Name']
    matches = planning[planning['Project ID'].str.startswith(prefix)]

    if area in target_areas:
        missing.at[idx, 'Bus  Name'] = strikethrough(row['Bus  Name'])
        missing.at[idx, 'Reasoning'] = 'Out-of-area'
        missing.at[idx, 'Lat'] = 'X'
        missing.at[idx, 'Long'] = 'X'

    elif not matches.empty:
        status = matches.iloc[0]['Status']
        
        if status in ['WITHDRAWN', 'RETRACTED']:
            missing.at[idx, 'Bus  Name'] = strikethrough(row['Bus  Name'])
            missing.at[idx, 'Reasoning'] = 'Withdrawn'
            missing.at[idx, 'Lat'] = 'X'
            missing.at[idx, 'Long'] = 'X'

        elif status == 'IN SERVICE':
            missing.at[idx, 'Reasoning'] = 'In Service'
            missing.at[idx, 'Real Name'] = matches.iloc[0]['Commercial Name']

# Save result
missing.to_excel("Missing_Buses.xlsx", index=False)

print("✅ Missing_Buses.xlsx updated correctly.")
