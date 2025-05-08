import pandas as pd

# Load files
missing = pd.read_excel("missing_location.xlsx")
working = pd.read_excel("Missing_Buses.xlsx")


working = working[['Bus Number', 'Lat', 'Long']].dropna()
coord_lookup = working.set_index('Bus Number')[['Lat', 'Long']].to_dict('index')

for idx, row in missing.iterrows():
    bus = row['Bus Number']
    if bus in coord_lookup:
        missing.at[idx, 'Lat'] = coord_lookup[bus]['Lat']
        missing.at[idx, 'Long'] = coord_lookup[bus]['Long']

missing.to_excel("missing_location.xlsx", index=False)
