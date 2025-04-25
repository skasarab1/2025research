import pandas as pd

# Load files
missing = pd.read_excel("Missing_Buses.xlsx")
working = pd.read_excel("WorkingSet.xlsx")

# Filter only valid lat/lon rows
working = working[['Bus  Number', 'latitude', 'longitude']].dropna()

# Create lookup dict
coord_lookup = working.set_index('Bus  Number')[['latitude', 'longitude']].to_dict('index')

# Fill Lat and Lon where available
for idx, row in missing.iterrows():
    bus = row['Bus  Number']
    if bus in coord_lookup:
        missing.at[idx, 'Lat'] = coord_lookup[bus]['latitude']
        missing.at[idx, 'Long'] = coord_lookup[bus]['longitude']

# Save back
missing.to_excel("Missing_Buses.xlsx", index=False)
