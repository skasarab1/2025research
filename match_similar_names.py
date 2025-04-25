import pandas as pd

# Load file
df = pd.read_excel("Missing_Buses.xlsx")

# Make sure column names are clean
df['Bus  Name'] = df['Bus  Name'].astype(str).str.strip()

# Extract queue ID (e.g., Q853) from bus name
df['Queue'] = df['Bus  Name'].str.extract(r'(Q\d+)')

# Fill missing Lat/Lon based on other entries with same queue
for queue_id, group in df.groupby('Queue'):
    lat = group['Lat'].dropna().values
    lon = group['Long'].dropna().values
    if lat.size > 0 and lon.size > 0:
        df.loc[df['Queue'] == queue_id, 'Lat'] = lat[0]
        df.loc[df['Queue'] == queue_id, 'Long'] = lon[0]

# Drop helper column
df.drop(columns='Queue', inplace=True)

# Save back to the original file
df.to_excel("Missing_Buses.xlsx", index=False)

