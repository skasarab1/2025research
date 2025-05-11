import pandas as pd

current = pd.read_excel("Fikir.xlsx")
working = pd.read_excel("Missing_Buses.xlsx")

current['Bus  Number'] = current['Bus  Number'].astype(str).str.strip()
working['Bus  Number'] = working['Bus  Number'].astype(str).str.strip()

for _, row in current.iterrows():
    busnum1 = row['Bus  Number']
    if pd.notna(row['latitude']):
        match_index = working[working['Bus  Number'] == busnum1].index
        if not match_index.empty:
            working.loc[match_index[0], 'Lat'] = row['latitude']
            working.loc[match_index[0], 'Long'] = row['longitude']

working.to_excel('Missing_Buses.xlsx')


