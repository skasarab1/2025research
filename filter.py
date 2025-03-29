import pandas as pd

# File paths
PSSE_Data = 'In_PSSEData.xlsx'
Bus_Locations = 'EI_Bus_loca.xlsx'
ISOne = 'ISONE.xlsx'
NYISO = 'NYISO.xlsx'

PSSE = pd.read_excel(PSSE_Data, sheet_name='PSSE_Buses')
Bus_Loc = pd.read_excel(Bus_Locations)
ISONE = pd.read_excel(ISOne)

def match_bus_number(df1, df2):
    for index, row in df1.iterrows():
        search_key = row['Bus  Number']

        # Filter rows where BUS_I exactly matches the numeric search_key
        matching_rows = df2[df2['BUS_I'] == search_key]

        if not matching_rows.empty:
            match = matching_rows.iloc[0]

            if pd.notna(match['latitude']) and pd.notna(match['longitude']):
                df1.at[index, 'latitude'] = match['latitude']
                df1.at[index, 'longitude'] = match['longitude']

# Identify unmatched bus numbers
unmatched_mask = ~PSSE['Bus  Number'].isin(Bus_Loc['BusNumber'])

unmatched_df = PSSE.loc[unmatched_mask, ['Bus  Number', 'Bus  Name',' Area Name' ,'Base kV']].copy()
unmatched_df['latitude'] = None
unmatched_df['longitude'] = None

match_bus_number(unmatched_df,ISONE)

# Export to Excel
unmatched_df.to_excel('result.xlsx', index=True)

