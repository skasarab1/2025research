import pandas as pd


PSSE_Data = 'In_PSSEData.xlsx'
UNKNOWN = 'missing_location.xlsx'

PSSE_Data = pd.read_excel(PSSE_Data, sheet_name='PSSE_Buses')
Unknown = pd.read_excel(UNKNOWN)

Unknown['Bus Name'] = None

def get_bus_name(bus_number, psse_df):
    match = psse_df[psse_df['Bus  Number'] == bus_number]
    if not match.empty:
        return match.iloc[0]['Bus Name']
    return None

for idx, row in Unknown.iterrows():
    busnum = row['Bus Number']
    row['Bus Name'] = get_bus_name(busnum, PSSE_Data)

Unknown.to_excel(UNKNOWN)





