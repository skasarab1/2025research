import pandas as pd

# File paths
PSSE_Data = 'In_PSSEData.xlsx'
Bus_Locations = 'EI_Bus_loca.xlsx'
ISOne = 'ISONE.xlsx'
NYISO = 'NYISO.xlsx'
UNKNOWN = 'TOBEFILLED.xlsx'

# Read files
PSSE_Lines = pd.read_excel(PSSE_Data, sheet_name='PSSE_Lines')
PSSE_Data = pd.read_excel(PSSE_Data, sheet_name='PSSE_Buses')
Bus_Loc = pd.read_excel(Bus_Locations)
ISONE = pd.read_excel(ISOne)
Unknown = pd.read_excel(UNKNOWN)

# Initialize new columns
Unknown['Closest Name'] = None
Unknown['Closest Lat'] = None
Unknown['Closest Lon'] = None
Unknown['N Levels'] = None
Unknown['Same Bus?'] = None

Unknown = Unknown[Unknown['Base kV']>= 69]


def get_bus_name(bus_number, psse_df):
    match = psse_df[psse_df['Bus  Number'] == bus_number]
    if not match.empty:
        return match.iloc[0]['Bus  Name']
    return None

# Loop through each unknown bus
for idx, row in Unknown.iterrows():
    BusNum = row['Bus  Number']

    # Get directly connected buses (level 1)
    MatchingRows = PSSE_Lines[
        (PSSE_Lines['From Bus  Number'] == BusNum) |
        (PSSE_Lines['To Bus  Number'] == BusNum)
    ]

    ConnectedBuses = []
    
    for _, line in MatchingRows.iterrows():
        if line['From Bus  Number'] == BusNum:
            ConnectedBuses.append({
            'BusNumber': line['To Bus  Number'],
            'Reactance': line['Line X (pu)']
            })
        else:
            ConnectedBuses.append({
            'BusNumber': line['From Bus  Number'],
            'Reactance': line['Line X (pu)']
            })


    # LEVEL 1 SEARCH
    found = False
    for conn in ConnectedBuses:

        target_bus_num = conn['BusNumber']
        reactance = conn['Reactance']

        match = Bus_Loc[Bus_Loc['BusNumber'] == target_bus_num]
        
        if not match.empty:

            if reactance <= .001 :
                Unknown.at[idx, 'latitude'] = match.iloc[0]['Lat']
                Unknown.at[idx,'longitude'] = match.iloc[0]['Long']
                Unknown.at[idx, 'N Levels'] = 0
                Unknown.at[idx, 'Closest Lat'] = 'X'
                Unknown.at[idx, 'Closest Lon'] = 'X'
                Unknown.at[idx, 'Closest Name'] = 'X'
                Unknown.at[idx, 'Same Bus?'] = 'Yes'
                found = True
                break

            else:
                Unknown.at[idx, 'Closest Name'] = get_bus_name(target_bus_num, PSSE_Data)
                Unknown.at[idx, 'Closest Lat'] = match.iloc[0]['Lat']
                Unknown.at[idx, 'Closest Lon'] = match.iloc[0]['Long']
                Unknown.at[idx, 'N Levels'] = 1
                found = True
                break

    # LEVEL 2 SEARCH
    if not found:
        for conn in ConnectedBuses:
            level1_bus = conn['BusNumber']

            MatchingRows2 = PSSE_Lines[
                (PSSE_Lines['From Bus  Number'] == level1_bus) |
                (PSSE_Lines['To Bus  Number'] == level1_bus)
            ]

            ConnectedBuses2 = []
            for _, line2 in MatchingRows2.iterrows():
                if line2['From Bus  Number'] == level1_bus:
                    ConnectedBuses2.append({
                        'BusNumber': line2['To Bus  Number'],
                    })
                else:
                    ConnectedBuses2.append({
                        'BusNumber': line2['From Bus  Number'],
                    })

            for conn2 in ConnectedBuses2:
                target_bus_num2 = conn2['BusNumber']
                match2 = Bus_Loc[Bus_Loc['BusNumber'] == target_bus_num2]
                if not match2.empty:
                    Unknown.at[idx, 'Closest Name'] = get_bus_name(target_bus_num2, PSSE_Data)
                    Unknown.at[idx, 'Closest Lat'] = match2.iloc[0]['Lat']
                    Unknown.at[idx, 'Closest Lon'] = match2.iloc[0]['Long']
                    Unknown.at[idx, 'N Levels'] = 2
                    found = True
                    break

            if found:
                break

# Save results
Unknown.to_excel('help.xlsx')