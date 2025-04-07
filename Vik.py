import pandas as pd

# File paths
PSSE_Data = 'In_PSSEData.xlsx'
Bus_Locations = 'EI_Bus_loca.xlsx'
ISOne = 'ISONE.xlsx'
NYISO = 'NYISO.xlsx'
UNKNOWN = 'TOBEFILLED.xlsx'

#Read Files
PSSE_Lines = pd.read_excel(PSSE_Data, sheet_name='PSSE_Lines')
PSSE_Data = pd.read_excel(PSSE_Data, sheet_name='PSSE_Buses')
Bus_Loc = pd.read_excel(Bus_Locations)
ISONE = pd.read_excel(ISOne)
Unknown = pd.read_excel(UNKNOWN)

Unknown['Closest Name'] = None
Unknown['Closest Lat'] = None
Unknown['Closest Lon'] = None
Unknown['N Levels'] = None

#Bus Name from Bus Number
def get_bus_name(bus_number, psse_df):
    # Look for the bus number in either From or To columns
    match = psse_df[
        (psse_df['Bus  Number'] == bus_number)
    ]

    if not match.empty:
        # Return the matching name based on which side the number was found
        row = match.iloc[0]
        if row['Bus  Number'] == bus_number:
            return row['Bus  Name']

for idx, row in Unknown.iterrows():
    BusNum = row['Bus  Number'] #Bus Number from Unknown Data set
    BusName = row['Bus  Name'] #Bus Name from Unknown Data set

    #sees all rows where the bus number matches is from or to
    MatchingRows = PSSE_Lines[
        (PSSE_Lines['From Bus  Number'] == BusNum) |
        (PSSE_Lines['To Bus  Number'] == BusNum)
    ]

    ConnectedBuses = [] #empty data set for found buses

    for _, line in MatchingRows.iterrows():

        if line['From Bus  Number'] == BusNum:
            ConnectedBuses.append(line['To Bus  Number'])
        else:
            ConnectedBuses.append(line['From Bus  Number'])
        
    NumConnections = len(set(ConnectedBuses)) # Number of Connections

    #check to see if the connected bus is known or unknown filter it out
    

    for target_bus_num in ConnectedBuses:
        match = Bus_Loc[Bus_Loc['BusNumber'] == target_bus_num]

        if not match.empty:
            Unknown.at[idx, 'Closest Name'] = get_bus_name(target_bus_num, PSSE_Data)
            Unknown.at[idx, 'Closest Lat'] = match.iloc[0]['Lat']
            Unknown.at[idx, 'Closest Lon'] = match.iloc[0]['Long']
            Unknown.at[idx, 'N Levels'] = 1
            break
        
        #N Level 2
        BusNum2 = target_bus_num
        BusName2 = row['Bus  Name'] #Bus Name from Unknown Data set

         #sees all rows where the bus number matches is from or to
        MatchingRows2 = PSSE_Lines[
                (PSSE_Lines['From Bus  Number'] == BusNum2) |
                (PSSE_Lines['To Bus  Number'] == BusNum2)
                 ]

        ConnectedBuses2 = [] #empty data set for found buses

        for i, line2 in MatchingRows2.iterrows():

            if line2['From Bus  Number'] == BusNum2:
                ConnectedBuses2.append(line2['To Bus  Number'])
            else:
                ConnectedBuses2.append(line2['From Bus  Number'])

        for target_bus_num2 in ConnectedBuses2:
            match2 = Bus_Loc[Bus_Loc['BusNumber'] == target_bus_num2]

            if not match2.empty:
                Unknown.at[idx, 'Closest Name'] = get_bus_name(target_bus_num2, PSSE_Data)
                Unknown.at[idx, 'Closest Lat'] = match2.iloc[0]['Lat']
                Unknown.at[idx, 'Closest Lon'] = match2.iloc[0]['Long']
                Unknown.at[idx, 'N Levels'] = 2

                break


Unknown.to_excel('help.xlsx')
    




