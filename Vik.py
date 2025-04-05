import pandas as pd

# File paths
PSSE_Data = 'In_PSSEData.xlsx'
Bus_Locations = 'EI_Bus_loca.xlsx'
ISOne = 'ISONE.xlsx'
NYISO = 'NYISO.xlsx'
UNKNOWN = 'TOBEFILLED.xlsx'

#Read Files
PSSE = pd.read_excel(PSSE_Data, sheet_name='PSSE_Lines')
Bus_Loc = pd.read_excel(Bus_Locations)
ISONE = pd.read_excel(ISOne)
Unknown = pd.read_excel(UNKNOWN)

#Filter kVs
Unknown = Unknown[(Unknown['Base kV'] >= 69) & (Unknown['Base kV'] <= 999)]

Unknown['Closest Name'] = None
Unknown['Closest Lat'] = None
Unknown['Closest Lon'] = None
Unknown['N Levels'] = None

for idx, row in Unknown.head(4).iterrows():
    BusNum = row['Bus  Number'] #Bus Number from Unknown Data set
    BusName = row['Bus  Name'] #Bus Name from Unknown Data set

    MatchingRows = PSSE[
        (PSSE['From Bus  Number'] == BusNum) |
        (PSSE['To Bus  Number'] == BusNum)
    ]

    for idx, line in MatchingRows.iterrows():
        ConnectedBuses = [] #empty data set for found buses
        if line['From Bus  Number'] == BusNum:
            ConnectedBuses.append(line['To Bus  Number'])
        else:
            ConnectedBuses.append(line['From Bus  Number'])
        
    NumConnections = len(set(ConnectedBuses))
    print(NumConnections)
    




