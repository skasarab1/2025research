import pandas as pd

# File paths
PSSE_Data = 'In_PSSEData.xlsx'
Bus_Locations = 'EI_Bus_loca.xlsx'
ISOne = 'ISONE.xlsx'
NYISO = 'NYISO.xlsx'
NYISOq = 'NYISO-Interconnection-Queue.xlsx'
UNKNOWN = 'TOBEFILLED.xlsx'

#Read Files
PSSE = pd.read_excel(PSSE_Data, sheet_name='PSSE_Lines')
Bus_Loc = pd.read_excel(Bus_Locations)
ISONE = pd.read_excel(ISOne)
NYSIOQ = pd.read_excel(NYISOq)
Unknown = pd.read_excel(UNKNOWN)

print(Unknown.head(4))
print(Bus_Loc.head(4))
print(PSSE.head(4))

Unknown['Closest Name'] = 'X'
Unknown['Closest Lat'] = 'X'
Unknown['Closest Lon'] = 'X'
Unknown['N Levels'] = 'X'

Bus_NUM_startRow = Unknown[Unknown['Closest Name'] == 'X'].iloc[0]
print(Bus_NUM_startRow)
Bus_NUM_start = Bus_NUM_startRow['Bus  Number']

Bus_NUMS_second = PSSE[PSSE['From Bus  Number'] == 'Bus_NUM_start']['From Bus  Name'].tolist()

print(Bus_NUM_start)
print(Bus_NUMS_second)

