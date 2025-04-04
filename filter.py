import pandas as pd

# File paths
PSSE_Data = 'In_PSSEData.xlsx'
Bus_Locations = 'EI_Bus_loca.xlsx'
ISOne = 'ISONE.xlsx'
NYISO = 'NYISO.xlsx'

PSSE = pd.read_excel(PSSE_Data, sheet_name='PSSE_Buses')
Bus_Loc = pd.read_excel(Bus_Locations)
ISONE = pd.read_excel(ISOne)
# Export to Excel


