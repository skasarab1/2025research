import pandas as pd

# Load data
psse_data = pd.read_excel('In_PSSEData.xlsx', sheet_name='PSSE_Buses')
bus_locations = pd.read_excel('EI_Bus_loca.xlsx')

# Standardize column names
psse_bus_numbers = psse_data['Bus  Number'].astype(int)
located_bus_numbers = bus_locations['BusNumber'].astype(int)

# Find missing buses
missing_buses = psse_data[~psse_data['Bus  Number'].isin(located_bus_numbers)]

# Save to file
missing_buses.to_excel('Missing_Buses.xlsx', index=False)


