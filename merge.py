import pandas as pd

current = pd.read_excel("Warren.xlsx")
working = pd.read_excel("Missing_Buses.xlsx")

current['Bus  Number'] = current['Bus  Number'].astype(str).str.strip()
working['Bus  Number'] = working['Bus  Number'].astype(str).str.strip()

merged = pd.merge(working, current[['Bus  Number', 'Lat', 'Long']], on = 'Bus  Number', how = 'left')

filled_count = merged['Lat'].notna().sum()
empty_count = merged['Lat'].isna().sum()

print(f"Filled bus locations: {filled_count}")
print(f"Missing bus locations: {empty_count}")

merged.toexcel('test.xlsx')







    



