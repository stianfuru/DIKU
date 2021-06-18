import pandas as pd

data = pd.read_excel('Diku.xlsx', sheet_name='DIKU')

bool_series = pd.isnull(data["Emneansvarlig"])
print(data[bool_series])
