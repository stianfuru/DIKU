import pandas as pd

data = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,I')

bool_series = pd.isnull(data["Emneansvarlig"])

md = open("resultat.md", "w+")
md.write(str(data[bool_series]))
