import pandas as pd

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')


keywords = ['fluidstatikk', 'programmering']

output =  df.query('"fluidstatikk" in "Læringsutbytte - Kunnskap"')

print(output)

#md = open("resultat.md", "w+")
#md.write(str(df))