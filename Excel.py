import pandas as pd

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')
df.dropna(inplace = True)


keywords = ['fluidstatikk', 'programmering', 'bygningsfysiske']

output =  df['Læringsutbytte - Kunnskap'].str.contains(keywords[0])
df_drop = df[output].drop('Læringsutbytte - Kunnskap', axis=1)
df_drop1 = df_drop.drop('Læringsutbytte - Ferdigheter', axis=1)
df_final = df_drop1.drop('Læringsutbytte - Generell Kompetanse', axis=1)
print(df_final)

md = open("resultat.md", "w+")
md.write(str(df_final))