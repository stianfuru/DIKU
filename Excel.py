from os import sep
import pandas as pd
import nltk
import sklearn 
import string
from nltk.corpus import stopwords

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')
df.dropna(inplace = True)

kolonner = ['Læringsutbytte - Kunnskap','Læringsutbytte - Ferdigheter','Læringsutbytte - Generell Kompetanse']
keywords = ['fluidstatistik', 'programmering', 'bygningsfysiske']

def text_process(frame):
    nopunc = [char for char in frame if char not in string.punctuaton]
    nopunc = ''.join(nopunc)
    nopunc = [word for word in nopunc.split() if word.lower not in stopwords.words('english')]
    nopunc = [word for word in nopunc.split() if word.lower not in stopwords.words('norwegian')]
    return nopunc

text_process(df)

print(df.describe())

#her har kamalan kommentert 
#test 2 - push

#df_final = df.drop(kolonner, axis=1)

#print(df.groupby('Læringsutbytte - Kunnskap').describe())

#md = open("resultat.md", "w+")
#md.write(str(df_final))             