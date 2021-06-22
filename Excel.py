import pandas as pd
import nltk
import string
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import CountVectorizer

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')
df.dropna(inplace = True)

kolonner = ['Læringsutbytte - Kunnskap','Læringsutbytte - Ferdigheter','Læringsutbytte - Generell Kompetanse']
keywords = ['fluidstatikk', 'betong', 'matematikk']

def text_process(frame):
    nopunc = [char for char in frame if char not in string.punctuation]
    nopunc = ''.join(nopunc)   
    nopunc = [word for word in nopunc.split() if word.lower not in stopwords.words('norwegian')]
    return nopunc

#tokeized og prosessert versjon av læringsutbytte kolonnene
LUK = df['Læringsutbytte - Kunnskap'].apply(text_process)
LUF = df['Læringsutbytte - Ferdigheter'].apply(text_process)
LUG = df['Læringsutbytte - Generell Kompetanse'].apply(text_process)

Emnekode = df['Emnekode']


def wordsearch(frame):
    j = 0
    k = 0
    i = 0
    unique = 0
    words = 0
    for _ in keywords:
        
        for _ in frame:
            bow_transformer = CountVectorizer(analyzer=text_process).fit(frame[i])
            unique += (len(bow_transformer.vocabulary_))
            for _ in frame[i]:
                if keywords[k] == frame[i][j]:    
                    print(frame[i][j]+' '+ Emnekode[i])
                j = j + 1
            j = 0
            i = i + 1
        i = 0
        k = k + 1


wordsearch(LUG) 
#print(unique)
#print(words)
#print(LUK[0][0])
#df_final = df.drop(kolonner, axis=1)

#print(df.groupby('Læringsutbytte - Kunnskap').describe())

#md = open("resultat.md", "w+")
#md.write(str(df_final))             