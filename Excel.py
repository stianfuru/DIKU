import pandas as pd
import nltk
import string
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import CountVectorizer

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')
df.dropna(inplace = True)

kolonner = ['Læringsutbytte - Kunnskap','Læringsutbytte - Ferdigheter','Læringsutbytte - Generell Kompetanse']
keywords = ['Fluidstatikk', 'betong', 'matematikk', 'væske']

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
    k = 0 #teller for keyword
    i = 0 #teller for celle
    j = 0 #teller for ord
   
    unique = 0
    for _ in keywords:
        
        for _ in frame:
            #bow_transformer = CountVectorizer(analyzer=text_process).fit(frame[i])
            #unique += (len(bow_transformer.vocabulary_))
            for _ in frame[i]:
                if keywords[k].lower() == frame[i][j].lower():    
                    #print(frame[i][j]+' '+ Emnekode[i])
                    arraystr = ' '.join(map(str, frame[i]))
                    print(Emnekode[i]+': '+arraystr)
                    md.write(Emnekode[i]+': '+arraystr+'\n')
                j = j + 1
            j = 0
            i = i + 1
        i = 0
        k = k + 1


 
#print(unique)
#print(words)
md = open("resultat.md", "w+")

def main():
    print('LUK: ')
    md.write('LUK: \n')
    wordsearch(LUK)
    print('LUF: ')
    md.write('LUF: \n')
    wordsearch(LUF)
    print('LUG:')
    md.write('LUG: \n')
    wordsearch(LUG)

main()