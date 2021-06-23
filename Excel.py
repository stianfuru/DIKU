import pandas as pd
import nltk
import string
import re
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import CountVectorizer

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')
df.dropna(inplace = True)

#kolonner = ['Læringsutbytte - Kunnskap','Læringsutbytte - Ferdigheter','Læringsutbytte - Generell Kompetanse']
keywords = ['Fluid', 'Dynamiske systemer'] #søkeord


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
    
    
   
    unique = 0
    for _ in keywords:
        i = 0 #teller for celle
        print(keywords[k]+':')
        md.write(keywords[k]+':\n')
        for _ in frame:
            #bow_transformer = CountVectorizer(analyzer=text_process).fit(frame[i])
            #unique += (len(bow_transformer.vocabulary_))
            j = 0 #teller for ord
            arraystr = ' '.join(map(str, frame[i])) #setter sammen igjen meldingen for printing 
            search = re.search(keywords[k].lower(),arraystr.lower()) #søkefunskjon
            if str(search) != 'None': #sjekker at det er match
                #print(frame[i][j]+' '+ Emnekode[i])
                                   
                print(Emnekode[i]+': '+arraystr + '\n') #printer ut emnekode og meldingen
                md.write(Emnekode[i]+': '+arraystr+'\n\n') #skriver det samme til resultat.md
                break                
            i = i + 1
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