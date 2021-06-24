import pandas as pd
#import nltk
import string
import re
from nltk.corpus import stopwords
#from sklearn.feature_extraction.text import CountVectorizer

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')
df.dropna(inplace = True)

#kolonner = ['Læringsutbytte - Kunnskap','Læringsutbytte - Ferdigheter','Læringsutbytte - Generell Kompetanse']
keywords = ['digital tvilling', 'virtuell', ' vr ', ' ar ', ' xr ','hololens','big room','revit','programmvare','trimble'
,' bim ','digital samhand','digitalisering','modell','kunstlig intelligens',' ice ',' vdc ','concurrent','engineering',' ipd ','lean', 'maskinlæring',' ai ',' ifc '] #søkeord


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


def wordsearch(k):
    p = 0   #indeks for frame
    for _ in range(3):
        if p == 0:
            print('LUK:')
            md.write('LUK: \n')
            search_in_frame(LUK, k)
        elif p == 1:
            print('LUF:')
            md.write('LUF: \n')
            search_in_frame(LUF, k)
        else: 
            print('LUG:')
            md.write('LUG: \n')
            search_in_frame(LUG, k)
        p = p + 1

count_max = 0


def search_in_frame(frame, k):
    i = 0 #indeks for celle
    count = 0
    global count_max
    for _ in frame:
        
        #bow_transformer = CountVectorizer(analyzer=text_process).fit(frame[i])
        #unique += (len(bow_transformer.vocabulary_))
        arraystr = ' '.join(map(str, frame[i])) #setter sammen igjen meldingen for printing 
        search = re.search(keywords[k].lower(),arraystr.lower()) #søkefunskjon
        if str(search) != 'None': #sjekker at det er match                                  
            print(Emnekode[i]+': '+arraystr+'\n') #printer ut emnekode og meldingen
            md.write(Emnekode[i]+': '+arraystr+'\n\n') #skriver det samme til resultat.m
            count = count + 1
            i = i + 1
            continue
        i = i + 1
    print(str(count)+' treff av 48 mulige\n')
    md.write(str(count)+' treff av 48 mulige\n\n')
    count_max += count
    if str(frame) == str(LUG):
        print(str(count_max)+' treff av totalt 144 mulige på søkeordet: '+keywords[k])
        md.write(str(count_max)+' treff av totalt 144 mulige på søkeordet: '+keywords[k]+'\n')
        count_max = 0
#print(unique)
#print(words)
md = open("resultat.md", "w+")

def main():
    k = 0 #indeks for keyword
    for _ in keywords:
        print('\n'+keywords[k]+':')
        md.write('\n'+keywords[k]+':\n')
        wordsearch(k)
        k = k + 1
        

main()