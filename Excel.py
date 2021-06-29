import pandas as pd
#import nltk
import string
import re
from openpyxl import load_workbook
from nltk.corpus import stopwords
#from sklearn.feature_extraction.text import CountVectorizer

df = pd.read_excel('Diku.xlsx', sheet_name='DIKU', usecols='B,C,O,P,Q')
df.dropna(inplace = True)

#kolonner = ['Læringsutbytte - Kunnskap','Læringsutbytte - Ferdigheter','Læringsutbytte - Generell Kompetanse']
keywords = ['digital tvilling','virtuell',' VR[- ]',' AR[- ]',' XR[- ]','hololens','big room','revit','programmvare','trimble'
,' BIM[- ]','digital samhand','digitalisering','modell','kunstlig intelligens',' ICE[- ]',' VDC[- ]','samtidig prosjektering'
,' IPD[- ]','lean', 'maskinlæring',' AI[- ]',' IFC[- ]'] #søkeord

tegnsetting = """!"#$%&'()*+,./:;<=>?@[\]^_`{|}~'"""

def text_process(frame): 
    nopunc = [char for char in frame if char not in tegnsetting]
    nopunc = ''.join(nopunc)   
    nopunc = [word for word in nopunc.split() if word.lower not in stopwords.words('norwegian')]
    return nopunc

#tokenized og prosessert versjon av læringsutbytte kolonnene
LUK = df['Læringsutbytte - Kunnskap'].apply(text_process)
LUF = df['Læringsutbytte - Ferdigheter'].apply(text_process)
LUG = df['Læringsutbytte - Generell Kompetanse'].apply(text_process)

wb = load_workbook(filename='resultat.xlsx')
ws = wb.active
wb.save(filename='resultat.xlsx')

Emnekode = df['Emnekode']

count_max = 0
actual_max = 0

md = open("resultat.md", "w+")

def search_in_frame(frame, k):
    i = 0 #indeks for celle
    count = 0
    global count_max
    global actual_max
    for _ in frame:
        
        #bow_transformer = CountVectorizer(analyzer=text_process).fit(frame[i])
        #unique += (len(bow_transformer.vocabulary_))
        arraystr = ' '.join(map(str, frame[i])) #setter sammen igjen meldingen for printing 
        search = re.search(keywords[k].lower(),arraystr.lower()) #søkefunskjon
        if str(search) != 'None': #sjekker at det er match                                  
            print(Emnekode[i]+': '+arraystr+'\n') #printer ut emnekode og meldingen
            md.write(Emnekode[i]+': '+arraystr+'\n\n') #skriver det samme til resultat.md
            count = count + 1
            i = i + 1
            continue
        i = i + 1
    print(str(count)+' treff av 48 mulige\n') #printer ut antall treff
    md.write(str(count)+' treff av 48 mulige\n\n')
    count_max += count #legger til dette i max-count for keyword
    if str(frame) == str(LUK):
        ws.cell(k+2,2,count) #skriver til excel ark
    elif str(frame) == str(LUF):
       ws.cell(k+2,3,count)
    else: #sjekker at det er siste frame
        print(str(count_max)+' treff av totalt 144 mulige på søkeordet: '+keywords[k]) #printer ut max_count
        md.write(str(count_max)+' treff av totalt 144 mulige på søkeordet: '+keywords[k]+'\n')
        ws.cell(k+2,4,count)
        ws.cell(k+2,5,count_max)
        actual_max += count_max
        count_max = 0
       
def wordsearch(k):
    p = 0   #indeks for frame
    for _ in range(3):
        if p == 0: #går første gjennom LUK
            print('LUK:')
            md.write('LUK: \n')
            search_in_frame(LUK, k)
        elif p == 1: #så LUF
            print('LUF:')
            md.write('LUF: \n')
            search_in_frame(LUF, k)
        else: #til sist LUG
            print('LUG:')
            md.write('LUG: \n')
            search_in_frame(LUG, k)
        p = p + 1

def main():
    k = 0 #indeks for keyword

    for _ in keywords:
        if keywords[k].startswith(' '):
            x = keywords[k].replace(' ','',1)
            print('\n'+x+':')
            md.write('\n'+x+':\n')
            ws.cell(k+2,1,x)
        else:
            print('\n'+keywords[k]+':')
            md.write('\n'+keywords[k]+':\n')
            ws.cell(k+2,1,keywords[k])
        wordsearch(k) #kjører søk
        k = k + 1 #neste søkeord

    max_mulige = 144 * len(keywords)

    print('\n\n'+str(actual_max)+' treff av totalt '+str(max_mulige)+' mulige.')
    md.write('\n\n'+str(actual_max)+' treff av totalt '+str(max_mulige)+' mulige.')
    ws.cell(2,7,actual_max)
    ws.cell(2,13,max_mulige)

    for _ in range(ws.max_row): #fjerner ord fra excel-arket som ikke er i keywords
        if ws.cell(1,len(keywords)+2) != None:
            ws.delete_rows(len(keywords)+2)
            
    wb.save(filename='resultat.xlsx')    



main()


#TO-DO: hvis det er treff på søkeord, lag ark for søke ordet med emnekode som det er treff på