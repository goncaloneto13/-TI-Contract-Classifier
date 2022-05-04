import pandas as pd
import sklearn.model_selection as sk
from sklearn.svm import LinearSVC
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from openpyxl import Workbook, load_workbook


#arq = 'BD - Descrição de Objetos.xlsx'
arq = 'Contratos de TI  da Consulta SQL.xlsx'
base_completa = pd.read_excel('BD - Descrição de Objetos.xlsx')
base_empenhos = pd.read_excel(arq)
objetos = base_completa['DESCRIÇÃO']
y = base_completa['TI']

vectorizer = TfidfVectorizer()
x = vectorizer.fit_transform(objetos)

X_train, X_test, y_train, y_test = sk.train_test_split(x, y, train_size=0.80, shuffle=True,random_state=42)

print('X_train',X_train.shape[0])
print('X_test',X_test.shape[0])

clfs = [LinearSVC()]

def classificar(clf, texto):

    teste = vectorizer.transform([texto])
    result = clf.predict(teste)

    return result    

wb = load_workbook(arq)
for sheet in wb:
    print(sheet)
planilha = wb.active

for clf in clfs:
    clf.fit(X_train, y_train)

    cont = 0
    contlic = 0  

    for i in y_test:
        if i == True:
            cont += 1
    score = clf.score(X_test,y_test)
    print(clf, end = ' - ')
    print(score)
    print(cont)

    linha = 2
    cont_ti = 0

    for obj in base_empenhos['OBJETO']:
       
        result = classificar(clf,obj)
        celula = 'P'+str(linha)
        planilha[celula] = str(result)
        linha += 1    
        #print(linha)

        if result == True:
            cont_ti += 1

    wb.save(arq)
    print(arq, cont_ti)
