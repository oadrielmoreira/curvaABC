#importando bibliotecas
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import datetime as dt


#importando databases
db = pd.read_excel('#####')
curvaABC = pd.read_excel('#####')



#TRATANDO A BASE DB CONSOLIDADO

#Criando Referência Curta que será utilizada para localizar o produto na Curva ABC
db['Referência Curta'] = db['Referência'].str[:13]
db['DataA'] = db['Data'].dt.strftime('%Y-%m')
db['DataA'] = pd.to_datetime(db['DataA'], format='%Y-%m')

#Tirando as devoluções
db = db.drop(db.loc[db['QTD']<0].index)


#CRIANDO E ADICIONANDO AS COLUNAS DOS MESES

#Definindo a variável Today
def today():
    return dt.date.today()

#DF dos meses
meses = pd.date_range(start='2021-06-01', end= today(), freq='MS')
nomes_colunas = [mes.strftime('%Y-%m') for mes in meses]
df_meses = pd.DataFrame(columns=nomes_colunas)

#Adicionando df_meses
curvaABC = pd.concat([curvaABC, df_meses])



#Somando as quantidades [QTD] a partir da igualdade de ANO/MÊS e Referência Curta (SKU do produto)
for col in curvaABC.columns[8:]:
    mes_ano = pd.to_datetime(col, format='%Y-%m')
    mes = mes_ano.month
    ano = mes_ano.year
    curvaABC[col] = curvaABC['Referência Curta'].apply(lambda ref: db.loc[(db['Referência Curta'] == ref) & (db['Data'].dt.month == mes) & (db['Data'].dt.year == ano), 'QTD'].sum())


#CRIANDO A COLUNA TOTAL

# somando as últimas 12 colunas e criando a coluna "Total"
curvaABC['Total'] = curvaABC.iloc[:, -13:-1].sum(axis=1)

# ordenando em ordem decrescente pela coluna "Total"
curvaABC = curvaABC.sort_values(by='Total', ascending=False)



#CRIANDO A COLUNA PERCENTUAL
curvaABC['Percentual'] = curvaABC['Total'] / curvaABC['Total'].sum() 


#CRIANDO A COLUNA PERCENTUAL ACUMULADA
curvaABC['Percentual Acumulado'] = curvaABC['Percentual'].cumsum() / curvaABC['Percentual'].sum()


#CRIANDO A CLASSIFICAÇÃO ABC
def classificar(valor):
    if valor < 0.50:
        return 'A'
    elif valor < 0.80:
        return 'B'
    else:
        return 'C'

curvaABC['Curva'] = curvaABC['Percentual Acumulado'].apply(classificar)



curvaABC.to_excel('Curva ABC.xlsx', index = False)
