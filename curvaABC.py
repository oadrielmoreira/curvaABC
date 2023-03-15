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


#Exportando Planilha com classificação ABC
curvaABC.to_excel('Curva ABC.xlsx', index = False)




#Criando tabela de frequência e gráfico representando a quantidade de produtos por classificação

tabfre = curvaABC['Curva'].value_counts()
tabfre = pd.DataFrame({'Curva': tabfre.index, 'Frequências': tabfre.values})
tabfre['Percentual'] = tabfre['Frequências'] / tabfre['Frequências'].sum() *100
tabfre = tabfre.iloc[::-1].reset_index(drop=True)

clf = tabfre['Curva']
freq = tabfre['Frequências']

cores = ['#00a60b', '#f7d200', '#828282']
plt.bar(clf, freq, color=cores)

for i, (v1, v2) in enumerate(zip(tabfre['Frequências'], tabfre['Percentual'])):
    plt.text(i, v1+1, 'Total: ' + f'{v1}\n Percentual: {v2:.1f}%\n', ha='center')
    
plt.ylim([0, max(tabfre['Frequências']) * 1.2])



plt.title('Quantidade de produtos por classificação')
plt.xlabel('Curva')
plt.ylabel('Quantidade')



#Criando dataframe de quantidade de produtos vendidos por classificação e criando gráfico de barras que indica a quantidade de produtos vendidos por classificação

df_soma = curvaABC.groupby('Curva')['Total'].sum().reset_index()
df_soma['Percentual'] = df_soma['Total'] / df_soma['Total'].sum() *100

clf = df_soma['Curva']
freq = df_soma['Total']

plt.bar(df_soma['Curva'], df_soma['Total'], color=cores)

for i, (v1, v2) in enumerate(zip(df_soma['Total'], df_soma['Percentual'])):
    plt.text(i, v1+1, 'Total: ' f'{v1}\n Percentual: {v2:.1f}%\n', ha='center')

plt.ylim([0, max(df_soma['Total']) * 1.2])


plt.title('Total de Produtos Vendidos por Classificação')
plt.xlabel('Classificações')
plt.ylabel('Total de Produtos Vendidos')

plt.show()



#Criando Gráfico de Linhas do comportamento dos meses analisados dos produtos A
dados = curvaABC.iloc[:, -17:-5]
dados_A = dados[curvaABC['Curva'] == 'A']

fig = plt.figure(figsize=(12, 6))
plt.plot(dados_A.T)

plt.grid(axis='y')
plt.ylim(0, 3000)

    
plt.title('Dados dos Produtos Classificados como A')
plt.xlabel('Mês')
plt.ylabel('Quantidade')
plt.show()


#Criando Gráfico de Linhas do comportamento dos meses analisados dos produtos B

dados = curvaABC.iloc[:, -17:-5]
dados_B = dados[curvaABC['Curva'] == 'B']

fig = plt.figure(figsize=(12, 6))
plt.plot(dados_B.T)

plt.grid(axis='y')
plt.ylim(0, 3000)

    
plt.title('Dados dos Produtos Classificados como B')
plt.xlabel('Mês')
plt.ylabel('Quantidade')
plt.show()



#Criando Gráfico de Linhas do comportamento dos meses analisados dos produtos C

dados = curvaABC.iloc[:, -17:-5]
dados_C = dados[curvaABC['Curva'] == 'C']

fig = plt.figure(figsize=(12, 6))
plt.plot(dados_C.T)

plt.grid(axis='y')
plt.ylim(0, 3000)

    
plt.title('Dados dos Produtos Classificados como C')
plt.xlabel('Mês')
plt.ylabel('Quantidade')
plt.show()
