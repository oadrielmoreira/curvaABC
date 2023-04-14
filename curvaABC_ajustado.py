#importando bibliotecas utilizadas no projeto
import pandas as pd
from datetime import datetime, timedelta
import datetime as dt
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

#importando database
db = pd.read_excel('###')
curvaABC = pd.read_excel('###')
#OBS: db é o dataframe que contém o database de vendas da empresa, enquanto o curvaABC contém uma base de infomações de todos os produtos que serão analisados


#TRATANDO A BASE DB CONSOLIDADO
#Criando Referência Curta e o Código sem a distinção de tributação que será utilizada para localizar o produto na Curva ABC
db['Código'] = db['Código'].str.slice(stop=11)
db['Referência Curta'] = db['Referência'].str[:13]
#Tratando a Data para formar uma coluna apenas com anos e meses
db['DataA'] = db['Data'].dt.strftime('%Y-%m')
db['DataA'] = pd.to_datetime(db['DataA'], format='%Y-%m')
#Tirando as devoluções
db = db.drop(db.loc[db['QTD']<0].index)


#CRIANDO E ADICIONANDO AS COLUNAS DOS MESES NO DATAFRAME curvaABC
#Definindo a variável Today, pois queremos que essa análise leve em consideração o dia a extração 
def today():
    return dt.date.today()
#DF dos meses indicando a criação de colunas com Anos e Meses a partir do primeiro mês monitorado "2021-06-01" termninando na função today() descrita acima
meses = pd.date_range(start='2021-06-01', end= today(), freq='MS')
nomes_colunas = [mes.strftime('%Y-%m') for mes in meses]
df_meses = pd.DataFrame(columns=nomes_colunas)
#Concatenando o dataframe df_meses no curvaABC
curvaABC = pd.concat([curvaABC, df_meses])


#Somando as quantidades [QTD] com duas condicionais: ANO/MÊS (nas colunas) e o PN (SKU do produto que está nas linhas)
for col in curvaABC.columns[6:]:
    mes_ano = pd.to_datetime(col, format='%Y-%m')
    mes = mes_ano.month
    ano = mes_ano.year
    curvaABC[col] = curvaABC['PN'].apply(lambda ref: db.loc[(db['Código'] == ref) & (db['Data'].dt.month == mes) & (db['Data'].dt.year == ano), 'QTD'].sum())


#CRIANDO A COLUNA TOTAL
#Somando as últimas 12 colunas, desconsiderando a do mês vigente da extração dos dados, e criando a coluna "Total"
curvaABC['Total'] = curvaABC.iloc[:, -13:-1].sum(axis=1)
#Ordenando em ordem decrescente pela coluna "Total"
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

#Exportando resultado em excel
curvaABC.to_excel('Curva ABC - AJUSTADA.xlsx', index = False)



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
#O resultado é um gráfico de colunas com o Total e Percentual de cada Produto nas Curvas A, B e C


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
#O resultado é um gráfico de colunas com o Total e Percentual de Vendas de cada Produto nas Curvas A, B e C


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
#O objetivo desse gráfico é analisar a amplitude histórica da quantidade vendida no período analisado na Curva A


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
#O objetivo desse gráfico é analisar a amplitude histórica da quantidade vendida no período analisado na Curva B


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
#O objetivo desse gráfico é analisar a amplitude histórica da quantidade vendida no período analisado na Curva C


#Criando DataFrame product
product = pd.DataFrame({
    'PN': curvaABC['PN'],
    '2021-06' : curvaABC['2021-06'],
    '2021-07' : curvaABC['2021-07'],
    '2021-08' : curvaABC['2021-08'],
    '2021-09' : curvaABC['2021-09'],
    '2021-10' : curvaABC['2021-10'],
    '2021-11' : curvaABC['2021-11'],
    '2021-12' : curvaABC['2021-12'],
    '2022-01' : curvaABC['2022-01'],
    '2022-02' : curvaABC['2022-02'],
    '2022-03' : curvaABC['2022-03'],
    '2022-04' : curvaABC['2022-04'],
    '2022-05' : curvaABC['2022-05'],
    '2022-06' : curvaABC['2022-06'],
    '2022-07' : curvaABC['2022-07'],
    '2022-08' : curvaABC['2022-08'],
    '2022-09' : curvaABC['2022-09'],
    '2022-10' : curvaABC['2022-10'],
    '2022-11' : curvaABC['2022-11'],
    '2022-12' : curvaABC['2022-12'],
    '2023-01' : curvaABC['2023-01'],
    '2023-02' : curvaABC['2023-02'],
    '2023-02' : curvaABC['2023-03'],
})


#Fazendo a transposição do dataframe (colunas para linhas)
product = product.set_index('PN').T

#CRIANDO BOXPLOT DOS PRODUTOS SELECIONADOS A PARTIR DO DATAFRAME product
#Criando uma figura com um único eixo de plotagem
fig, ax = plt.subplots(figsize=(10, 6))
#Escolhendo os produtos pelo SKU
curvaA = ['RZ.AU.KR.08', 'RZ.MO.DA.30','RZ.MO.DA.31','RZ.AU.KR.07','RZ.AC.PL.02','RZ.AU.KR.19']
#Criando Boxplot
sns.boxplot(data=product[curvaA], orient='v', ax=ax)
ax.set_title('Boxplots de Vendas por Produtos Curva A', fontsize=18)
ax.set_xlabel('Produto', fontsize=14)
ax.set_ylabel('Quantidade de Vendas', fontsize=14)
