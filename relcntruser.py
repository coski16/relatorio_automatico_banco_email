#Criacao de relatorio automatico em xlsx do banco de dados mysql#

#Packages Necessarios 
import mysql.connector
from mysql.connector import Error
import numpy as np
import pandas as pd
from datetime import date
import emailenvio #servico de envio de e-mail na mesma pasta

#listas
CNTR = [] 
DIA = []
MES = []
ANO = []
USER = []
HOJE = date.today()
diaatual = data_em_texto = '{}'.format(HOJE.day)
mestual = data_em_texto = '{}'.format(HOJE.month)
anotual = data_em_texto = '{}'.format(HOJE.year)

#CRIANDO CONEXAO COM O BANCO
Host = 'xxx.xxx.x.x' #ip do servidor
User = 'root' #usuario root
Password = 'senha' #senha
Database = 'database' #database
conexao_db = mysql.connector.connect(host = Host , user = User, password = Password , database = Database, auth_plugin = 'mysql_native_password')#inicia conexao com o banco
cursor = conexao_db.cursor() #define o cursou do mysql
query = """SELECT A.nracond AS CNTR,DAY(A.dtpesagem) AS Dia ,MONTH(A.dtpesagem) AS Mes,YEAR(A.dtpesagem) AS Ano,U1.nomeUser AS USER FROM cotton.acond AS A
INNER JOIN cotton.os AS O
    ON O.codos = A.codos
INNER JOIN cotton.user AS U1
    ON U1.codUser = A.codUser
    WHERE O.safra LIKE 2022 
    ;"""#query do banco
cursor.execute(query) #executa a query do banco
dados = cursor.fetchall() #retorna o resultado da query
for i in range(0,len(dados)):
            for j in range(0,1):
                CNTR.append(str(dados[i][j]))
for i in range(0,len(dados)):
            for j in range(1,2):
                DIA.append(dados[i][j])
for i in range(0,len(dados)):
            for j in range(2,3):
                MES.append(str(dados[i][j]))                   
for i in range(0,len(dados)):
            for j in range(3,4):
                ANO.append(str(dados[i][j]))   
for i in range(0,len(dados)):
            for j in range(4,5):
                USER.append(str(dados[i][j]))   
        
tabela0 = pd.DataFrame({'CONTAINER':CNTR,'DIA':DIA,'MES': MES,'ANO': ANO ,'USUARIO':USER}) #cria a tabela em dataframe

tabela1 = tabela0.sort_values(['MES']) 

tabela1_filter = tabela1[tabela1['MES'].isin([str(mestual)])] #fitro
tabela1_filter2 = tabela1[tabela1['ANO'].isin([str(anotual)])] #filtro

tabela3 = tabela1_filter2.pivot_table(index=['USUARIO'], values= 'CONTAINER',columns=['MES'], aggfunc='count').fillna(0) #cria duas dataframes com a partir do filtro
tabela4 = tabela1_filter.pivot_table(index=['USUARIO'], values= 'CONTAINER',columns=['DIA'], aggfunc='count').fillna(0)  

tabela3.to_excel('qtdcrtmes.xlsx',index=True) #exporta o arquivo em xlsx (excel)
print('Arquivo criado qtdcrtmes.xlsx') #confirma no terminal que o arquivo gerou
tabela4.to_excel('qtdcrtdia.xlsx',index=True)
print('Arquivo criado qtdcrtdia.xlsx')

emailenvio.enviar_email() #chama o servico de envio de e-mail automatico

#Mauricio Lascoski -

