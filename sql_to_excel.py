import pandas as pd
import cx_Oracle
from query import clob_sql
import xlsxwriter
linha = '-------------------------------------------------------------------------------'
def usuario():
    hosts = []
    arq = open('usu.txt','r')
    lines = arq.readlines()
    arq.close()
    for l in range(len(lines)):
        lines[l] = lines[l].replace('\n','')
        hosts.append(lines[l])
    
    return lines[0], lines[1], hosts
user, psw, hosts = usuario() 


nome = 'nome_arquivo'
#hosts = ['DATABASE']
writer = pd.ExcelWriter(f'{nome}.xlsx', engine='xlsxwriter')
for x in range(len(hosts)):
    try:
        conn = cx_Oracle.connect(user, psw, hosts[x] , encoding="UTF-8")
        print(f'Entrou na {hosts[x]}!')
        print(linha)
        df = pd.read_sql(clob_sql, con=conn)
        print('Carregou o DF.')
        print(linha)
        df.to_excel(writer, sheet_name=hosts[x], index=False)
        print(f'Exportou para EXCEL a filial {hosts[x]}!')
        print(linha)
        conn.close()
    except:
        print('Erro!')
        
df2 = pd.DataFrame([[clob_sql]], columns=['SQL'])
df2.to_excel(writer, sheet_name='SQL', index=False)
writer.save()

