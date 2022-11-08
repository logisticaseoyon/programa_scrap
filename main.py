def mont_fab_1(data_entrada, data_final, dir, ignore):
    import pandas as pd
    import numpy as np
    from tkinter.filedialog import askopenfilename
    import datetime
    from datetime import date
    import os
    import time


    # Obtendo lista com arquivos Montagem fabrica 1
    arq = dir
    lista_nomes = []
    ignorar = ignore
    for _, _, arquivo in os.walk(arq):
        for i in arquivo:
            if i in ignorar:
                pass
            else:
                lista_nomes.append(i) 

    lista_tabelas = []
    for i in lista_nomes:
        try:
            df = pd.read_excel(arq+'/'+i,sheet_name = 'PRODUÇÃO',header = 2)
        except:
            print("Erro ao importar ", i)
            input("Pressione enter")
            print("Processando...")
        # Filtra coluna data com datas acima da data informada
        df['DATA'] = pd.to_datetime(df['DATA'], format = '%Y-%m-%d')
        df = df[(df['DATA'] >= data_entrada) & (df['DATA'] <= data_final)]
        df['DATA'] = df['DATA'].dt.strftime('%d/%m/%Y')

        # Verifica se existe a coluna "Ordem de produção", se não, cria uma
        if "ORDEM DE PRODUÇÃO" in df.columns:
            pass
        else:
            df['ORDEM DE PRODUÇÃO'] = ""

        # Correções em tabelas diferentes
        colunas_nomes = ['Unnamed: 0', 'DATA', 'Código', 'QUANTIDADE', 'DX/SX', 'Nome', 'PADRÃO',
           'HSTD', 'FALHA DE INJEÇÃO', 'REBARBAS | FUROS OBSTRUÍDOS',
           'CORTE DE FACA', 'MANCHAS', 'BOLHAS', 'QUEBRA OU TRINCA',
           'RETENÇÃO OU RASPAGEM', 'CHUPAGEM', 'DEFORMAÇÃO', 'ARRANHÕES', 'BRILHO',
           'MATERIAL CONTAMINADO', 'SOLDA DEFICIENTE',
           'LIMPEZA DE ROSCA NA TROCA DE COR', 'BORRA NA TEXTURA',
           'REFUGO DE SETUP', 'TOTAL', 'Porcentagem', 'ORDEM DE PRODUÇÃO']

        if df.columns[11] == 'nan.1':
            pass
            df.columns = lista_tabelas[-1].columns
        elif 'BORRA NA TEXTURA.1' in df.columns:
            df.rename(columns={'BORRA NA TEXTURA.1': 'REFUGO DE SETUP'}, inplace = True)
        elif 'REFUGO DE SETUP' not in df.columns:
            df['REFUGO DE SETUP'] = 0
            df = df[colunas_nomes]
        df.columns = colunas_nomes

        lista_tabelas.append(df)

    # Concatena tabelas em 1
    df_completo = pd.concat(lista_tabelas, axis = 0, ignore_index = True)

    # Montando tabelas com colunas na ordem correta

    colunas = ['DATA','ORDEM DE PRODUÇÃO','Código','QUANTIDADE',
             'FALHA DE INJEÇÃO','REBARBAS | FUROS OBSTRUÍDOS','CORTE DE FACA',
             'MANCHAS','BOLHAS','QUEBRA OU TRINCA','RETENÇÃO OU RASPAGEM',
             'CHUPAGEM','DEFORMAÇÃO','ARRANHÕES','BRILHO','MATERIAL CONTAMINADO',
             'SOLDA DEFICIENTE','LIMPEZA DE ROSCA NA TROCA DE COR','BORRA NA TEXTURA',
             'REFUGO DE SETUP','TOTAL']

    df_completo = df_completo[colunas]
    df_completo['Código'] = ""

    df_completo.fillna(0,inplace = True)

    # Somando colunas para obter coluna "TOTAL"
    cols_numericas = df_completo.iloc[:,4:20]
    df_completo['TOTAL'] = cols_numericas.sum(axis = 1)
    df_completo['QUANTIDADE'] = ""
    df_completo = df_completo[df_completo['TOTAL'] != 0]
    return df_completo

#-----------------------------------------------------------------------------------------------

def inj_fab_1(data_entrada, data_final, dir, ignore):
    import numpy as np
    from tkinter.filedialog import askopenfilename
    import datetime
    from datetime import date
    import os
    import time
    import pandas as pd


    # Obtendo lista com arquivos injeção fabrica 1
    arq = dir
    lista_nomes = []
    ignorar = ignore
    for _, _, arquivo in os.walk(arq):
        for i in arquivo:
            if i in ignorar:
                pass
            else:
                lista_nomes.append(i) 

    lista_tabelas = []
    for i in lista_nomes:
        try:
            df = pd.read_excel(arq+'/'+i,sheet_name = 'PRODUÇÃO',header = 2)
        except:
            print("Erro ao importar ", i)
            input("Pressione enter")
            print("Processando...")
        # Filtra coluna data com datas acima da data informada
        try:
            df['DATA'] = pd.to_datetime(df['DATA'], format = '%Y-%m-%d')
            df = df[(df['DATA'] >= data_entrada) & (df['DATA'] <= data_final)]
            df['DATA'] = df['DATA'].dt.strftime('%d/%m/%Y')
        except:
            print("Erro data ", i)

        # Verifica se existe a coluna "Ordem de produção", se não, cria uma
        if "ORDEM DE PRODUÇÃO" in df.columns:
            pass
        else:
            df['ORDEM DE PRODUÇÃO'] = ""

        # Padroniza nomes de colunas
        if 'DESCRIÇÃO PRODUTO' in df.columns:
            df.rename(columns={'DESCRIÇÃO PRODUTO': 'DESCRIÇÃO DO PRODUTO'}, inplace = True)
        elif 'DESCRIÇÃO DA PEÇA' in df.columns:
            df.rename(columns={'DESCRIÇÃO DA PEÇA': 'DESCRIÇÃO DO PRODUTO'}, inplace = True)
        elif 'DESSCRIÇÃO DO PRODUTO' in df.columns:
            df.rename(columns={'DESSCRIÇÃO DO PRODUTO': 'DESCRIÇÃO DO PRODUTO'}, inplace = True)
        elif 'BORRA NA TEXTURA.1' in df.columns:
            df.rename(columns={'BORRA NA TEXTURA.1': 'REFUGO DE SETUP'}, inplace = True)

        lista_tabelas.append(df)
    col_names = lista_tabelas[0].columns
    for i in lista_tabelas:
        i.columns = col_names
    # Concatena tabelas em uma
    df_completo = pd.concat(lista_tabelas, axis = 0, ignore_index = True)

    # Montando tabelas com colunas na ordem correta

    colunas = ['DATA','ORDEM DE PRODUÇÃO','CÓDIGO.1','MOLDE',
    'FALHA DE INJEÇÃO','REBARBAS | FUROS OBSTRUÍDOS','CORTE DE FACA',
    'MANCHAS','BOLHAS','QUEBRA OU TRINCA','RETENÇÃO OU RASPAGEM',
    'CHUPAGEM','DEFORMAÇÃO','ARRANHÕES','BRILHO','MATERIAL CONTAMINADO',
    'SOLDA DEFICIENTE','LIMPEZA DE ROSCA NA TROCA DE COR','BORRA NA TEXTURA',
    'REFUGO DE SETUP','TOTAL']

    df_completo = df_completo[colunas]
    df_completo['CÓDIGO.1'] = ""

    df_completo.fillna(0,inplace = True)

    # Somando colunas para obter coluna "TOTAL"
    cols_numericas = df_completo.iloc[:,4:20]
    df_completo['TOTAL'] = cols_numericas.sum(axis = 1)
    #df_completo['QUANTIDADE'] = ""
    df_completo = df_completo[df_completo['TOTAL'] != 0]
    return df_completo



#-----------------------------------------------------------------------------------------------------

def mont_fab_2(data_entrada, data_final, dir, ignore):
    import pandas as pd
    import numpy as np
    import datetime
    from datetime import date
    import os
    import time


    # Obtendo lista com arquivos Montagem fabrica 2
    arq = dir
    lista_nomes = []
    ignorar = ignore
    for _, _, arquivo in os.walk(arq):
        for i in arquivo:
            if i in ignorar:
                pass
            else:
                lista_nomes.append(i) 

    lista_tabelas = []
    for i in lista_nomes:
        try:
            df = pd.read_excel(arq+'/'+i,sheet_name = 'PRODUÇÃO',header = 2)

            # Filtra coluna data com datas acima da data informada
            df['DATA'] = pd.to_datetime(df['DATA'], format = '%Y-%m-%d')
            df = df[(df['DATA'] >= data_entrada) & (df['DATA'] <= data_final)]
            df['DATA'] = df['DATA'].dt.strftime('%d/%m/%Y')
            
            # Verifica se existe a coluna "Ordem de produção", se não, cria uma
            if "ORDEM DE PRODUÇÃO" in df.columns:
                pass
            else:
                df['ORDEM DE PRODUÇÃO'] = ""
            lista_tabelas.append(df)
        except:
            print("Erro ao importar ", i)
            input("Pressione enter")
            print("Processando...")
            
    col_names = lista_tabelas[0].columns
    for i in lista_tabelas:
        i.columns = col_names

    # Concatena tabelas em 1
    df_completo = pd.concat(lista_tabelas, axis = 0, ignore_index = True)

    # Montando tabelas com colunas na ordem correta

    colunas = ['DATA','ORDEM DE PRODUÇÃO','Código','QUANTIDADE',
               'FALHA DE INJEÇÃO','REBARBAS | FUROS OBSTRUÍDOS','CORTE DE FACA',
               'MANCHAS','BOLHAS','QUEBRA OU TRINCA','RETENÇÃO OU RASPAGEM',
               'CHUPAGEM','DEFORMAÇÃO','ARRANHÕES','BRILHO','MATERIAL CONTAMINADO',
               'SOLDA DEFICIENTE','LIMPEZA DE ROSCA NA TROCA DE COR','BORRA NA TEXTURA',
               'REFUGO DE SETUP','TOTAL']

    df_completo = df_completo[colunas]
    df_completo['Código'] = ""

    df_completo.fillna(0,inplace = True)

    # Somando colunas para obter coluna "TOTAL"
    cols_numericas = df_completo.iloc[:,4:20]
    df_completo['TOTAL'] = cols_numericas.sum(axis = 1)
    df_completo['QUANTIDADE'] = ""
    df_completo = df_completo[df_completo['TOTAL'] != 0]
    return df_completo

#------------------------------------------------------------------------------------------------


def inj_fab_2(data_entrada, data_final, dir):
    import pandas as pd
    import numpy as np
    import os
    from datetime import datetime


    # Obtendo lista com arquivos inj fabrica 2
    arq = dir
    nomes_arqs = []
    for _, _, arquivo in os.walk(arq):
        for i in arquivo:
            nomes_arqs.append(i)


    lista_arqs = []
    for i in nomes_arqs:
        if len(i) == 13:
            if (pd.to_datetime(datetime.strptime(i[:-5], '%Y%m%d').date()) - data_entrada).days >= 0:
                if (pd.to_datetime(datetime.strptime(i[:-5], '%Y%m%d').date()) - data_final).days <= 0:
                    lista_arqs.append(i)


    dfs_concat = []
    if len(lista_arqs) > 0:
        for i in lista_arqs:
            data = pd.to_datetime(datetime.strptime(i[:-5], '%Y%m%d').date())
            df = pd.read_excel(arq+"/"+ i, header = 6)
            
            # Limpeza dos dados
            df[['OP','Máquina']] = df[['OP','Máquina']].fillna(method = 'ffill')
            ops = df['OP'].unique()
            dfs_ops = []
            for i in ops:
                dfs_ops.append(df[(df['OP'] == i)  & (df['Refugo'] != 'Total')].dropna(subset=['Refugo']))
            for i in dfs_ops:
                i['tip_ref'] = i['Refugo'].apply(lambda x: x.split('-')[0])


            # Monta novo df alocando as informações nas colunas correspondentes
            cols = ['DATA','ORDEM DE PRODUÇÃO','Código','MOLDE','FALHA DE INJEÇÃO','REBARBAS | FUROS OBSTRUÍDOS','CORTE DE FACA',
            'MANCHAS','BOLHAS','QUEBRA OU TRINCA','RETENÇÃO OU RASPAGEM','CHUPAGEM','DEFORMAÇÃO','ARRANHÕES','BRILHO',
            'MATERIAL CONTAMINADO','SOLDA DEFICIENTE','LIMPEZA DE ROSCA NA TROCA DE COR','BORRA NA TEXTURA','REFUGO DE SETUP']
            df_final = pd.DataFrame(columns = cols)

            mat = []
            for i in dfs_ops:
                    if i.shape[0] != 0:
                        for y in range(len(i)):
                            if i.iloc[y,-1:].item() == '1 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0], i.iloc[y,-5:-4].item(),"","","","","","","","","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '2 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"",i.iloc[y,-5:-4].item(),"","","","","","","","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '3 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","",i.iloc[y,-5:-4].item(),"","","","","","","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '4 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","",i.iloc[y,-5:-4].item(),"","","","","","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '5 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","",i.iloc[y,-5:-4].item(),"","","","","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '6 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","",i.iloc[y,-5:-4].item(),"","","","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '7 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","",i.iloc[y,-5:-4].item(),"","","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '8 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","",i.iloc[y,-5:-4].item(),"","","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '9 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","",i.iloc[y,-5:-4].item(),"","","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '10 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","","",i.iloc[y,-5:-4].item(),"","","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '11 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","","","",i.iloc[y,-5:-4].item(),"","","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '12 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","","","","",i.iloc[y,-5:-4].item(),"","","",""]
                                mat.append(vet)
                            elif i.iloc[y,-1:].item() == '13 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","","","","","",i.iloc[y,-5:-4].item(),"","",""]
                                mat.append(vet)                 
                            elif i.iloc[y,-1:].item() == '14 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","","","","","","",i.iloc[y,-5:-4].item(),"",""]
                                mat.append(vet)                   
                            elif i.iloc[y,-1:].item() == '15 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","","","","","","","",i.iloc[y,-5:-4].item(),""]
                                mat.append(vet)                     
                            elif i.iloc[y,-1:].item() == '16 ':
                                vet = [data, i.reset_index()['OP'][0],"",i.reset_index()['Máquina'][0],"","","","","","","","","","","","","","","",i.iloc[y,-5:-4].item()]
                                mat.append(vet)
                            else:
                                pass

            df_final = pd.DataFrame(mat,columns=cols)


            # Converte colunas numericas
            num_cols = ['ORDEM DE PRODUÇÃO','FALHA DE INJEÇÃO','REBARBAS | FUROS OBSTRUÍDOS','CORTE DE FACA',
            'MANCHAS','BOLHAS','QUEBRA OU TRINCA','RETENÇÃO OU RASPAGEM','CHUPAGEM','DEFORMAÇÃO','ARRANHÕES','BRILHO',
            'MATERIAL CONTAMINADO','SOLDA DEFICIENTE','LIMPEZA DE ROSCA NA TROCA DE COR','BORRA NA TEXTURA','REFUGO DE SETUP']
            for i in num_cols:
                df_final[i] = pd.to_numeric(df_final[i])

            # Agrupa dados por OP
            df_final = df_final.groupby(['DATA','ORDEM DE PRODUÇÃO','Código','MOLDE']).sum().reset_index()
            dfs_concat.append(df_final)


        dfs_concat = pd.concat(dfs_concat)
        dfs_concat['DATA'] = dfs_concat['DATA'].apply(lambda x: x.strftime('%d/%m/%Y'))
        dfs_concat.rename(columns={'MOLDE': 'QUANTIDADE'}, inplace = True)
        return dfs_concat
    else:
        print("Inj Fab 2 não possui arquivos para data informada")
        input("Pressione enter")
        print("Processando...")
        return None


#------------------------------------------------------------------------------------------------

import pandas as pd
import numpy as np
from tkinter.filedialog import askopenfilename
import datetime
from datetime import date
import os
import time

# Configuração de variaveis
import pandas as pd
variaveis = pd.read_csv('variaveis.csv', sep = ';')
dir_inj_fab_1 = variaveis['diretorio_inj_fab_1'][0].replace('\\','/')
dir_inj_fab_2 = variaveis['diretorio_inj_fab_2'][0].replace('\\','/')
dir_mont_fab_1 = variaveis['diretorio_mont_fab_1'][0].replace('\\','/')
dir_mont_fab_2 = variaveis['diretorio_mont_fab_2'][0].replace('\\','/')
exceptions_inj_fab_1 = variaveis['excecoes_inj_fab_1'].values
exceptions_inj_fab_2 = variaveis['excecoes_inj_fab_2'].values
exceptions_mont_fab_1 = variaveis['excecoes_mont_fab_1'].values
exceptions_mont_fab_2 = variaveis['excecoes_mont_fab_2'].values


print("----------------- FAVOR INFORMAR CASO OCORRA A ADIÇÃO DE UMA NOVA PLANILHA -----------------")
input("Pressione enter")

x = 0
while x == 0:
    dia = int(input("Data inicial (dia): "))
    mes = int(input("Data inicial (mês): "))
    ano = int(input("Data inicial (ano): "))

    print("Confirmar data: ", dia, "/", mes,"/", ano, "  ?")
    confirma = input("Pressione, 'ENTER' para confirmar ou 'N' para corrigir. ")
    if confirma == "N" or confirma == "n":
        pass
    else:
        data_entrada = pd.to_datetime(datetime.date(ano,mes,dia))
        x = 1

informar = input("Informar data final? \n'S': Sim\n'N': Não\n")
if informar == 'S' or informar == 's':
    x = 0
    while x == 0:
        print("Digite a data final: ")
        dia_2 = int(input("Data final (dia): "))
        mes_2 = int(input("Data final (mês): "))
        ano_2 = int(input("Data final (ano): "))

        print("Confirmar data: ", dia_2, "/", mes_2,"/", ano_2, "  ?")
        confirma = input("Pressione, 'ENTER' para confirmar ou 'N' para corrigir. ")
        if confirma == "N" or confirma == "n":
            pass
        else:
            data_final = pd.to_datetime(datetime.date(ano_2,mes_2,dia_2))
            if np.timedelta64(data_final - data_entrada, 'D').astype(int) > 0:             
                x = 1
            else:
                input("Data final deve ser maior que a inicial. Pressione 'ENTER' para digitar novamente.")
                
else:
    data_final = pd.to_datetime(datetime.datetime.now())
    
print('Processando...')



antes = datetime.datetime.now()

mont_fab_2 = mont_fab_2(data_entrada, data_final, dir_mont_fab_2, exceptions_mont_fab_2)
inj_fab_1 = inj_fab_1(data_entrada, data_final,dir_inj_fab_1, exceptions_inj_fab_1)
mont_fab_1 = mont_fab_1(data_entrada, data_final, dir_mont_fab_1, exceptions_mont_fab_1)
inj_fab_2 = inj_fab_2(data_entrada, data_final, dir_inj_fab_2)

depois = datetime.datetime.now()

# Ajustes
inj_fab_1.rename(columns={'CÓDIGO.1': 'Código'}, inplace = True)
inj_fab_1.rename(columns={'MOLDE': 'QUANTIDADE'}, inplace = True)
# Concatenação
df_completo = pd.concat([mont_fab_1,mont_fab_2,inj_fab_1, inj_fab_2], axis = 0, ignore_index = True)
# Ajuste coluna Molde
df_completo.rename(columns={'QUANTIDADE': 'MAQUINA'}, inplace = True)

# substitui zeros
df_completo = df_completo.replace(0,"")

# Muda formato de coluna 'DATA' para datetime
df_completo['DATA'] = pd.to_datetime(df_completo['DATA'], format = '%d/%m/%Y' )

# Elimila coluna total
df_completo.drop('TOTAL', axis = 1, inplace = True)
# Exporta
df_completo.to_excel('C:/temp/refugo.xlsx', index = False)

# Tempo de execução
print('Finalizado!')
print("Tempo de execução: ",depois - antes)
time.sleep(2)

from tkinter import messagebox

messagebox.showinfo('Mensagem', \
      'Base atualizada com sucesso')


os.system('start C:/temp/refugo.xlsx')
