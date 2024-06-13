import pyautogui
import os
import pandas as pd
import numpy as np

caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)

def processar_coluna_data(df, coluna):
    df[coluna] = df[coluna].astype(str).str.strip()  # Converter para string e remover espaços em branco
    df[coluna] = pd.to_datetime(df[coluna], format='%Y-%m-%d', errors='coerce')  # Converter para datetime
    df[coluna] = df[coluna] + pd.Timedelta(hours=22)  # Adicionar o horário '22:00'
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y %H:%M')  # Formatar no formato desejado

def processar_datas(df, colunas):
    for coluna in colunas:
        if 'Data' in coluna:
            df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=1)
        elif 'Fim' in coluna:
            df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=2)
        df[coluna] = df[coluna].dt.strftime('%d/%m/%Y %H:%M')

def formatar_datas(df, colunas):
    for coluna in colunas:
        df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
        df[coluna] = df[coluna].dt.strftime('%d/%m/%Y %H:%M')

# Carregar a planilha do Excel
Planilha_CC19 = pd.read_excel("planilhaderotascc19.xlsx")
colunas_para_remover = ['Série', 'Cnpj cliente', 'N° Carga', 'Status da baixa','Cliente','Cidade','Ct-e/OST','Peso','Qtde','Vlr Merc.','Entrega Canhoto Físico']
Planilha_CC19.drop(columns=colunas_para_remover, inplace=True)
Planilha_CC19['N° NF'] = pd.to_numeric(Planilha_CC19['N° NF'], errors='coerce')
Planilha_CC19.dropna(subset=['N° NF'], inplace=True)
Planilha_CC19['N° NF'] = Planilha_CC19['N° NF'].astype(int)
Planilha_CC19 = Planilha_CC19.rename(columns={'N05/06/2024 22:00° NF': 'NF', 'Data NF': 'DATA NOTA FISCAL', 'Data': 'Data Chegada', 'Status da entrega': 'STATUS'})
Planilha_CC19['Data Entrega'] = Planilha_CC19['Data Chegada']
Planilha_CC19['Fim Descarreg.'] = Planilha_CC19['Data Chegada']
Planilha_CC19 = Planilha_CC19.dropna(axis=1, how='all')
Planilha_CC19['DATA NOTA FISCAL'] = pd.to_datetime(Planilha_CC19['DATA NOTA FISCAL'])
Planilha_CC19['DATA NOTA FISCAL'] = Planilha_CC19['DATA NOTA FISCAL'].dt.strftime('%d/%m/%Y')
colunas_de_data = ['Data Chegada', 'Data Entrega', 'Fim Descarreg.']
for coluna in colunas_de_data:
    processar_coluna_data(Planilha_CC19, coluna)
colunas_de_data_horas = ['Data Entrega', 'Fim Descarreg.']
processar_datas(Planilha_CC19, colunas_de_data_horas)
#print(Planilha_CC19)

# Carregar a planilha do Excel
Planilha_CC15 = pd.read_excel("planilhaderotascc15.xlsx")
colunas_para_remover = ['Série', 'Cnpj cliente', 'N° Carga', 'Status da baixa','Cliente','Cidade','Ct-e/OST','Peso','Qtde','Vlr Merc.','Entrega Canhoto Físico']
Planilha_CC15.drop(columns=colunas_para_remover, inplace=True)
Planilha_CC15['N° NF'] = pd.to_numeric(Planilha_CC15['N° NF'], errors='coerce')
Planilha_CC15.dropna(subset=['N° NF'], inplace=True)
Planilha_CC15['N° NF'] = Planilha_CC15['N° NF'].astype(int)
Planilha_CC15 = Planilha_CC15.rename(columns={'N° NF': 'NF', 'Data NF': 'DATA NOTA FISCAL', 'Data': 'Data Chegada', 'Status da entrega': 'STATUS'})
Planilha_CC15['Data Entrega'] = Planilha_CC15['Data Chegada']
Planilha_CC15['Fim Descarreg.'] = Planilha_CC15['Data Chegada']
Planilha_CC15['STATUS'] = Planilha_CC15['STATUS'].fillna('EM ROTA')
Planilha_CC15 = Planilha_CC15.dropna(axis=1, how='all')
Planilha_CC15['DATA NOTA FISCAL'] = pd.to_datetime(Planilha_CC15['DATA NOTA FISCAL'])
Planilha_CC15['DATA NOTA FISCAL'] = Planilha_CC15['DATA NOTA FISCAL'].dt.strftime('%d/%m/%Y')
colunas_de_data = ['Data Chegada', 'Data Entrega', 'Fim Descarreg.']
for coluna in colunas_de_data:
    processar_coluna_data(Planilha_CC15, coluna)
Planilha_CC15['Data Entrega'] = pd.to_datetime(Planilha_CC15['Data Entrega'], format='%d/%m/%Y %H:%M')
Planilha_CC15['Fim Descarreg.'] = pd.to_datetime(Planilha_CC15['Fim Descarreg.'], format='%d/%m/%Y %H:%M')
processar_datas(Planilha_CC15, colunas_de_data_horas)
#print(Planilha_CC15)

# # # Carregar a planilha do Excel
Planilha_Bahia = pd.read_excel("EntregaT2.xlsx")
Planilha_Bahia = Planilha_Bahia.dropna(axis=1, how='all')
Planilha_Bahia = Planilha_Bahia[['NF', 'STATUS', 'DT NF', 'CHEGADA', 'FIM DESCARGA', 'ENTREGA']]
Planilha_Bahia = Planilha_Bahia[(Planilha_Bahia['STATUS'] == 'ATRASADA') | (Planilha_Bahia['STATUS'] == 'NO PRAZO')]
Planilha_Bahia = Planilha_Bahia.rename(columns={'DT NF': 'DATA NOTA FISCAL','CHEGADA': 'Data Chegada', 'FIM DESCARGA': 'Data Entrega', 'ENTREGA': 'Fim Descarreg.'})
Planilha_Bahia['NF'] = Planilha_Bahia['NF'].astype(np.int64)#problema do .0
Planilha_Bahia['STATUS'] = 'Entregue'
Planilha_Bahia['DATA NOTA FISCAL'] = pd.to_datetime(Planilha_Bahia['DATA NOTA FISCAL'])
Planilha_Bahia['DATA NOTA FISCAL'] = Planilha_Bahia['DATA NOTA FISCAL'].dt.strftime('%d/%m/%Y')
colunas_para_formatar = ['Data Chegada', 'Data Entrega', 'Fim Descarreg.']
formatar_datas(Planilha_Bahia, colunas_para_formatar)
#print(Planilha_Bahia)

BASE_DADOS = pd.read_excel("BASE_DADOS.xlsx")
BASE_DADOS = BASE_DADOS.dropna(axis=1, how='all')
#print(BASE_DADOS)

# # # # Juntar os dados das duas planilhas
combined_df = pd.concat([BASE_DADOS,Planilha_CC19, Planilha_CC15,Planilha_Bahia], ignore_index=True)
combined_df = combined_df[combined_df['STATUS'] == 'Entregue']
combined_df = combined_df.drop_duplicates(subset='NF', keep='first')
combined_df['BAIXADO'] = combined_df['BAIXADO'].fillna('NAO')
#print(combined_df)

numero_linhas = len(combined_df)

for i, linha in enumerate(combined_df.index):
    nf = combined_df.loc[linha, "NF"]
    data_nota_fiscal = combined_df.loc[linha, "DATA NOTA FISCAL"]
    data_chegada = combined_df.loc[linha, "Data Chegada"]   
    data_entrega = combined_df.loc[linha, "Data Entrega"]   
    data_fim_descarregamento =  combined_df.loc[linha, "Fim Descarreg."]  
    baixado = str(combined_df.loc[linha, "BAIXADO"])  
    #print(linha)
    if baixado == "SIM":
        continue
    else:    
        status = 'ENTREGUE'
        falta = numero_linhas - i 
        #print(f'nota:{nf} data da nota:{data_nota_fiscal} data chegada:{data_chegada} data entrega:{data_entrega} data fim descarregamento:{data_fim_descarregamento} falta:{falta}')
        pyautogui.sleep(2)
        pyautogui.write(str(data_chegada))

        



