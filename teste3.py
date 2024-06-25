import pandas as pd
from openpyxl import load_workbook
import os
import numpy as np
import pyautogui
import ctypes
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys          
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
import pyautogui
from selenium.webdriver.common.action_chains import ActionChains
import shutil


caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)


def remover_hora(data_str):
    if pd.notna(data_str) and isinstance(data_str, str):
        return data_str[:-5]
    return None


def processar_coluna_data(df, coluna):
    df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
    df[coluna] = df[coluna].dt.date
    df[coluna] = df[coluna].astype(str).str.strip()  # Converter para string e remover espaços em branco
    df[coluna] = pd.to_datetime(df[coluna], format='%Y-%m-%d', errors='coerce')  # Converter para datetime
    df[coluna] = df[coluna] + pd.Timedelta(hours=22)  # Adicionar o horário '22:00'
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y %H:%M')  # Formatar no formato desejado

def processar_coluna_chegada(df, coluna):
    df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
    df[coluna] = df[coluna] + pd.Timedelta(hours=22)
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')

def processar_datas(df, colunas):
    for coluna in colunas:
        if 'Data Entrega' in coluna:
            df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=1)
        elif 'Fim Descarreg.' in coluna:
            df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=2)
        df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')
        
def formatar_datas(df, colunas):
    for coluna in colunas:
        df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
        df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')

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
Planilha_CC15['DATA NOTA FISCAL'] = pd.to_datetime(Planilha_CC15['DATA NOTA FISCAL'])
Planilha_CC15['DATA NOTA FISCAL'] = Planilha_CC15['DATA NOTA FISCAL'].dt.strftime('%d/%m/%Y')
colunas_de_data = ['Data Entrega', 'Fim Descarreg.']
for coluna in colunas_de_data:
    processar_coluna_data(Planilha_CC15, coluna) 
processar_datas(Planilha_CC15, colunas_de_data)
processar_coluna_chegada(Planilha_CC15,'Data Chegada')
Planilha_CC15 = Planilha_CC15.dropna(axis=1, how='all')
#print(Planilha_CC15)

Planilha_CC19 = pd.read_excel("planilhaderotascc19.xlsx")
colunas_para_remover = ['Série', 'Cnpj cliente', 'N° Carga', 'Status da baixa','Cliente','Cidade','Ct-e/OST','Peso','Qtde','Vlr Merc.','Entrega Canhoto Físico','NF com problema','Imagem Salva - Com Erro']
Planilha_CC19.drop(columns=colunas_para_remover, inplace=True)
Planilha_CC19['N° NF'] = pd.to_numeric(Planilha_CC19['N° NF'], errors='coerce')
Planilha_CC19.dropna(subset=['N° NF'], inplace=True)
Planilha_CC19['N° NF'] = Planilha_CC19['N° NF'].astype(int)
Planilha_CC19 = Planilha_CC19.rename(columns={'N° NF': 'NF', 'Data NF': 'DATA NOTA FISCAL', 'Data': 'Data Chegada', 'Status da entrega': 'STATUS'})
Planilha_CC19['Data Entrega'] = Planilha_CC19['Data Chegada']
Planilha_CC19['Fim Descarreg.'] = Planilha_CC19['Data Chegada']
Planilha_CC19['DATA NOTA FISCAL'] = pd.to_datetime(Planilha_CC19['DATA NOTA FISCAL'])
Planilha_CC19['DATA NOTA FISCAL'] = Planilha_CC19['DATA NOTA FISCAL'].dt.strftime('%d/%m/%Y')
colunas_de_data = ['Data Entrega', 'Fim Descarreg.']
for coluna in colunas_de_data:
    processar_coluna_data(Planilha_CC19, coluna)
processar_datas(Planilha_CC19, colunas_de_data)
processar_coluna_chegada(Planilha_CC19,'Data Chegada')
Planilha_CC19 = Planilha_CC19.dropna(axis=1, how='all')
#print(Planilha_CC19)


for i, linha in enumerate(Planilha_CC15.index):
    data_chegada = Planilha_CC15.loc[linha, "Data Chegada"]
    data_entrega = Planilha_CC15.loc[linha, "Data Entrega"]   
    data_fim_descarregamento =  Planilha_CC15.loc[linha, "Fim Descarreg."]  
    data_nota = Planilha_CC15.loc[linha, "DATA NOTA FISCAL"]  
    nota = Planilha_CC15.loc[linha, "NF"]
    print(f'nota:{nota} data:{data_nota}')
