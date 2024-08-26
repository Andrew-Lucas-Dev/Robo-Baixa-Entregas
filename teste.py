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
from datetime import datetime, timedelta

time = 0.5
caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)


Planilha_CC19 = pd.read_excel("planilhaderotascc19.xlsx")
colunas_para_remover = ['Série', 'Cnpj cliente', 'Status da baixa','Cliente','Cidade','Ct-e/OST','Peso','Qtde','Vlr Merc.','Entrega Canhoto Físico','NF com problema','Imagem Salva - Com Erro']
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
#print(Planilha_CC19['Data Chegada'])