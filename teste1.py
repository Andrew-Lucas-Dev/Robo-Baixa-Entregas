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
import re


caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)

def click_selenium(selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")

def remover_hora(data_str):
    if pd.notna(data_str) and isinstance(data_str, str):
        return data_str[:-5]
    return None

def click_image(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def baixada_ou_nao(yes, ok, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    yes_path = os.path.join(current_dir, caminho_imagem, yes) 
    ok_path = os.path.join(current_dir, caminho_imagem, ok) 
    while True:
        try:
            position = pyautogui.locateOnScreen(yes_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Nota sera baixada.")
                break
        except Exception as e:
            print("Imagem 'yes' não encontrada na tela. Aguardando...")
        
        try:
            position = pyautogui.locateOnScreen(ok_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Nota ja baixada.")
                return True
        except Exception as e:
            print("Imagem 'ok' não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)
    return False

def imagem_encontrada(image_path, confidence=0.9, max_attempts=5):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_path)
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Nota encontrada na tela.")
                return True
        except Exception as e:
            print("Nota não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

    click_image('CTRC.png')
    for i in range(2):
        pyautogui.press("down")
    pyautogui.sleep(0.5)
    pyautogui.press("enter")   
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Nota encontrada na tela é uma OST.")
                return True
        except Exception as e:
            print("Nota não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)  
    
    click_image('cancelar.png')
    click_image('digitar_data.png')
    for i in range(10):
        pyautogui.press("backspace")
    pyautogui.write(str(data_nota_fiscal))
    pyautogui.sleep(1)  
    click_nota('digitar_nota.png')
    pyautogui.sleep(1)
    pyautogui.write('00')
    pyautogui.sleep(0.2)
    pyautogui.write(str(nf))
    pyautogui.sleep(1)
    click_image('atualizar.png')
    pyautogui.sleep(1)
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Nota encontrada na tela.")
                return True
        except Exception as e:
            print("Nota não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)
    print("Número máximo de tentativas atingido. Nota não encontrada.")
    return False
 
def verificar_campo(image_name, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_name)    
    while True:
        found = False
        for i in range(5): 
            try:
                position = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if position:
                    print("Campo preenchido.")
                    found = True
                    break
            except Exception as e:
                print(f"Erro ao procurar o campo: {e}")
            print("Campo não preenchido. Aguardando...")
            pyautogui.sleep(1)       
        if found:
            break
        else:
            click_image('salvar_filial.png')
            for i in range(5):
                pyautogui.press("backspace")
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            for i in range(5):
                pyautogui.press("backspace")
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            for i in range(5):
                pyautogui.press("backspace")            
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")

def click_nota(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                center_x += 90  # Adicionar deslocamento para a direita
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def image_erro(image_name,image_name2,image_name3,image_name4,confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_name)
    image_path2 = os.path.join(caminho_imagem, image_name2)
    image_path3 = os.path.join(caminho_imagem, image_name3)
    image_path4 = os.path.join(caminho_imagem, image_name4)
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)         
            if position:
                click_image('ok_aviso.png')
                try:
                    position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
                    if position:
                        center_x = position3.left + position3.width // 2
                        center_y = position3.top + position3.height // 2
                        pyautogui.moveTo(center_x, center_y)  # Movendo o cursor para a posição da imagem               
                        pyautogui.moveRel(-80, 0)  # Movendo o cursor para cima
                        pyautogui.click()  # Clicando no local da imagem
                        for i in range(2): 
                            pyautogui.press("tab")
                        click_image('efetuar_baixa.png')
                        print("Imagem de erro encontrada na tela.")
                        break
                except Exception as e:
                    print("Imagem não encontrada na tela. Aguardando...")
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                 break
        except Exception as e:
            print("Imagem de erro não encontrada na tela.")
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                print("Imagem de Baixa efetuada")
                break
        except Exception as e:
            print("Imagem de Baixa não encontrada na tela.")
        try:
            position4 = pyautogui.locateOnScreen(image_path4, confidence=confidence)
            if position4:
                print("Imagem de Baixa efetuada")
                break
        except Exception as e:
            print("Imagem de Baixa não encontrada na tela.")       
        pyautogui.sleep(1)

def erro_efetuar(image_path, confidence=0.9, max_attempts=5):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, r'IMAGENS', image_path)
    attempts = 0
    while attempts < max_attempts:
        try:
            position = pyautogui.locateOnScreen(caminho_imagem, confidence=confidence)
            if position:
                click_image('OK.png')
                pyautogui.sleep(2)
                click_image('OK.png')
                click_image('efetuar_baixa.png')
                break
        except Exception as e:
            print(f"Imagem não encontrada na tela. Tentativa {attempts + 1} de {max_attempts}. Aguardando...")
        attempts += 1
        pyautogui.sleep(1)

def atualizar_ocorrencia(image_path, confidence=0.9, max_attempts=5):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, r'IMAGENS', image_path)
    attempts = 0
    while attempts < max_attempts:
        try:
            position = pyautogui.locateOnScreen(caminho_imagem, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Baixa de ocorrencia feita com sucesso.")
                break
        except Exception as e:
            print(f"Imagem da baixa nao encontrada. Tentativa {attempts + 1} de {max_attempts}. Aguardando...")
        attempts += 1
        pyautogui.sleep(1)

def verificar_ok_final(image_path, confidence=0.9, max_attempts=5):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, r'IMAGENS', image_path)
    attempts = 0
    while attempts < max_attempts:
        try:
            position = pyautogui.locateOnScreen(caminho_imagem, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Ok final encontrado com sucesso.")
                break
        except Exception as e:
            print(f"Imagem do ok nao encontrada. Tentativa {attempts + 1} de {max_attempts}. Aguardando...")
        attempts += 1
        pyautogui.sleep(1)

def check_caps_lock():
    return ctypes.windll.user32.GetKeyState(0x14) & 0xffff != 0

def alt_press(key):
    pyautogui.keyDown('alt')
    pyautogui.press(key)
    pyautogui.keyUp('alt')

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



# Função para extrair o número completo do MDF-e e formatá-lo
def extrair_mdf_e(mdf_e):
    if pd.isna(mdf_e):
        return None
    match = re.search(r"(?:1/2|2/1|2/2|/5/1|5/2)/([\d.]+)", str(mdf_e))
    return match.group(1).replace(".", "") if match else None

# Função para extrair o número da filial (antes da barra)
def extrair_filial(mdf_e):
    if pd.isna(mdf_e):
        return None
    match = re.search(r"MDF-e: (\d+)/", str(mdf_e))
    return int(match.group(1)) if match else None


# Função comum para processar a planilha
def processar_planilha(nome_arquivo, colunas_para_remover, coluna_data, coluna_mdf_filial):
    Planilha = pd.read_excel(nome_arquivo)
    Planilha.drop(columns=colunas_para_remover, inplace=True)

    Planilha["MDF-e"] = Planilha[coluna_mdf_filial].where(
        Planilha[coluna_mdf_filial].astype(str).str.contains("MDF-e:", na=False)
    ).fillna(method="ffill")
    
    Planilha["Filial"] = Planilha["MDF-e"]
    Planilha["MDF-e"] = Planilha["MDF-e"].apply(extrair_mdf_e)
    Planilha["Filial"] = Planilha["Filial"].apply(extrair_filial)

    Planilha['N° NF'] = pd.to_numeric(Planilha['N° NF'], errors='coerce')
    Planilha.dropna(subset=['N° NF'], inplace=True)
    Planilha['N° NF'] = Planilha['N° NF'].astype(int)

    Planilha = Planilha.rename(columns={'N° NF': 'NF', 'Data NF': 'DATA NOTA FISCAL', 'Data': 'Data Chegada', 'DATA': 'Data Chegada', 'Status da entrega': 'STATUS'})
    Planilha['Data Entrega'] = Planilha['Data Chegada']
    Planilha['Fim Descarreg.'] = Planilha['Data Chegada']

    Planilha['DATA NOTA FISCAL'] = pd.to_datetime(Planilha['DATA NOTA FISCAL']).dt.strftime('%d/%m/%Y')

    colunas_de_data = ['Data Entrega', 'Fim Descarreg.']
    for coluna in colunas_de_data:
        processar_coluna_data(Planilha, coluna)
    processar_datas(Planilha, colunas_de_data)
    processar_coluna_chegada(Planilha, 'Data Chegada')

    Planilha = Planilha.dropna(axis=1, how='all')
    
    if 'STATUS' in Planilha.columns:
        Planilha.loc[Planilha['STATUS'] == 'FINALIZADO', 'STATUS'] = 'Entregue'
    try:
        Planilha.drop(columns='Plano', inplace=True)
    except:
        pass

    return Planilha

# Parâmetros de colunas a remover
colunas_para_remover_cc19 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Vlr Merc.',
    'Entrega Canhoto Físico', 'Imagem Salva - Com Erro', 'NF com problema', 'Status da baixa', 'Data Emissão Ct-e'
]
colunas_para_remover_cc15 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Vlr Merc.',
    'Entrega Canhoto Físico', 'Status da baixa'
]
colunas_para_remover_cc16 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Manifesto', 'EM BRANCO', 'Bairro'
]
colunas_para_remover_cc21 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Vlr Merc.', 'Entrega Canhoto Físico', 'Manifesto'
]

# Processar cada planilha
Planilha_CC19 = processar_planilha("CC19.xlsx", colunas_para_remover_cc19, 'N° Carga', 'N° Carga')
Planilha_CC15 = processar_planilha("CC15.xlsx", colunas_para_remover_cc15, 'N° Carga', 'N° Carga')
Planilha_CC16 = processar_planilha("CC16.xlsx", colunas_para_remover_cc16, 'Plano', 'Plano')
Planilha_CC21 = processar_planilha("CC21.xlsx", colunas_para_remover_cc21, 'N° Carga', 'N° Carga')

# Exibir as planilhas processadas (opcional)
# print(Planilha_CC19)
# print(Planilha_CC15)
# print(Planilha_CC16)
# print(Planilha_CC21)

BASE_DADOS = pd.read_excel("BASE_DADOS.xlsx")
BASE_DADOS = BASE_DADOS.dropna(axis=1, how='all')
#print(BASE_DADOS)

# # # # Juntar as 4 planilhas
combined_df = pd.concat([BASE_DADOS,Planilha_CC19,Planilha_CC15,Planilha_CC16,Planilha_CC21], ignore_index=True)
# combined_df = combined_df[combined_df['STATUS'] == 'Entregue']
combined_df = combined_df.drop_duplicates(subset='NF', keep='first')
combined_df['BAIXADO'] = combined_df['BAIXADO'].fillna('NAO')
#print(combined_df)

# Obter os manifestos únicos
manifestos_unicos = combined_df["MDF-e"].unique()

resultado = []  # Lista para salvar os manifestos ou status

# Processar cada manifesto
for man in manifestos_unicos:
    linhas_manifesto = combined_df[combined_df["MDF-e"] == man]
    
    if (linhas_manifesto["STATUS"] == "Entregue").all():
        resultado.append({"Manifesto": man, "Filial": linhas_manifesto["Filial"].iloc[0]})
    else:
        resultado.append({"Manifesto": "NAO BAIXAR", "Filial": linhas_manifesto["Filial"].iloc[0]})
        
        # Baixar notas individuais
        print(f"Baixando notas do manifesto {man}...")
        for _, linha in linhas_manifesto.iterrows():
            if linha["STATUS"] == "Entregue":
                print(f"Filial:{linha['Filial']} Nota:{linha['NF']} ")

# Exibir resultados
for item in resultado:
    if item["Manifesto"] != "NAO BAIXAR":
        print(f"Filial:{item['Filial']} Manifesto:{item['Manifesto']} ")
