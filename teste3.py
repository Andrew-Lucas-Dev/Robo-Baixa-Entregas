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
multi = False

# def confirmacao_preenchido(image_path,image_path2,image_path3, confidence=0.9):
#     current_dir = os.path.dirname(__file__)  # Diretório atual do script
#     caminho_imagem = caminho + r'\IMAGENS'
#     image_path = os.path.join(current_dir, caminho_imagem, image_path)
#     image_path2 = os.path.join(current_dir, caminho_imagem, image_path2) 
#     image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)  
#     while True:
#         try:
#             position = pyautogui.locateOnScreen(image_path, confidence=confidence)
#             position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
#             position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)

#             if position:
#                 print('Algum filtro 1 nao foi digitado.')
#             # if position2:
#             #     print('Algum filtro 2 nao foi digitado.')
#             # if position3:
#             #     print('Algum filtro 3 nao foi digitado.')
#         except:
#             print("Filtros todos digitados.")
#             break  # Saia do loop em caso de erro inesperado
7
#         pyautogui.sleep(1)

# confirmacao_preenchido('referencia.png','digitar_data.png','digitar_nota.png')

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

def confirmacao_preenchido(image_path, image_path2, image_path3, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    a, b, c = 0, 0, 0
    multi = True
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem 1 foi encontrada na tela.")
                a = 0  # Redefine 'a' se a imagem 1 for encontrada
                click_image('referencia.png')
                for i in range(10):
                    pyautogui.press("backspace")
                pyautogui.write(str(referencia))
            else:
                a = 1  # Define 'a' como 1 se a imagem 1 não for encontrada
        except Exception as e:
            print("Erro ao procurar imagem 1:", e)
            a = 1
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                print("Imagem 2 foi encontrada na tela.")
                b = 0  # Redefine 'b' se a imagem 2 for encontrada
                click_image('digitar_data.png')
                for i in range(10):
                    pyautogui.press("backspace")
                pyautogui.write(str(data_nota_fiscal))
                pyautogui.sleep(time)  
            else:
                b = 1  # Define 'b' como 1 se a imagem 2 não for encontrada
        except Exception as e:
            print("Erro ao procurar imagem 2:", e)
            b = 1
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position3:
                print("Imagem 3 foi encontrada na tela.")
                c = 0  # Redefine 'c' se a imagem 3 for encontrada
                click_nota('digitar_nota.png')
                pyautogui.sleep(time)
                if multi == True:
                    pyautogui.write('00')
                    pyautogui.sleep(0.2)
                pyautogui.write(str(nf))
                pyautogui.sleep(time)
            else:
                c = 1  # Define 'c' como 1 se a imagem 3 não for encontrada
        except Exception as e:
            print("Erro ao procurar imagem 3:", e)
            c = 1
        # Se todas as três imagens não forem encontradas, sai do loop
        if a == 1 and b == 1 and c == 1:
            print("Nenhuma das três imagens foi encontrada. Saindo do loop.")
            click_image('atualizar.png')
            pyautogui.sleep(time)
            break
        pyautogui.sleep(1)

referencia = 123456789
data_nota_fiscal = '25/07/2024'
nf = 121565165

confirmacao_preenchido('referencia.png', 'digitar_data.png', 'digitar_nota.png')