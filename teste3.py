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
                print("Imagem foi encontrada na tela.")
                
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
            break
        pyautogui.sleep(1)

click_image('referencia.png')
click_image('digitar_data.png')
click_image('digitar_nota.png')