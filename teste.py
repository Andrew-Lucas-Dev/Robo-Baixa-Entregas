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
         
    multi = True
    click_image('cancelar.png')
    click_image('referencia.png')
    for i in range(10):
        pyautogui.press("backspace")
    pyautogui.write(str(referencia))
    click_image('digitar_data.png')
    for i in range(15):
        pyautogui.press("backspace")
    pyautogui.write(str(data_nota_fiscal))
    pyautogui.sleep(1)  
    click_nota('digitar_nota.png')
    pyautogui.sleep(1)
    pyautogui.write('00')
    pyautogui.sleep(0.2)
    pyautogui.write(str(nf))
    pyautogui.sleep(1)
    confirmacao_preenchido('referencia.png','digitar_data.png','digitar_nota.png')
    for i in range(15):
        pyautogui.press("left")
    pyautogui.sleep(1)
    click_image('atualizar.png')
    pyautogui.sleep(1)
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Nota encontrada na tela.")
                confirmacao_preenchido('referencia.png','digitar_data.png','digitar_nota.png')
                return True
        except Exception as e:
            print("Nota não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)
    print("Número máximo de tentativas atingido. Nota não encontrada.")
    return False

if imagem_encontrada('nota_encontrada.png'):
    combined_df.loc[linha, "BAIXADO"] = "SIM" 