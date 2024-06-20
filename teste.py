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

def click_image(image_path, confidence=0.8):
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
                print("Imagem encontrada e clicada.")
                break
        except:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

click_image('logo_rodopar_areatrabalho_resumido.png')#PC ESCRITORIO
#click_image('logo_rodopar_areatrabalho.png')#PC CASA
pyautogui.click()
click_image('conectar_rodopar.png')
#click_image('conectar_rodopar1.png')
click_image('senha_rodopar_1.png')
#click_image('senha_rodopar_2.png')
pyautogui.write("17@mudar")
click_image('ok_primeiro_login.png')
#click_image('ok_primeiro_login2.png')
click_image('sim_primeiro_login.png')
#click_image('sim_primeiro_login2.png')
click_image('segundo_login.png')    
pyautogui.sleep(1)
pyautogui.write("anascimento")
pyautogui.press("tab")
pyautogui.write("990607")
for i in range(2): 
    pyautogui.press("enter")
click_image('filial_1.png')
pyautogui.press("enter")

click_image('botao_frota.png')
pyautogui.press("alt")
pyautogui.press("alt")
pyautogui.press("right")
for i in range(2): 
    pyautogui.press("down")
pyautogui.press("right")
for i in range(10): 
    pyautogui.press("down")
pyautogui.press("enter")
