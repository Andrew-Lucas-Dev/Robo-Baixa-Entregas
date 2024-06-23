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

def click_selenium(selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")

driver = webdriver.Chrome()
driver.get("https://jettatransporte-my.sharepoint.com/:f:/g/personal/jetta_bi_jettatransporte_onmicrosoft_com/EiA6eCcrmHVOi0SVjgVS4eYBTgW6NmdHNlvRSINLlAOW5g?e=qyl9wK")
driver.maximize_window()

click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/div[1]/span/span/button')


click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[3]/div/div/div[3]/div/div[1]/span/span/button')
try:
    print("Pasta Planilha CC19...")
    corpo_email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div/div/div/div[3]/div/div[1]/span/span/button')))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)
click_selenium(By.XPATH, '/html/body/div[4]/div/div/div/div/div/div/ul/li[4]/button/div/span')

pyautogui.sleep(5)
driver.quit()

# pyautogui.sleep(2)
# driver.back()
# pyautogui.sleep(2)

