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
import os

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

click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/div[1]/span/span[2]/button')
try:
    print("Pasta Planilha Bahia...")
    corpo_email = WebDriverWait(driver, 360).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/div[1]/span/span[2]/button')))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)
click_selenium(By.XPATH, '/html/body/div[3]/div/div/div/div/div/div/ul/li[4]/button/div/span')
pyautogui.sleep(2)
driver.back()
pyautogui.sleep(2)

click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[3]/div/div[1]/span/span[2]/button')
try:
    print("Pasta Planilha CC15...")
    corpo_email = WebDriverWait(driver, 360).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div/div/div/div[3]/div/div[1]/span/span[2]/button')))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)
click_selenium(By.XPATH, '/html/body/div[3]/div/div/div/div/div/div/ul/li[4]/button/div/span')
pyautogui.sleep(2)
driver.back()
pyautogui.sleep(2)

click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[3]/div/div/div[3]/div/div[1]/span/span[2]/button')
try:
    print("Pasta Planilha CC19...")
    corpo_email = WebDriverWait(driver, 360).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div/div/div/div[3]/div/div[1]/span/span[2]/button')))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)

click_selenium(By.XPATH, '/html/body/div[3]/div/div/div/div/div/div/ul/li[4]/button/div/span')
pyautogui.sleep(5)

driver.quit()

arquivos = [
    'EntregaT2.xlsx',
    'planilhaderotascc15.xlsx',
    'planilhaderotascc19.xlsx'
]

diretorio_origem = r'C:/Users/Usuario/Downloads/'
diretorio_destino = r'C:\Users\Usuario\Desktop\Robo-Baixa-Entregas'

if not os.path.exists(diretorio_destino):
    os.makedirs(diretorio_destino)

for arquivo in arquivos:
    caminho_origem = os.path.join(diretorio_origem, arquivo)
    caminho_destino = os.path.join(diretorio_destino, arquivo)
    try:
        if os.path.exists(caminho_origem):
            shutil.move(caminho_origem, caminho_destino)
            print(f"Arquivo '{arquivo}' movido com sucesso para '{diretorio_destino}'")
        else:
            print(f"O arquivo de origem '{caminho_origem}' não existe.")
    except Exception as e:
        print(f"Erro ao mover o arquivo '{arquivo}': {e}")