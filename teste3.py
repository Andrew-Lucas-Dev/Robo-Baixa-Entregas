from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import pandas as pd
import cv2
import numpy as np
from selenium import webdriver
from PIL import Image
from io import BytesIO
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.keys import Keys          
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

driver = webdriver.Chrome()
driver.get("https://doriprod.runteccorp.com/dori/logoff.php")
driver.maximize_window()

def click_selenium(selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")

click_selenium(By.XPATH, '//*[@id="dialog-cookies"]/div/div/div[3]/button')

def captura_salva_imagem(driver, x1, y1, x2, y2, output_file):
    pyautogui.sleep(3)  # Pausa de 3 segundos para garantir que a página esteja carregada

    # Capturar a tela da página atual
    screenshot = driver.get_screenshot_as_png()

    # Converter a screenshot em uma imagem Pillow
    image = Image.open(BytesIO(screenshot))

    # Recortar a imagem para a área específica
    cropped_image = image.crop((x1, y1, x2, y2))

    # Salvar a imagem recortada
    cropped_image.save(output_file)
    print(f"Imagem recortada salva como {output_file}")

def process_and_extract_text(image_path):
    # Especificar o caminho para o executável do Tesseract se não estiver no PATH do sistema
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    # Abrir a imagem usando Pillow
    image = Image.open(image_path)

    # Pré-processamento da imagem
    # Converter para escala de cinza
    gray_image = image.convert('L')

    # Aumentar o contraste
    enhancer = ImageEnhance.Contrast(gray_image)
    gray_image = enhancer.enhance(2)

    # Aplicar um filtro de nitidez
    sharp_image = gray_image.filter(ImageFilter.SHARPEN)

    # Converter a imagem para array numpy
    image_np = np.array(sharp_image)

    # Aplicar uma operação de fechamento para remover ruídos
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    closed_image = cv2.morphologyEx(image_np, cv2.MORPH_CLOSE, kernel)

    # Usar pytesseract para extrair texto
    custom_config = r'--oem 3 --psm 6'
    texto = pytesseract.image_to_string(closed_image, config=custom_config, lang='por')

    # Remover espaços e caracteres finais
    texto = texto.replace(" ", "").rstrip()

    return texto, len(texto)



i = 0
pyautogui.sleep(2)
while True:
    driver.refresh()
    driver.execute_script("window.scrollBy(0, -100);")
    try:
        print("Preenchendo campo de e-mail...")
        email_field = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'usuario')))
        email_field.click() 
        email_field.send_keys(Keys.CONTROL + "a")
        email_field.send_keys(Keys.BACKSPACE)
        email_field.send_keys('jeniffer.david@jettatransportes.com.br')
        print("Campo de e-mail preenchido com sucesso!")       
    except Exception as e:
        print("Erro ao preencher campo de e-mail:", e)
        pyautogui.sleep(2)
    try:
        print("Preenchendo campo de senha...")
        password_field = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'senha')))
        password_field.click() 
        password_field.send_keys(Keys.CONTROL + "a")
        password_field.send_keys(Keys.BACKSPACE)
        password_field.send_keys('@MUDARsenha01')
        print("Campo de senha preenchido com sucesso!")
    except Exception as e:
        print("Erro ao preencher campo de senha:", e)
    if i == 0:
        x1, y1, x2, y2 = 335, 360, 465, 410
    else:
        x1, y1, x2, y2 = 335, 495, 465, 535
    
    captura_salva_imagem(driver, x1, y1, x2, y2, 'recorte.png')
    image_path = 'recorte.png'
    texto, comprimento = process_and_extract_text(image_path)
    print(texto)
    if len(texto) == 4:
        print("A variável contém exatamente quatro caracteres.")
        try:
            print("Preenchendo campo da imagem...")
            password_field = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.ID, 'imagem')))
            password_field.send_keys(texto)
            print("Campo de senha preenchido com sucesso!")
        except Exception as e:
            print("Erro ao preencher campo de senha:", e)
    else:
        print("A variável não contém exatamente quatro caracteres.")
        continue

    pyautogui.sleep(1)
    driver.execute_script("window.scrollBy(0, 100);")
    pyautogui.sleep(1)
    click_selenium(By.XPATH, '/html/body/div[4]/div[2]/div[1]/form/button[2]')
    i += 1
    try:
        print("Clicando menu...")
        password_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/nav/button/span')))
        password_field.click()
        print("Menu clicado com sucesso!")
        break
    except Exception as e:
        print("Erro ao clicar no menu:", e)
        continue

print(texto)

pyautogui.sleep(5)

# Fechar o navegador
driver.quit()