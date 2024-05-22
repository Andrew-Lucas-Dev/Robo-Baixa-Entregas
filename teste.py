import pyautogui
import os

caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)

def click_image(image_path, confidence=0.8):
    # Construir o caminho completo para a imagem
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

def click_nota(image_path, confidence=0.8):
    # Construir o caminho completo para a imagem
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




for i in range(2):
    click_image('cancelar.png')
    click_image('digitar_data.png')
    for i in range(10):
        pyautogui.press("backspace")
    pyautogui.write("14/03/2024")
    click_nota('digitar_nota.png')
    pyautogui.write("1181331")
    click_image('atualizar.png')
    click_image('salvar_filial.png')
    pyautogui.write("5")
    click_image('salvar_ocorrencia.png')
    pyautogui.write("1")
    click_image('salvar_observ.png')
    pyautogui.write("1")
    click_image('salvar_datachegada.png')
    for i in range(10):
        pyautogui.press("backspace")
    pyautogui.write("18/03/202417:00")
    pyautogui.press('tab')
    pyautogui.write("18/03/202417:01")
    pyautogui.press('tab')
    pyautogui.write("18/03/202417:02")
    pyautogui.press('tab')
    pyautogui.write("aaa")
    for i in range(2):
        pyautogui.press("tab")
    pyautogui.write("111")
    #click_image('efetuar_baixa.png')
    pyautogui.sleep(1)
    click_image('cancelar.png')










