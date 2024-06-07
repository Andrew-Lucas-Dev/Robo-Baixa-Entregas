import pyautogui
import os

caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)


# Definindo a função para clicar na imagem
def click_image(image_name, confidence=0.9):
    # Construir o caminho completo para a imagem
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_name)
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
            print(f"Erro ao procurar a imagem: {e}")
        print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)


def verificar_campo(image_name, confidence=0.9):
    # Construir o caminho completo para a imagem
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
            # Se não encontrou após 5 tentativas, realiza as ações adicionais
            click_image('salvar_filial.png')
            pyautogui.press("backspace")
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            pyautogui.press("backspace")
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            pyautogui.press("backspace")
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")

# Exemplo de uso
verificar_campo('campo_filial.png')
verificar_campo('campo_observacao.png')
verificar_campo('campo_ocorrencia.png')







