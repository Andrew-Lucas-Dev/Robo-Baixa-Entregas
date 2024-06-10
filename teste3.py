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
            print("Imagem não encontrada na tela. Aguardando...")
        
        pyautogui.sleep(1)

def image_erro(image_name,image_name2,image_name3, confidence=0.9):
    # Construir o caminho completo para a imagem
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_name)
    image_path2 = os.path.join(caminho_imagem, image_name2)
    image_path3 = os.path.join(caminho_imagem, image_name3)
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
                print("Baixa efetuada")
                break
        except Exception as e:
            print("Imagem não encontrada na tela.")
        
        pyautogui.sleep(1)


# def verificar_campo(image_name, confidence=0.9):
#     # Construir o caminho completo para a imagem
#     current_dir = os.path.dirname(__file__)  # Diretório atual do script
#     caminho_imagem = os.path.join(current_dir, 'IMAGENS')
#     image_path = os.path.join(caminho_imagem, image_name)    
#     while True:
#         found = False
#         for i in range(5): 
#             try:
#                 position = pyautogui.locateOnScreen(image_path, confidence=confidence)
#                 if position:
#                     print("Campo preenchido.")
#                     found = True
#                     break
#             except Exception as e:
#                 print(f"Erro ao procurar o campo: {e}")
#             print("Campo não preenchido. Aguardando...")
#             pyautogui.sleep(1)       
#         if found:
#             break
#         else:
#             # Se não encontrou após 5 tentativas, realiza as ações adicionais
#             click_image('salvar_filial.png')
#             pyautogui.press("backspace")
#             pyautogui.write("1")
#             pyautogui.sleep(1)   
#             pyautogui.press("tab")
#             pyautogui.press("backspace")
#             pyautogui.write("1")
#             pyautogui.sleep(1)   
#             pyautogui.press("tab")
#             pyautogui.press("backspace")
#             pyautogui.write("1")
#             pyautogui.sleep(1)   
#             pyautogui.press("tab")

# # Exemplo de uso
# verificar_campo('campo_filial.png')
# verificar_campo('campo_observacao.png')
# verificar_campo('campo_ocorrencia.png')

#image_erro('mensagem_erro.png','mensagem_deu_certo.png','campo_filial.png')

def imagem_encontrada(image_path, confidence=0.9, max_attempts=5):
    # Construir o caminho completo para a imagem
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_path)
    
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem foi encontrada na tela.")
                return True
        except Exception as e:
            print(f"Erro ao procurar a imagem: {e}")
        print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)
    
    return False

# Exemplo de uso
if imagem_encontrada('nota_encontrada.png'):
    print("Imagem foi encontrada.")
else:
    print("Imagem não foi encontrada.")







