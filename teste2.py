import pandas as pd
from openpyxl import load_workbook
import os
import numpy as np
import pyautogui
import ctypes

caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)


def click_image(image_path, confidence=0.9):
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

def imagem_encontrada(image_path, confidence=0.9, max_attempts=5):
    # Construir o caminho completo para a imagem
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
    
    return False
 
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

def click_nota(image_path, confidence=0.9):
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

def image_erro(image_name,image_name2,image_name3,image_name4,confidence=0.9):
    # Construir o caminho completo para a imagem
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_name)
    image_path2 = os.path.join(caminho_imagem, image_name2)
    image_path3 = os.path.join(caminho_imagem, image_name3)
    image_path4 = os.path.join(caminho_imagem, image_name4)
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
                print("Imagem de Baixa efetuada")
                break
        except Exception as e:
            print("Imagem de Baixa não encontrada na tela.")
        try:
            position4 = pyautogui.locateOnScreen(image_path4, confidence=confidence)
            if position4:
                print("Imagem de Baixa efetuada")
                break
        except Exception as e:
            print("Imagem de Baixa não encontrada na tela.")
        
        pyautogui.sleep(1)

def erro_efetuar(image_path, confidence=0.9, max_attempts=5):
    # Construir o caminho completo para a imagem
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, r'IMAGENS', image_path)
    attempts = 0
    while attempts < max_attempts:
        try:
            position = pyautogui.locateOnScreen(caminho_imagem, confidence=confidence)
            if position:
                click_image('OK.png')
                pyautogui.sleep(2)
                click_image('OK.png')
                click_image('efetuar_baixa.png')
                break
        except Exception as e:
            print(f"Imagem não encontrada na tela. Tentativa {attempts + 1} de {max_attempts}. Aguardando...")
        attempts += 1
        pyautogui.sleep(1)

def check_caps_lock():
    return ctypes.windll.user32.GetKeyState(0x14) & 0xffff != 0

def alt_press(key):
    pyautogui.keyDown('alt')
    pyautogui.press(key)
    pyautogui.keyUp('alt')

# Carregar os dados dos arquivos Excel
Planilha_1 = pd.read_excel("EntregaT2.xlsx", skiprows=4)
Planilha_1 = Planilha_1.rename(columns={'Nota Fiscal': 'NF','Data Prevista': 'DATA NOTA FISCAL'})
colunas_para_remover = ['Cobrou Descarga?', 'Motivo Atraso', 'Serie', 'Motivo Devolução', 'Cliente', 'Emp Carga', 'Bren', 'SP', 'Número Carga']
Planilha_1.drop(columns=colunas_para_remover, inplace=True)
Planilha_1 = Planilha_1.dropna(axis=1, how='all')
#print(Planilha_1)

Planilha_2 = pd.read_excel("BASE_DADOS.xlsx")
Planilha_2 = Planilha_2.dropna(axis=1, how='all')
#print(Planilha_2)

Planilha_sumare = pd.read_excel("planilha_sumare.xlsx")
colunas_para_remover = ['Série', 'Cnpj cliente', 'N° Carga','Status da Baixa']
Planilha_sumare.drop(columns=colunas_para_remover, inplace=True)
Planilha_sumare = Planilha_sumare.rename(columns={'N° NF': 'NF', 'Data NF': 'DATA NOTA FISCAL', 'DATA  ENTREGA': 'Data Chegada', 'Status da entrega': 'STATUS'})
Planilha_sumare['Data Entrega'] = Planilha_sumare['Data Chegada']
Planilha_sumare['Fim Descarreg.'] = Planilha_sumare['Data Chegada']
Planilha_sumare = Planilha_sumare.dropna(axis=1, how='all')
#print(Planilha_sumare)

# # Juntar os dados das duas planilhas
combined_df = pd.concat([Planilha_2, Planilha_1,Planilha_sumare], ignore_index=True)
combined_df = combined_df[combined_df['STATUS'] == 'Entregue']
combined_df = combined_df.drop_duplicates(subset='NF', keep='first')
combined_df['BAIXADO'] = combined_df['BAIXADO'].fillna('NAO')
#print(combined_df)

# # # #LOGIN
# # if check_caps_lock():
# #     pyautogui.press("capslock")  # Desativa o CAPS LOCK se estiver ativado
# # pyautogui.keyDown('win')
# # pyautogui.press("m")
# # pyautogui.keyUp('win')
# # click_image('logo_rodopar_areatrabalho.png')#PC ESCRITORIO
# # #click_image('logo_rodopar_areatrabalho_casa.png')#PC CASA
# # pyautogui.click()
# # click_image('conectar_rodopar.png')
# # click_image('senha_rodopar_1.png')
# # pyautogui.write("17@mudar")
# # click_image('ok_primeiro_login.png')
# # click_image('sim_primeiro_login.png')
# # click_image('segundo_login.png')
# # pyautogui.sleep(1)
# # pyautogui.write("anascimento")
# # pyautogui.press("tab")
# # pyautogui.write("99060767")
# # for i in range(2): 
# #     pyautogui.press("enter")
# # click_image('escolha_filial.png')
# # pyautogui.press("enter")

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
numero_linhas = len(combined_df)
#print(numero_linhas)

for i, linha in enumerate(combined_df.index):
    nf = combined_df.loc[linha, "NF"]
    data_nota_fiscal = combined_df.loc[linha, "DATA NOTA FISCAL"]
    data_chegada = combined_df.loc[linha, "Data Chegada"]   
    data_entrega = combined_df.loc[linha, "Data Entrega"]   
    data_fim_descarregamento =  combined_df.loc[linha, "Fim Descarreg."]  
    baixado = str(combined_df.loc[linha, "BAIXADO"])  
    #print(linha)
    if baixado == "SIM":
        continue
    else:    
        data_chegada = data_chegada.strftime('%d/%m/%Y')
        data_entrega = data_entrega.strftime('%d/%m/%Y')
        data_fim_descarregamento = data_fim_descarregamento.strftime('%d/%m/%Y')
        data_um_mes_antes = pd.to_datetime(data_entrega, format='%d/%m/%Y')
        data_um_mes_antes = data_um_mes_antes - pd.DateOffset(weeks=2)
        data_um_mes_antes = data_um_mes_antes.strftime('%d/%m/%Y')    
        status = 'ENTREGUE'
        combined_df.loc[linha, "BAIXADO"] = "SIM"
        falta = numero_linhas - i 
        print(f'nota:{nf} data da nota:{data_um_mes_antes} data chegada:{data_chegada} data entrega:{data_entrega} data fim descarregamento:{data_fim_descarregamento} falta:{falta}')
        
        click_image('cancelar.png')
        click_image('digitar_data.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(data_um_mes_antes))
        pyautogui.sleep(1)  
        click_nota('digitar_nota.png')
        pyautogui.sleep(1)
        pyautogui.write(str(nf))
        pyautogui.sleep(1)
        click_image('atualizar.png')
        pyautogui.sleep(1)
        if imagem_encontrada('nota_encontrada.png'):        
            click_image('salvar_filial.png')
            pyautogui.write("1")
            pyautogui.sleep(1)       
            click_image('salvar_ocorrencia.png')
            pyautogui.write("1")
            pyautogui.sleep(1)      
            click_image('salvar_observ.png')
            pyautogui.write("1")
            pyautogui.sleep(1)           
            pyautogui.press("tab")
            verificar_campo('campo_filial.png')
            verificar_campo('campo_observacao.png')
            verificar_campo('campo_ocorrencia.png')
            click_image('salvar_datachegada.png')
            for i in range(10):
                pyautogui.press("backspace")
            pyautogui.write(str(data_chegada))
            pyautogui.write("22:00")  ## 22 HORAS É UMA MARCAÇÃO ARBITRARIA DO NOSSO TIME
            pyautogui.sleep(1)
            pyautogui.press('tab')
            pyautogui.write(str(data_entrega))
            pyautogui.write("22:01")
            pyautogui.sleep(1)
            pyautogui.press('tab')
            pyautogui.write(str(data_fim_descarregamento))
            pyautogui.write("22:02")
            pyautogui.sleep(1)
            pyautogui.press('tab')
            pyautogui.write("aaa")
            pyautogui.sleep(1)
            for i in range(2):
                pyautogui.press("tab")
            pyautogui.write("111")
            pyautogui.sleep(1)
            click_image('efetuar_baixa.png')
            image_erro('mensagem_erro.png','mensagem_deu_certo.png','campo_filial.png','imagem_baixar_entrega.png')
            erro_efetuar('Erro_sistema_caiu.png')
            click_image('yes.png')
            pyautogui.sleep(2)
            click_image('yes.png')
            pyautogui.sleep(2)
            click_image('OK.png')
            pyautogui.sleep(2)


#print(combined_df)  
combined_df.to_excel('BASE_DADOS.xlsx', index=False)    
click_image('botao_voltar.png')