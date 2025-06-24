import sqlite3
import pandas as pd
import re
from datetime import datetime
import pyautogui
import os
import shutil
import subprocess
import ctypes
from selenium import webdriver
from selenium.webdriver.common.keys import Keys          
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

time = 0.5
caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)
app_path = r"C:\Users\Usuario\Desktop\VR_PROD_JETTA.rdp"
multi = False

def processar_notas_em_lote(lista_notas):
    """
    Processa um lote de notas fiscais, inserindo ou atualizando conforme necessário.
    
    :param lista_notas: Lista de dicionários contendo as informações das notas fiscais.
    """
    with sqlite3.connect('banco_dados_entregas.db') as conexao:
        cursor = conexao.cursor()

        for nota in lista_notas:
            try:
                filial = nota['filial']
                serie = nota['serie']
                numero_nota = nota['nota']
                data_nota = nota['data_nota']
                data_chegada = nota['data_chegada']
                data_entrega = nota['data_entrega']
                data_descarreg = nota['data_descarreg']
                status = nota['status']
                man = nota['man']
                num_carga = nota['num_carga']
                num_carga = str(num_carga)
                baixado = nota['baixado']

                # Verifica se a combinação nota + MDFe já existe
                cursor.execute('SELECT id FROM notas WHERE nota = ? AND MDFe = ?', (numero_nota, man))
                resultado = cursor.fetchone()

                if resultado:
                    print(f"A nota {numero_nota} já está associada ao MDFe {man}. Registrando novo histórico.")

                # Tenta inserir uma nova linha no banco de dados
                cursor.execute(''' 
                INSERT INTO notas (filial,serie, nota, data_nota_fiscal, data_chegada, data_entrega, data_descarreg, status, baixado, MDFe, num_carga)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'NAO', ?, ?)''', 
                (filial,serie, numero_nota, data_nota, data_chegada, data_entrega, data_descarreg, status, man, num_carga))
                
                print(f"Nota {numero_nota} processada com sucesso no MDFe {man}.")

            except sqlite3.IntegrityError:
                print(f"Nota {numero_nota} ja inserida, verificando novo status.")
                # Atualiza o registro existente caso o status seja 'Entregue'
                if status == 'Entregue' and baixado == 'NAO':
                    cursor.execute(''' 
                    UPDATE notas
                    SET 
                        filial = ?, 
                        serie = ?,
                        data_nota_fiscal = ?, 
                        data_chegada = ?, 
                        data_entrega = ?, 
                        data_descarreg = ?, 
                        status = ?, 
                        baixado = '?', 
                        MDFe = ?, 
                        num_carga = ?
                    WHERE nota = ? AND MDFe = ?''', 
                    (filial,serie, data_nota, data_chegada, data_entrega, data_descarreg, status,baixado, man, num_carga, numero_nota, man))
                    
                    print(f"Nota {numero_nota} atualizada para status 'Entregue' no MDFe {man}.")
                else:
                    print(f"Nota {numero_nota} ja inserida e nao atualizada.")
                
        conexao.commit()


def visualizar_tabela():
    conexao = sqlite3.connect('banco_dados_entregas.db')
    cursor = conexao.cursor()

    cursor.execute('SELECT * FROM notas')
    notas = cursor.fetchall()
    for nota in notas:
        print(f'filial:{nota[1]} nota:{nota[2]} man:{nota[9]} status:{nota[7]} baixado:{nota[8]}')
    conexao.close()
    return

def listar_notas():
    conexao = sqlite3.connect('banco_dados_entregas.db')
    cursor = conexao.cursor()

    # Selecionar todas as notas não baixadas e ordenar por MDFe e filial
    cursor.execute("""
        SELECT * 
        FROM notas 
        WHERE baixado = 'NAO' 
        ORDER BY MDFe ASC, filial ASC
    """)
    notas = cursor.fetchall()
    conexao.close()
    return notas

def atualizar_status(nota, baixado, manifesto, filial):
    conexao = sqlite3.connect('banco_dados_entregas.db')
    cursor = conexao.cursor()

    cursor.execute('''
    UPDATE notas
    SET baixado = ?
    WHERE nota = ? AND MDFe = ? AND filial = ?
    ''', (baixado, nota, manifesto, filial))
    
    conexao.commit()
    print(f"Status da nota {nota} atualizado para {baixado}.")
    conexao.close()

def atualizar_status_man(filial, man, baixado):
    conexao = sqlite3.connect('banco_dados_entregas.db')
    cursor = conexao.cursor()

    # Atualizar todas as notas com o mesmo MDFe e filial
    cursor.execute('''
        UPDATE notas
        SET baixado = ?
        WHERE MDFe = ? AND filial = ?
    ''', (baixado, man, filial))
    
    conexao.commit()
    print(f"Status de todas as notas do manifesto {man} na filial {filial} atualizado para {baixado}.")
    conexao.close()

def visualizar_tabela_man(filial, man):
    conexao = sqlite3.connect('banco_dados_entregas.db')
    cursor = conexao.cursor()

    cursor.execute('SELECT * FROM notas')

    cursor.execute('''
        SELECT * FROM notas
        WHERE MDFe = ? AND filial = ?
    ''', (man, filial))

    notas = cursor.fetchall()
    for nota in notas:
        print(f'filial:{nota[1]} nota:{nota[2]} man:{nota[9]} status:{nota[7]} baixado:{nota[8]}')
    conexao.close()
    return

def excluir_nota(nota):
    conexao = sqlite3.connect('banco_dados_entregas.db')
    cursor = conexao.cursor()

    cursor.execute('''
    DELETE FROM notas
    WHERE nota = ?
    ''', (nota,))
    
    conexao.commit()
    print(f"Nota {nota} excluída com sucesso.")
    conexao.close()



def check_caps_lock():
    return ctypes.windll.user32.GetKeyState(0x14) & 0xffff != 0

def alt_press(key):
    pyautogui.keyDown('alt')
    pyautogui.press(key)
    pyautogui.keyUp('alt')

def login():
    try:
        # Verifica se o arquivo existe antes de abrir
        if os.path.exists(app_path):
            subprocess.Popen(["mstsc", app_path])  # mstsc é o cliente RDP no Windows
        else:
            print(f"O arquivo {app_path} não foi encontrado.")
    except Exception as e:   
        print(f"Ocorreu um erro: {e}")
    
    #LOGIN
    if check_caps_lock():                                                                                                                                                                                                      
        pyautogui.press("capslock")
    click_image('conectar_rodopar.png')
    #click_image('conectar_rodopar1.png')
    click_image('senha_rodopar_1.png')
    #click_image('senha_rodopar_2.png')
    pyautogui.write("10mudar")
    click_image('ok_primeiro_login.png')
    #click_image('ok_primeiro_login2.png')
    click_image('sim_primeiro_login.png')
    # pyautogui.sleep(15)
    # pyautogui.press('enter')
    # pyautogui.sleep(10)
    # pyautogui.press('enter')
    # pyautogui.sleep(5)
    # pyautogui.press('enter')
    #   click_image('sim_primeiro_login2.png')
    click_image('segundo_login.png')            
    pyautogui.sleep(1)
    pyautogui.write("llima")
    pyautogui.press("tab")  
    pyautogui.write("1234")
    for i in range(2): 
        pyautogui.press("enter")
    click_image('filial_1.png')
    pyautogui.press("enter")
    click_image('botao_faturamento.png')
    click_image('botao_faturamento_movimentacao.png')
    click_image('botao_faturamento_movimentacao_entregas.png')
    click_image('cancelar.png')

def click_selenium(selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")

def click_(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    try:
        position = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if position:
            center_x = position.left + position.width // 2
            center_y = position.top + position.height // 2
            pyautogui.click(center_x, center_y)
            print("Imagem foi encontrada na tela.")
    except Exception as e:
        print("Imagem não encontrada na tela. Aguardando...")

def click_image(image_path, confidence=0.9):
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

def na_tela(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    for i in range(3):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Nenhuma nota selecionada.")
                pyautogui.sleep(2)
                pyautogui.press('enter')
                return True
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")

            pyautogui.sleep(1)
    return False

def click_nota(image_path, confidence=0.9):
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

def confirmacao_preenchido(image_path, image_path2, image_path3, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    a, b, c = 0, 0, 0
    while True:
        click_('ok.png')
        click_('ok_marcado.png')
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem 1 foi encontrada na tela.")
                a = 0  # Redefine 'a' se a imagem 1 for encontrada
                click_image('referencia.png')
                for i in range(10):
                    pyautogui.press("backspace")
                pyautogui.write(str(referencia))
            else:
                a = 1  # Define 'a' como 1 se a imagem 1 não for encontrada
        except Exception as e:
            print("Erro ao procurar imagem 1:", e)
            a = 1
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                print("Imagem 2 foi encontrada na tela.")
                b = 0  # Redefine 'b' se a imagem 2 for encontrada
                click_image('digitar_data.png')
                for i in range(10):
                    pyautogui.press("backspace")
                pyautogui.write(str(data_nota_fiscal))
                pyautogui.sleep(time)  
            else:
                b = 1  # Define 'b' como 1 se a imagem 2 não for encontrada
        except Exception as e:
            print("Erro ao procurar imagem 2:", e)
            b = 1
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position3:
                print("Imagem 3 foi encontrada na tela.")
                c = 0  # Redefine 'c' se a imagem 3 for encontrada
                click_nota('digitar_nota.png')
                pyautogui.sleep(time)
                if multi == True:
                    pyautogui.write('00')
                    pyautogui.sleep(0.2)
                pyautogui.write(str(nf))
                pyautogui.sleep(time)
            else:
                c = 1  # Define 'c' como 1 se a imagem 3 não for encontrada
        except Exception as e:
            print("Erro ao procurar imagem 3:", e)
            c = 1
        # Se todas as três imagens não forem encontradas, sai do loop
        if a == 1 and b == 1 and c == 1:
            print("Nenhuma das três imagens foi encontrada. Saindo do loop.")
            click_image('atualizar.png')
            pyautogui.sleep(time)
            for i in range(2):
                pyautogui.press('tab')
            break

def verificar_campo(filial,image_name, confidence=0.9):
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
            click_image('salvar_filial.png')
            for i in range(5):
                pyautogui.press("backspace")
            pyautogui.write(str(filial))
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            for i in range(5):
                pyautogui.press("backspace")
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            for i in range(5):
                pyautogui.press("backspace")            
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")

def click_filial_man(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                center_x += 50  # Adicionar deslocamento para a direita
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def click_serie_man(image_path, confidence=0.9):
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

def click_man(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                center_x += 130  # Adicionar deslocamento para a direita
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def manifesto_encontrada(image_path,image_path2,image_path3, confidence=0.9, max_attempts=7):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_path)
    image_path2 = os.path.join(caminho_imagem, image_path2)
    image_path3 = os.path.join(caminho_imagem, image_path3)
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position or position2:
                print("Manifesto nao encontrado na tela. Aguardando...")
        except Exception as e:
            print("Manifesto encontrada na tela.")  
            return True
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position3:
                print("Manifesto encontrada na tela.")
                return True
        except Exception as e:
            print("Manifesto não encontrada na tela. Aguardando...")

        pyautogui.sleep(1)
    return False    

def imagem_encontrada(image_path, confidence=0.9, max_attempts=7):
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
    pyautogui.sleep(time) 
    # for i in range(2):
    #     pyautogui.press("down")
    #     pyautogui.sleep(0.5)
    # pyautogui.press("enter") 
    click_image('OST.png')  
    pyautogui.sleep(time)   
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
    pyautogui.sleep(0.5)
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
    pyautogui.sleep(1)
    click_image('atualizar.png')
    pyautogui.sleep(1)
    for i in range(2):
        pyautogui.press('tab')
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

def finalizar_baixa(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Pronto para proxima baixa.")
                break
        except Exception as e:
            click_('yes.png')
            click_('yes_marcado.png')
            click_('ok.png')
            click_('yes2.png')
            click_('sim_marcado.png')
            click_('sim_confirmar_entrega.png')
            click_('sim2.png')
            click_('cancelar.png')
            

def remover_duplicados_por_filial(manifestos):
    vistos = set()  # Armazena as combinações (Filial, Manifesto) já vistas
    resultado = []  # Lista filtrada
    for manifesto in manifestos:
        chave = (manifesto['Filial'], manifesto['Manifesto'])  # Combinação única de Filial e Manifesto
        if chave not in vistos:  # Verifica se a combinação (Filial, Manifesto) já foi processada
            resultado.append(manifesto)  # Adiciona o manifesto se for único
            vistos.add(chave)  # Marca a combinação como vista
    return resultado

def processar_coluna_data(df, coluna):
    df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
    df[coluna] = df[coluna].dt.date
    df[coluna] = df[coluna].astype(str).str.strip()  # Converter para string e remover espaços em branco
    df[coluna] = pd.to_datetime(df[coluna], format='%Y-%m-%d', errors='coerce')  # Converter para datetime
    df[coluna] = df[coluna] + pd.Timedelta(hours=22)  # Adicionar o horário '22:00'
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y %H:%M')  # Formatar no formato desejado

def processar_coluna_chegada(df, coluna):
    df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
    df[coluna] = df[coluna] + pd.Timedelta(hours=22)
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')

def processar_datas(df, colunas):
    for coluna in colunas:
        if 'Data Entrega' in coluna:
            df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=1)
        elif 'Fim Descarreg.' in coluna:
            df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=2)
        df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')
        
def formatar_datas(df, colunas):
    for coluna in colunas:
        df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
        df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')

# Função para extrair o número completo do MDF-e e formatá-lo
def extrair_mdf_e(mdf_e):
    if pd.isna(mdf_e):
        return None
    match = re.search(r"(?:1/2|2/1|2/2|/5/1|5/2)/([\d.]+)", str(mdf_e))
    return match.group(1).replace(".", "") if match else None

# Função para extrair o número da filial (antes da barra)
def extrair_filial(mdf_e):
    if pd.isna(mdf_e):
        return None
    match = re.search(r"MDF-e: (\d+)/", str(mdf_e))
    return int(match.group(1)) if match else None

# Função para extrair o número da filial (antes da barra)
def extrair_serie(mdf_e):
    if pd.isna(mdf_e):  # Verifica se é NaN
        return None
    match = re.search(r"MDF-e: \d+/(\d+)/", str(mdf_e))
    return int(match.group(1)) if match else None

# Função comum para processar a planilha
def processar_planilha(nome_arquivo, colunas_para_remover, coluna_data, coluna_mdf_filial):
    Planilha = pd.read_excel(nome_arquivo)
    Planilha.drop(columns=colunas_para_remover, inplace=True)

    Planilha["MDF-e"] = Planilha[coluna_mdf_filial].where(
        Planilha[coluna_mdf_filial].astype(str).str.contains("MDF-e:", na=False)
    ).fillna(method="ffill")
    
    Planilha["Filial"] = Planilha["MDF-e"]
    Planilha["MDF-e"] = Planilha["MDF-e"].apply(extrair_mdf_e)
    Planilha["Serie"] = Planilha["Filial"].apply(extrair_serie)
    Planilha["Filial"] = Planilha["Filial"].apply(extrair_filial)

    Planilha['N° NF'] = pd.to_numeric(Planilha['N° NF'], errors='coerce')
    Planilha.dropna(subset=['N° NF'], inplace=True)
    Planilha['N° NF'] = Planilha['N° NF'].astype(int)

    Planilha = Planilha.rename(columns={'N° NF': 'NF', 'Data NF': 'DATA NOTA FISCAL', 'Data': 'Data Chegada', 'DATA': 'Data Chegada', 'Status da entrega': 'STATUS'})
    try:
        Planilha = Planilha.rename(columns={'Plano': 'N° Carga'})
    except:
        pass

    Planilha['Data Entrega'] = Planilha['Data Chegada']
    Planilha['Fim Descarreg.'] = Planilha['Data Chegada']

    Planilha['DATA NOTA FISCAL'] = pd.to_datetime(Planilha['DATA NOTA FISCAL']).dt.strftime('%d/%m/%Y')

    colunas_de_data = ['Data Entrega', 'Fim Descarreg.']
    for coluna in colunas_de_data:
        processar_coluna_data(Planilha, coluna)
    processar_datas(Planilha, colunas_de_data)
    processar_coluna_chegada(Planilha, 'Data Chegada')

    Planilha = Planilha.dropna(axis=1, how='all')
    
    if 'STATUS' in Planilha.columns:
        Planilha.loc[Planilha['STATUS'].str.upper().isin(['FINALIZADO', 'FINALIZADA']), 'STATUS'] = 'Entregue'

    # Tentar remover a coluna 'Plano' se existir
    # if 'Plano' in Planilha.columns:
    #     Planilha.drop(columns='Plano', inplace=True)


    return Planilha

driver = webdriver.Chrome()
driver.get("https://jettatransporte-my.sharepoint.com/:f:/g/personal/jetta_bi_jettatransporte_onmicrosoft_com/EiA6eCcrmHVOi0SVjgVS4eYBTgW6NmdHNlvRSINLlAOW5g?e=qyl9wK")
driver.maximize_window()        
pyautogui.sleep(2)

#click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/div[1]/span/span/button')
click_selenium(By.XPATH, '//*[@id="virtualized-list_3_page-0_579067"]/div[3]/span')
pyautogui.sleep(2)

# click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div/div[2]/div/div/div/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[3]/div/div/div[3]/div/div[1]/span/span/button')
                      
# try:
#     print("Pasta Planilha Bahia...")
#     corpo_email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div/div[2]/div/div/div/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div/div/div/div[3]/div/div[1]/span/span/button')))
#     action_chains = ActionChains(driver)                                                
#     action_chains.context_click(corpo_email).perform()
# except Exception as e:
#     print("Erro ao clicar no corpo do e-mail:", e)
# pyautogui.sleep(5)
# click_selenium(By.XPATH, "//span[text()='Baixar']")
# pyautogui.sleep(2)
# driver.back()
# pyautogui.sleep(2)

#click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[3]/div/div[1]/span/span/button')
click_selenium(By.XPATH, "//span[text()='Planilha CC15']")
pyautogui.sleep(5)
try:
    print("Pasta Planilha CC15...")
    corpo_email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='planilhaderotascc15.xlsx']")))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)

pyautogui.sleep(5)
click_selenium(By.XPATH, "//span[text()='Baixar']")
pyautogui.sleep(2)
driver.back()
pyautogui.sleep(2)

click_selenium(By.XPATH, "//span[text()='Planilha CC19']")
pyautogui.sleep(5)
try:
    print("Pasta Planilha CC19...")
    corpo_email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='planilhaderotascc19.xlsx']")))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)

pyautogui.sleep(5)
click_selenium(By.XPATH, "//span[text()='Baixar']")





pyautogui.sleep(2)
driver.back()
pyautogui.sleep(2)

click_selenium(By.XPATH, "//span[text()='Planilha CC16']")
pyautogui.sleep(5)
try:
    print("Pasta Planilha CC16...")
    corpo_email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='PLANILHA DE ROTAS CC16.xlsx']")))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)

pyautogui.sleep(5)
click_selenium(By.XPATH, "//span[text()='Baixar']")


pyautogui.sleep(2)
driver.back()
pyautogui.sleep(2)

click_selenium(By.XPATH, "//span[text()='Planilha CC21']")
pyautogui.sleep(5)
try:
    print("Pasta Planilha CC21...")
    corpo_email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Planilha de Baixas CC21.xlsx']")))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)

pyautogui.sleep(5)
click_selenium(By.XPATH, "//span[text()='Baixar']")

pyautogui.sleep(40)

arquivos = [
    'EntregaT2.xlsx',
    'planilhaderotascc15.xlsx',
    'planilhaderotascc19.xlsx',
    'PLANILHA DE ROTAS CC16.xlsx',
    'Planilha de Baixas CC21.xlsx'
]

diretorio_origem = r'C:/Users/Usuario/Downloads/'
diretorio_destino = r'C:\Users\Usuario\Desktop\Robo-Baixa-Entregas'
# diretorio_origem = r'C:/Users/Andrew/Downloads/'
# diretorio_destino = r'C:/Users/Andrew/Desktop/Robo-Baixa-Entregas'

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

driver.quit()
pyautogui.sleep(5)















# Parâmetros de colunas a remover
colunas_para_remover_cc19 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Vlr Merc.'#,
    #'Entrega Canhoto Físico'
]
colunas_para_remover_cc15 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Vlr Merc.',
    'Entrega Canhoto Físico', 'Status da baixa'
]
colunas_para_remover_cc16 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Manifesto', 'EM BRANCO', 'Bairro'
]
colunas_para_remover_cc21 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Vlr Merc.', 'Entrega Canhoto Físico', 'Manifesto'
]

# # Processar cada planilha
Planilha_CC19 = processar_planilha("planilhaderotascc19.xlsx", colunas_para_remover_cc19, 'N° Carga', 'N° Carga')
Planilha_CC15 = processar_planilha("planilhaderotascc15.xlsx", colunas_para_remover_cc15, 'N° Carga', 'N° Carga')
Planilha_CC16 = processar_planilha("PLANILHA DE ROTAS CC16.xlsx", colunas_para_remover_cc16, 'Plano', 'Plano')
Planilha_CC21 = processar_planilha("Planilha de Baixas CC21.xlsx", colunas_para_remover_cc21, 'N° Carga', 'N° Carga')
# print(Planilha_CC16)

# # # # Juntar as 4 planilhas
combined_df = pd.concat([Planilha_CC19,Planilha_CC15,Planilha_CC16,Planilha_CC21], ignore_index=True)

def ajustar_data(data_str):
    """
    Corrige a data caso o mês e o dia estejam invertid1    :param data_str: Data em formato string (DD/MM/YYYY)
    :return: Data corrigida em formato string (DD/MM/YYYY) ou a original se válida
    """
    if not data_str:
        return None  # Retorna None se a data estiver ausente
    
    try:
        # Tenta interpretar a data no formato correto
        data = datetime.strptime(data_str, '%d/%m/%Y')
        data_atual = datetime.now()

        # Verifica se a data parece futura (erro de troca de dia e mês)
        if data > data_atual:
            # Tenta corrigir interpretando como mês/dia/ano
            data_corrigida = datetime.strptime(data_str, '%m/%d/%Y')
            return data_corrigida.strftime('%d/%m/%Y')
        return data_str
    except ValueError:
        return data_str  # Retorna a data original se não puder ser interpretada

notas_para_processar = []

for index, row in combined_df.iterrows():
    try:
        # Extrair valores das colunas e ajustar as datas
        nota_dict = {
            'filial': row.get('Filial', 'Desconhecida'),
            'serie': row.get('Serie', 'Desconhecida'),
            'nota': row.get('NF'),
            'data_nota': ajustar_data(row.get('DATA NOTA FISCAL')),
            'data_chegada': ajustar_data(row.get('Data Chegada')),
            'data_entrega': ajustar_data(row.get('Data Entrega')),
            'data_descarreg': ajustar_data(row.get('Fim Descarreg.')),
            'status': row.get('STATUS', 'Pendente'),
            'man': row.get('MDF-e'),
            'num_carga': row.get('N° Carga', None),
            'baixado': row.get('baixado', None),
        }
        # Validação básica
        if not nota_dict['nota'] or not nota_dict['man']:
            raise ValueError(f"Nota ou MDFe ausente na linha {index}.")

        # Adicionar à lista de notas para processamento
        notas_para_processar.append(nota_dict)

    except Exception as e:
        print(f"Erro ao processar a linha {index}: {e}")
        continue

# Passar a lista de notas para a função de processamento em lote
if notas_para_processar:
    processar_notas_em_lote(notas_para_processar)
    print(f"{len(notas_para_processar)} notas processadas com sucesso.")
else:
    print("Nenhuma nota válida encontrada para processar.")

print('--------------------------------------------------------------')

login()

notas = listar_notas()

# Criar DataFrame
df = pd.DataFrame(notas, columns=['id', 'filial', 'serie', 'nota', 'data_nota_fiscal', 'data_chegada', 'data_entrega', 'data_descarreg', 'status', 'baixado', 'MDFe', 'num_carga'])

resultado = []  # Lista para salvar os manifestos ou status

# Agrupar por MDFe e Filial
grupos = df.groupby(["MDFe", "filial"])
print('Iniciando baixas ')

# Processar cada grupo
for (mdfe, filial), grupo in grupos:

    # Verificar se todos os status no grupo são "Entregue" ou "entregue"
    if grupo["status"].isin(["Entregue", "entregue", "ENTREGUE"]).all():

        # Usar valores representativos do grupo (primeira linha)
        primeira_linha = grupo.iloc[0]
        resultado.append({
            "Manifesto": mdfe,
            "Filial": filial,
            "serie": primeira_linha["serie"],
            "data_nota_fiscal": primeira_linha["data_nota_fiscal"],
            "data_chegada": primeira_linha["data_chegada"],
            "data_entrega": primeira_linha["data_entrega"],
            "data_descarreg": primeira_linha["data_descarreg"]
        })
    else:
        # Adicionar "NAO BAIXAR" com Filial
        resultado.append({
            "Manifesto": "NAO BAIXAR",
            "Filial": filial
        })
        # # Baixar notas individuais
        # print(f"Baixando notas do manifesto {mdfe} (Filial {filial})")
        # print(grupo)
        for _, linha in grupo.iterrows():
            if linha["status"] in ["Entregue", "entregue"]:
                nf = linha["nota"]
                manifesto = linha["MDFe"]
                data_nota_fiscal = linha["data_nota_fiscal"]
                data_chegada = linha["data_chegada"]
                data_entrega = linha["data_entrega"]
                data_descarregamento = linha["data_descarreg"]
                referencia = linha["num_carga"]
                baixado1 = linha["baixado"]
                nf = int(nf)
                filial = int(filial)
                if referencia is None:
                    continue

                # Convertendo a string para um objeto datetime
                data_referencia = datetime.strptime("15/12/2024", "%d/%m/%Y")
                data = datetime.strptime(data_nota_fiscal, "%d/%m/%Y")
                
                # Comparação
                if data_referencia > data:
                    print("A data é maior que 15/12/2024.")
                    continue
                
                print(f"Filial: {filial} Nota: {nf} man:{manifesto} carga:{referencia} data nota:{data_nota_fiscal} cheg:{data_chegada} entre:{data_entrega} desc:{data_descarregamento}")



                click_('ok_marcado.png')
                click_image('cancelar.png')
                click_('ok_marcado.png')
                pyautogui.sleep(time)
                click_image('cancelar.png')
                pyautogui.sleep(time)
                click_image('referencia.png')
                for i in range(10):
                    pyautogui.press("backspace")
                pyautogui.write(str(referencia))
                click_image('digitar_data.png') 
                for i in range(10):
                    pyautogui.press("backspace")
                pyautogui.write(str(data_nota_fiscal))
                pyautogui.sleep(time)  
                click_nota('digitar_nota.png')
                pyautogui.sleep(time)
                pyautogui.write(str(nf))
                pyautogui.sleep(time)
                click_image('atualizar.png')
                pyautogui.sleep(time)
                for i in range(2):
                    pyautogui.press('tab')
                if imagem_encontrada('nota_encontrada.png'):
                    confirmacao_preenchido('referencia.png','digitar_data.png','digitar_nota.png')
                    pyautogui.sleep(1)     
                    click_image('salvar_filial.png')
                    pyautogui.write(str(filial))
                    pyautogui.sleep(time)       
                    click_image('salvar_ocorrencia.png')
                    pyautogui.write("1")
                    pyautogui.sleep(time)      
                    click_image('salvar_observ.png')
                    pyautogui.write("1")
                    pyautogui.sleep(time)           
                    pyautogui.press("tab")

                    #COLOCAR IFS DE ACORDO COM O NUMERO DA FILIAL PRA PEGAR A FOTO CORRESPONDENTE. EX: IF FILIAL == 1
                    if filial == 1:
                        verificar_campo(filial,'campo_filial.png')
                    elif filial == 2:
                        verificar_campo(filial,'campo_filial_2.png')
                    elif filial == 3:
                        verificar_campo(filial,'.png')
                    elif filial == 4:
                        verificar_campo(filial,'.png')
                    elif filial == 5:
                        verificar_campo(filial,'.png')

                    verificar_campo(filial,'campo_observacao.png')
                    verificar_campo(filial,'campo_ocorrencia.png')
                    click_image('salvar_datachegada.png')
                    for i in range(10):
                        pyautogui.press("backspace")
                    pyautogui.write(str(data_chegada))
                    pyautogui.sleep(time)
                    pyautogui.press('tab')
                    pyautogui.write(str(data_entrega))
                    pyautogui.sleep(time)
                    pyautogui.press('tab')
                    pyautogui.write(str(data_descarregamento))
                    pyautogui.sleep(time)
                    pyautogui.press('tab')
                    pyautogui.write("aaa")
                    pyautogui.sleep(time)
                    for i in range(2):
                        pyautogui.press("tab")
                    pyautogui.write("111")
                    pyautogui.sleep(time)
                    click_image('efetuar_baixa.png')
                    pyautogui.sleep(2)
                    print('Verificando se teve nota selecionada')
                    if na_tela('nenhuma_nota_selecionada.png'):
                        continue                   
                    finalizar_baixa('digitar_data.png')
                    click_image('cancelar.png')
                    baixado = 'SIM'
                    atualizar_status(nf, baixado,manifesto,filial)
                else:
                    click_image('cancelar.png')
                


print('------------------------------------')
print('Baixando Manifestos')

manifestos_filtrados = remover_duplicados_por_filial(resultado)

# Exibir resultados
for item in manifestos_filtrados:
    filial = item["Filial"]
    man = item["Manifesto"]
    #baixar por manifesto
    if man != "NAO BAIXAR":
        serie = item["serie"]
        data_nota_fiscal = item["data_nota_fiscal"]
        data_chegada = item["data_chegada"]
        data_entrega = item["data_entrega"]
        data_descarregamento = item["data_descarreg"]
        man = int(man)
        filial = int(filial)

        # Convertendo a string para um objeto datetime
        data_referencia = datetime.strptime("15/12/2024", "%d/%m/%Y")
        data = datetime.strptime(data_nota_fiscal, "%d/%m/%Y")
        # if man == '21362':
        #     print(man)
        # Comparação
        if data_referencia > data:
            print("A data é maior que 15/12/2024.")
            # print(man)
            continue
        print(f"Filial: {filial} serie:{serie} Manifesto: {man} data nota:{data_nota_fiscal} cheg:{data_chegada} entre:{data_entrega} desc:{data_descarregamento}")
       

        pyautogui.sleep(time)
        click_image('cancelar.png')
        pyautogui.sleep(time)
        click_filial_man('digitar_man.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(filial))
        # pyautogui.press('tab')
        pyautogui.sleep(time)
        click_serie_man('digitar_man.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(serie))
        # pyautogui.press('tab')
        pyautogui.sleep(time)
        click_man('digitar_man.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(man))
        # pyautogui.press('tab')
        pyautogui.sleep(time)
        click_image('atualizar.png')
        pyautogui.sleep(3)

        #CTES
        if manifesto_encontrada('grid_limpa.png','grid_limpa2.png','encontrou_man.png'):
            print('achou cte')
            pyautogui.sleep(1)     
            click_image('salvar_filial.png')
            pyautogui.write(str(filial))
            pyautogui.sleep(time)       
            click_image('salvar_ocorrencia.png')
            pyautogui.write("1")
            pyautogui.sleep(time)      
            click_image('salvar_observ.png')
            pyautogui.write("1")
            pyautogui.sleep(time)           
            pyautogui.press("tab")

            #COLOCAR IFS DE ACORDO COM O NUMERO DA FILIAL PRA PEGAR A FOTO CORRESPONDENTE. EX: IF FILIAL == 1
            if filial == 1:
                verificar_campo(filial,'campo_filial.png')
            elif filial == 2:
                verificar_campo(filial,'campo_filial_2.png')
            elif filial == 3:
                verificar_campo(filial,'.png')
            elif filial == 4:
                verificar_campo(filial,'.png')
            elif filial == 5:
                verificar_campo(filial,'.png')

            verificar_campo(filial,'campo_observacao.png')
            verificar_campo(filial,'campo_ocorrencia.png')
            click_image('salvar_datachegada.png')
            for i in range(10):
                pyautogui.press("backspace")
            pyautogui.write(str(data_chegada))
            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.write(str(data_entrega))
            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.write(str(data_descarregamento))
            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.write("aaa")
            pyautogui.sleep(time)
            for i in range(2):
                pyautogui.press("tab")
            pyautogui.write("111")
            pyautogui.sleep(time)
            click_image('efetuar_baixa.png')
            pyautogui.sleep(2)
            print('Verificando se teve nota selecionada')
            if na_tela('nenhuma_nota_selecionada.png'):
                continue        
            finalizar_baixa('em_branco.png')
            click_image('cancelar.png')
            baixado = 'SIM'
            atualizar_status_man(filial, man, baixado)
        pyautogui.sleep(2)


        click_image('cancelar.png')
        pyautogui.sleep(time)
        click_filial_man('digitar_man.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(filial))
        # pyautogui.press('tab')
        pyautogui.sleep(time)
        click_serie_man('digitar_man.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(serie))
        # pyautogui.press('tab')
        pyautogui.sleep(time)
        click_man('digitar_man.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(man))
        # pyautogui.press('tab')

        pyautogui.sleep(time)

        #OST
        click_image('CTRC.png')
        pyautogui.sleep(time) 
        # for i in range(2):
        #     pyautogui.press("down")
        #     pyautogui.sleep(0.5)
        # pyautogui.press("enter")
        click_image('OST.png')
        pyautogui.sleep(time) 
        # click_image('atualizar.png')
        pyautogui.sleep(3) 
        if manifesto_encontrada('grid_limpa.png','grid_limpa2.png','encontrou_man.png'):   
            print('achou ost')

            pyautogui.sleep(1)     
            click_image('salvar_filial.png')
            pyautogui.write(str(filial))
            pyautogui.sleep(time)       
            click_image('salvar_ocorrencia.png')
            pyautogui.write("1")
            pyautogui.sleep(time)      
            click_image('salvar_observ.png')
            pyautogui.write("1")
            pyautogui.sleep(time)           
            pyautogui.press("tab")

            #COLOCAR IFS DE ACORDO COM O NUMERO DA FILIAL PRA PEGAR A FOTO CORRESPONDENTE. EX: IF FILIAL == 1
            if filial == 1:
                verificar_campo(filial,'campo_filial.png')
            elif filial == 2:
                verificar_campo(filial,'campo_filial_2.png')
            elif filial == 3:
                verificar_campo(filial,'.png')
            elif filial == 4:
                verificar_campo(filial,'.png')
            elif filial == 5:
                verificar_campo(filial,'.png')

            verificar_campo(filial,'campo_observacao.png')
            verificar_campo(filial,'campo_ocorrencia.png')
            click_image('salvar_datachegada.png')
            for i in range(10):
                pyautogui.press("backspace")
            pyautogui.write(str(data_chegada))
            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.write(str(data_entrega))
            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.write(str(data_descarregamento))
            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.write("aaa")
            pyautogui.sleep(time)
            for i in range(2):
                pyautogui.press("tab")
            pyautogui.write("111")
            pyautogui.sleep(time)
            click_image('efetuar_baixa.png')
            pyautogui.sleep(2)
            print('Verificando se teve nota selecionada')
            if na_tela('nenhuma_nota_selecionada.png'):
                continue        
            finalizar_baixa('em_branco.png')
            click_image('cancelar.png')

            baixado = 'SIM'
            atualizar_status_man(filial, man, baixado)

        pyautogui.sleep(5)

click_image('cancelar.png')
click_image('botao_voltar.png')
pyautogui.sleep(2)
alt_press("f4")
click_image('fechar_rodopar.png')
print('fim')