import sqlite3
import pandas as pd
import re
import pyautogui
import os 
from datetime import datetime, timedelta
import subprocess
import ctypes
import shutil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys          
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

caminho = os.getcwd() 
app_path = r"C:\Users\Usuario\Desktop\VR_PROD_JETTA.rdp"
# caminho_comprovantes = r'C:\Users\Usuario\Documents\COMPRANTES'
# caminho_comprovantes_sistema = r'T:\Users\Usuario\Documents\COMPRANTES'
time = 0.5

def click_selenium(selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")

def processar_notas_em_lote(lista_notas):
    """
    Processa um lote de notas fiscais, inserindo ou atualizando conforme necessário.
    
    :param lista_notas: Lista de dicionários contendo as informações das notas fiscais.
    """
    with sqlite3.connect('banco_dados_entregas2.db') as conexao:
        cursor = conexao.cursor()

        for nota in lista_notas:
            # print(nota)
            try:
                filial = nota['filial']
                serie = nota['serie']
                numero_nota = nota['nota']
                data_nota = nota['data_nota']
                data_chegada = nota['data_chegada']
                data_entrega = nota['data_entrega']
                data_descarreg = nota['data_descarreg']
                status = nota['status']
                cod_oco = nota['cod_oco']
                cte = nota['cte']
                man = nota['man']
                num_carga = nota['num_carga']
                num_carga = str(num_carga)
                cc = nota['cc']
                atendente = nota['atendente']
                baixado = nota['baixado']

                # Verifica se a combinação nota + MDFe já existe
                cursor.execute('SELECT id FROM notas2 WHERE nota = ? AND MDFe = ?', (numero_nota, man))
                resultado = cursor.fetchone()

                if resultado:
                    print(f"A nota {numero_nota} já está associada ao MDFe {man}. Registrando novo histórico.")

                # Tenta inserir uma nova linha no banco de dados
                cursor.execute(''' 
                INSERT INTO notas2 (filial,serie, nota, data_nota_fiscal, data_chegada, data_entrega, data_descarreg, status,cod_oco,Cte, MDFe, num_carga, CC,Atendente, baixado)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?, ?, ?,?,'NAO')''', 
                (filial,serie, numero_nota, data_nota, data_chegada, data_entrega, data_descarreg, status,cod_oco,cte, man, num_carga, cc,atendente))
                
                print(f"Nota {numero_nota} processada com sucesso no MDFe {man}.")

            except sqlite3.IntegrityError:
                print(f"Nota {numero_nota} ja inserida, verificando novo status.")
                # Atualiza o registro existente caso o status seja 'Entregue'
                if status == 'Entregue' and baixado == 'NAO':
                    cursor.execute(''' 
                    UPDATE notas2
                    SET 
                        filial = ?, 
                        serie = ?,
                        data_nota_fiscal = ?, 
                        data_chegada = ?, 
                        data_entrega = ?, 
                        data_descarreg = ?, 
                        status = ?, 
                        cod_oco = ?,
                        Cte = ?,
                        MDFe = ?, 
                        num_carga = ?,
                        CC = ?,
                        Atendente = ?,
                        baixado = '?',
                    WHERE nota = ? AND MDFe = ?''', 
                    (filial,serie, data_nota, data_chegada, data_entrega, data_descarreg, status,cod_oco,cte,man, num_carga, numero_nota, man, cc,atendente, baixado))
                    
                    print(f"Nota {numero_nota} atualizada para status 'Entregue' no MDFe {man}.")
                else:
                    print(f"Nota {numero_nota} ja inserida e nao atualizada.")
                
        conexao.commit()

def atualizar_status(nota, baixado, manifesto, filial):
    conexao = sqlite3.connect('banco_dados_entregas2.db')
    cursor = conexao.cursor()

    cursor.execute('''
    UPDATE notas2
    SET baixado = ?
    WHERE nota = ? AND MDFe = ? AND filial = ?
    ''', (baixado, nota, manifesto, filial))
    
    conexao.commit()
    print(f"Status da nota {nota} atualizado para {baixado}.")
    conexao.close()



def listar_notas():
    conexao = sqlite3.connect('banco_dados_entregas2.db')
    cursor = conexao.cursor()

    # Selecionar todas as notas não baixadas e ordenar por MDFe e filial
    cursor.execute("""
        SELECT * 
        FROM notas2 
        WHERE baixado = 'NAO' 
        ORDER BY MDFe ASC, filial ASC
    """)
    notas = cursor.fetchall()
    conexao.close()
    return notas

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

def alt_press(key):
    pyautogui.keyDown('alt')
    pyautogui.press(key)
    pyautogui.keyUp('alt')

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

def imagem_na_tela(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

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

#enquanto o campo de data nota nao estiver vazio apertar os botoes
def finalizar_baixa(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            click_('yes.png')
            click_('yes_marcado.png')
            click_('ok.png')
            click_('cancelar.png')
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(0.5)

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

def verificar_campo(image_name, confidence=0.9):
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
            pyautogui.write("1")
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
    
    click_image('cancelar.png')
    click_image('digitar_data.png')
    for i in range(10):
        pyautogui.press("backspace")
    pyautogui.write(str(data_nota_fiscal))
    pyautogui.sleep(1)  
    click_nota('digitar_nota.png')
    pyautogui.sleep(1)
    pyautogui.write('00')
    pyautogui.sleep(0.2)
    pyautogui.write(str(nf))
    pyautogui.sleep(1)
    click_image('atualizar.png')
    pyautogui.sleep(1)
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Nota encontrada na tela.")
                return True
        except Exception as e:
            print("Nota não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)
    print("Número máximo de tentativas atingido. Nota não encontrada.")
    return False
 

# def processar_coluna_data(df, coluna):
#     df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
#     df[coluna] = df[coluna].dt.date
#     df[coluna] = df[coluna].astype(str).str.strip()  # Converter para string e remover espaços em branco
#     df[coluna] = pd.to_datetime(df[coluna], format='%Y-%m-%d', errors='coerce')  # Converter para datetime
#     df[coluna] = df[coluna] + pd.Timedelta(hours=22)  # Adicionar o horário '22:00'
#     df[coluna] = df[coluna].dt.strftime('%d/%m/%Y %H:%M')  # Formatar no formato desejado

# def processar_coluna_chegada(df, coluna):
#     df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
#     df[coluna] = df[coluna] + pd.Timedelta(hours=22)
#     df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')

# def processar_datas(df, colunas):
#     for coluna in colunas:
#         if 'Data Entrega' in coluna:
#             df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=1)
#         elif 'Fim Descarreg.' in coluna:
#             df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors='coerce') + pd.Timedelta(minutes=2)
#         df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')
        
def formatar_datas(df, colunas):
    for coluna in colunas:
        df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
        df[coluna] = df[coluna].dt.strftime('%d/%m/%Y%H:%M')

# Função para extrair o número completo do MDF-e e formatá-lo
def extrair_mdf_e(mdf_e):
    #print(mdf_e)
    if pd.isna(mdf_e):
        return None
    match = re.search(r"(?:1/2|2/1|2/2|/5/1|5/2|3/2|3/U|3/13|3/18|19/1)/([\d.]+)", str(mdf_e))
    return  match.group(1).replace(".", "") if match else None

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
    #match = re.search(r"MDF-e: \d+/(\d+)/", str(mdf_e))
    match = re.search(r"MDF-e: [^/]+/([^/]+)/[^/]+", str(mdf_e))

    try:
        filial = int(match.group(1))
    except:
        filial = str(match.group(1))
    return filial if match else None

# Função comum para processar a planilha
def processar_planilha(nome_arquivo, colunas_para_remover, coluna_data, coluna_mdf_filial):
    Planilha = pd.read_excel(nome_arquivo)
    
    Planilha.drop(columns=colunas_para_remover, inplace=True)



    #print(Planilha[coluna_mdf_filial])
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

    Planilha = Planilha.dropna(subset=['Status da entrega'])
    Planilha = Planilha[Planilha['Status da entrega'].str.strip() != ""]

    Planilha = Planilha.rename(columns={'N° NF': 'NF', 'Data NF': 'DATA NOTA FISCAL', 'Status da entrega': 'STATUS'})
    # Planilha['Data Entrega'] = Planilha['Data Chegada']
    # Planilha['Fim Descarreg.'] = Planilha['Data Chegada']

    Planilha['DATA NOTA FISCAL'] = pd.to_datetime(Planilha['DATA NOTA FISCAL']).dt.strftime('%d/%m/%Y')

    # colunas_de_data = ['Data Entrega', 'Fim Descarreg.']
    # for coluna in colunas_de_data:
    #     processar_coluna_data(Planilha, coluna)
    # processar_datas(Planilha, colunas_de_data)
    # processar_coluna_chegada(Planilha, 'Data Chegada')

    Planilha = Planilha.dropna(axis=1, how='all')
    
    # if 'STATUS' in Planilha.columns:
    #     Planilha.loc[Planilha['STATUS'] == 'FINALIZADO', 'STATUS'] = 'Entregue'



    # Planilha = Planilha.dropna(subset=['STATUS'])

    try:
        Planilha.drop(columns='Plano', inplace=True)
    except:
        pass

    return Planilha

driver = webdriver.Chrome()
driver.get("https://jettatransporte-my.sharepoint.com/:f:/g/personal/jetta_bi_jettatransporte_onmicrosoft_com/EiA6eCcrmHVOi0SVjgVS4eYBTgW6NmdHNlvRSINLlAOW5g?e=qyl9wK")
driver.maximize_window()        
pyautogui.sleep(2)

#click_selenium(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/div[1]/span/span/button')
click_selenium(By.XPATH, '//*[@id="virtualized-list_3_page-0_579067"]/div[3]/span')
pyautogui.sleep(2)

click_selenium(By.XPATH, "//span[text()='Planilha CC32']")
pyautogui.sleep(5)
try:
    print("Pasta Planilha CC32...")
    corpo_email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Planilha de Baixas CC32.xlsx']")))
    action_chains = ActionChains(driver)
    action_chains.context_click(corpo_email).perform()
except Exception as e:
    print("Erro ao clicar no corpo do e-mail:", e)

pyautogui.sleep(5)
click_selenium(By.XPATH, "//span[text()='Baixar']")

pyautogui.sleep(20)

arquivos = [
    'Planilha de Baixas CC32.xlsx'
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

colunas_para_remover_32 = [
    'Cnpj cliente', 'Cliente', 'Cidade', 'Peso', 'Qtde', 'Manifesto', 'Bairro'
]

Planilha_CC32 = processar_planilha('Planilha de Baixas CC32.xlsx', colunas_para_remover_32, 'Carga', 'Carga')
Planilha_Ocorrencias = pd.read_excel('OCORRENCIA.xlsx')
Planilha_Observação = pd.read_excel('OBSERVAÇÃO.xlsx')
#print(Planilha_CC32.columns)



def ajustar_date(data_str):
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
    
def ajustar_data(data_str,hora_str):
    """
    Corrige a data caso o mês e o dia estejam invertid1    :param data_str: Data em formato string (DD/MM/YYYY)
    :return: Data corrigida em formato string (DD/MM/YYYY) ou a original se válida
    """
    # if not data_str:
    #     return None  # Retorna None se a data estiver ausente
    # if not hora_str:
    #     return None  # Retorna None se a data estiver ausente
    try:

        data_str = data_str.strftime("%Y-%m-%d")
        hora_str = hora_str.strftime("%H:%M:%S")
        data_str = datetime.strptime(f"{data_str} {hora_str}", "%Y-%m-%d %H:%M:%S")
        data_str = data_str.strftime("%d/%m/%Y%H:%M")
        return data_str
    except ValueError:
        return data_str  # Retorna a data original se não puder ser interpretada

def remover_duplicados_por_filial(manifestos):
    vistos = set()  # Armazena as combinações (Filial, Manifesto) já vistas
    resultado = []  # Lista filtrada
    for manifesto in manifestos:
        chave = (manifesto['Filial'], manifesto['Manifesto'])  # Combinação única de Filial e Manifesto
        if chave not in vistos:  # Verifica se a combinação (Filial, Manifesto) já foi processada
            resultado.append(manifesto)  # Adiciona o manifesto se for único
            vistos.add(chave)  # Marca a combinação como vista
    return resultado

notas_para_processar = []

for index, row in Planilha_CC32.iterrows():
    #print(row.get)
    try:
        nota = row.get('NF')

        serie = row.get('Serie')
        try:
            serie = int(serie)
        except:
            serie = str(serie)

        #VER AS FORMATAÇOES DAS DATAS
        #-------------------------------------------------------------------------------------
        data_chegada = row.get('Data Chegada')  # Exemplo: '05/03/2025'
        hora_chegada = row.get('Horario Chegada')  # Exemplo: '14:30'
        
        if pd.isna(data_chegada):
            print('SEM DATA1')
            print(nota)
            continue

        if pd.isna(hora_chegada):
            hora_chegada = pd.to_datetime("00:10:00", format="%H:%M:%S").time()
            
        data_hora_chegada = ajustar_data(data_chegada,hora_chegada)


        data_entrega = row.get('Data Entrega')  # Exemplo: '05/03/2025'
        hora_entrega = row.get('Horario Entrega')  # Exemplo: '14:30'

        if pd.isna(data_entrega):
            print('SEM DATA2')
            print(nota)
            data_hora_entrega = ajustar_data(data_chegada,hora_chegada)

        else:
            data_hora_entrega = ajustar_data(data_entrega,hora_entrega)

        data_fim_descarreg = row.get('Data Entrega')  # Exemplo: '05/03/2025'
        horario_fim_descarreg = row.get('Horario Entrega')  # Exemplo: '14:30'

        if pd.isna(data_fim_descarreg):
            print('SEM DATA3')
            data_hora_fim_descarreg = ajustar_data(data_chegada,hora_chegada)
        else:
            data_hora_fim_descarreg = ajustar_data(data_fim_descarreg,horario_fim_descarreg)
        
        #-------------------------------------------------------------------------------------

        #COLOCAR CTE
        nota_dict = {
            'filial': row.get('Filial', 'Desconhecida'),
            'serie': serie,
            'nota': row.get('NF'),
            'data_nota': ajustar_date(row.get('DATA NOTA FISCAL')),
            'data_chegada': data_hora_chegada,
            'data_entrega': data_hora_entrega,
            'data_descarreg': data_hora_fim_descarreg,
            'status': row.get('STATUS', 'Pendente'),
            'cod_oco': row.get('Cod Observação', 'Pendente'),
            'cte': row.get('Ct-e/OST'),
            'man': row.get('MDF-e'),
            'num_carga': row.get('Carga', None),
            'atendente': row.get('Atendente', None),
            'cc': 32,
            'baixado': row.get('baixado', None),
        }

        # print(nota_dict)

        # Validação básica
        if not nota_dict['nota'] or not nota_dict['man']:
            raise ValueError(f"Nota ou MDFe ausente na linha {index}.")

        # Adicionar à lista de notas para processamento
        notas_para_processar.append(nota_dict)

    except Exception as e:
        print(f"Erro ao processar a linha {index}: {e}")
        print(row.get)
        continue

# Passar a lista de notas para a função de processamento em lote
if notas_para_processar:
    processar_notas_em_lote(notas_para_processar)
    #print(notas_para_processar)
    print(f"{len(notas_para_processar)} notas processadas com sucesso.")
else:
    print("Nenhuma nota válida encontrada para processar.")

print('--------------------------------------------------------------')

login()

notas = listar_notas()

# Criar DataFrame
df = pd.DataFrame(notas, columns=['id', 'filial', 'serie', 'nota', 'data_nota_fiscal', 'data_chegada', 'data_entrega', 'data_descarreg', 'status', 'cod_oco','Cte', 'MDFe', 'num_carga','cc','Atendente', 'baixado'])

resultado = []  # Lista para salvar os manifestos ou status

# Agrupar por MDFe e Filial
grupos = df.groupby(["MDFe", "filial"])
print('Iniciando baixas ')



for _, linha in df.iterrows():
    filial = linha["filial"]
    nf = linha["nota"]
    manifesto = linha["MDFe"]
    data_nota_fiscal = linha["data_nota_fiscal"]
    data_chegada = linha["data_chegada"]
    data_entrega = linha["data_entrega"]
    data_descarregamento = linha["data_descarreg"]
    referencia = linha["num_carga"]
    status = linha["status"]
    cod_oco = linha["cod_oco"]
    baixado1 = linha["baixado"]
    nf = int(nf)
    filial = int(filial)
    carga = linha["num_carga"]
    atendente = linha["Atendente"]
    cte = linha["Cte"]
    if referencia is None:
        continue

    # Convertendo a string para um objeto datetime
    data_referencia = datetime.strptime("15/12/2024", "%d/%m/%Y")
    data = datetime.strptime(data_nota_fiscal, "%d/%m/%Y")
    
    # Comparação
    if data_referencia > data:
        print("A data é maior que 15/12/2024.")
        continue
    
    print(f"Filial: {filial} Nota: {nf} man:{manifesto} carga:{referencia} data nota:{data_nota_fiscal} cheg:{data_chegada} entre:{data_entrega} desc:{data_descarregamento} status:{status} cod:{cod_oco}")




    if pd.isna(atendente):  
        atendente = 'NAO INFORMADO'
    else:
        try:
            atendente = int(float(atendente))
        except:
            atendente = str(atendente)
    cte = str(cte)
    
    if status in Planilha_Ocorrencias['DESCRICAO OCORRENCIA'].values:
        ocorrencia = Planilha_Ocorrencias.loc[Planilha_Ocorrencias[Planilha_Ocorrencias['DESCRICAO OCORRENCIA'] == status].index.values, 'COD OCORRENCIA'].values[0]
    else:
        print('Cadastrar ocorrencia')
        continue

    if cod_oco in Planilha_Observação['DESCRICAO OBSERVACAO'].values:
        cod_observacao = Planilha_Observação.loc[Planilha_Observação[Planilha_Observação['DESCRICAO OBSERVACAO'] == cod_oco].index.values, 'COD OBS'].values[0]
    else:
        print('Cadastrar observação')
        continue


    #EDITAVEL EQUIPE ANA
    try:
        cod_ocorrencia =  Planilha_CC32.loc[linha, "Cod Ocorrencia"] 
    except:
        cod_ocorrencia = ''
    try:
        observacao =  Planilha_CC32.loc[linha, "Observação"] 
    except:
        observacao = ''
    
    
 


    # data_chegada = data_chegada.strftime("%d/%m/%Y")
    # data_chegada = str(data_chegada)
    # try:
    #     data_entrega = data_entrega.strftime("%d/%m/%Y")
    #     data_fim_descarregamento = data_fim_descarregamento.strftime("%d/%m/%Y")
    #     data_entrega = str(data_entrega)
    #     data_fim_descarregamento = str(data_fim_descarregamento)
    # except:
    #     pass


    # # Converter para datetime adicionando uma data fictícia
    # hor_chegada = datetime.combine(datetime.today(), hor_chegada)

    # # Somar 1 minuto
    # hor_chegada_dev = (hor_chegada + timedelta(minutes=5)).strftime("%H%M")
    # hor_chegada_dev2 = (hor_chegada + timedelta(minutes=10)).strftime("%H%M")

    # # Converter hor_chegada original para string HHMM
    # hor_chegada = hor_chegada.strftime("%H%M")

    # # Convertendo para string (opcional, pois já é string após strftime)
    # hor_chegada = str(hor_chegada)
    # hor_chegada_dev = str(hor_chegada_dev)
    # hor_chegada_dev2 = str(hor_chegada_dev2)

    # try:
    #     hor_entrega = hor_entrega.strftime("%H%M")
    #     hor_fim_descarregamento = hor_fim_descarregamento.strftime("%H%M")
    #     hor_entrega = str(hor_entrega)
    #     hor_fim_descarregamento = str(hor_fim_descarregamento)
    # except:
    #     pass

    print(f'nf:{nf} carga:{carga} man:{manifesto} status:{status} data:{data_chegada}')


    click_image('cancelar.png')
    pyautogui.sleep(time)
    click_image('referencia.png')
    for i in range(10):
        pyautogui.press("backspace")
    pyautogui.write(str(carga))
    pyautogui.sleep(time) 
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
    if imagem_encontrada('nota_encontrada.png'):
     
        click_image('salvar_filial.png')
        pyautogui.write("3")
        pyautogui.sleep(time)       
        click_image('salvar_ocorrencia.png')
        pyautogui.write(str(ocorrencia))
        pyautogui.sleep(time)      
        click_image('salvar_observ.png')
        pyautogui.write(str(cod_observacao))
        pyautogui.sleep(time)           
        pyautogui.press("tab")

        #AJUSTAR CONFIRMAÇÃO DE PREENCHIMENTO DE CAMPO
        # verificar_campo('campo_filial.png')
        # verificar_campo('campo_observacao.png')
        # verificar_campo('campo_ocorrencia.png')

        click_image('salvar_datachegada.png')
        for i in range(10):
            pyautogui.press("backspace")
        pyautogui.write(str(data_chegada))
        pyautogui.sleep(time)
        pyautogui.press('tab')
        if ocorrencia == 1:
            try:
                pyautogui.write(str(data_entrega))
                pyautogui.sleep(time)
                pyautogui.press('tab')
                pyautogui.write(str(data_descarregamento))
            except:
                pyautogui.write(str(data_chegada))
                pyautogui.sleep(time)
                pyautogui.press('tab')
                pyautogui.write(str(data_chegada))

            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.write("aaa")
            pyautogui.sleep(time)
            for i in range(2):
                pyautogui.press("tab")
            pyautogui.write("111")

        else:
            pyautogui.sleep(time)
            pyautogui.write(str(data_chegada))
            click_nota('nota_dev.png')
            pyautogui.write('1')
            pyautogui.sleep(time)

        pyautogui.sleep(time)
        click_nota('observacao.png')
        pyautogui.sleep(1) 
        pyautogui.write('Atendende: ')
        pyautogui.sleep(1) 
        pyautogui.write(str(atendente))
        pyautogui.write(' ')
        pyautogui.sleep(1) 
        if observacao:  # Verifica se a variável não está vazia
            pyautogui.write('Codigo ocorrencia:')
            pyautogui.sleep(1) 
            pyautogui.write(str(cod_ocorrencia))
            pyautogui.write(' ')
            pyautogui.sleep(1) 
            pyautogui.write('Observacao: ') 
            pyautogui.sleep(1) 
            pyautogui.write(str(observacao))
            pyautogui.write(' ')
            pyautogui.sleep(time)
            pyautogui.press('tab')
            pyautogui.sleep(5)
            
        click_image('efetuar_baixa.png')
        pyautogui.sleep(time)
        finalizar_baixa('digitar_data.png')  

        click_image('cancelar.png')
        baixado = 'SIM'
        atualizar_status(nf, baixado,manifesto,filial)

click_image('cancelar.png')
click_image('botao_voltar.png')
pyautogui.sleep(2)
alt_press("f4")
click_image('fechar_rodopar.png')
print('fim')