import pandas as pd
import os
import pyautogui
from datetime import datetime, timedelta

caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)

data_ontem = datetime.now().date() - timedelta(days=1)
data_ontem = data_ontem.strftime("%d/%m/%Y")

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

# # Caminho para o arquivo TXT
# arquivo_txt = caminho + r'\BAIXAS EBBA.txt'
# arquivo_excel = caminho +r'\arquivotxt.xlsx'

# # Verificar se o arquivo TXT existe
# if not os.path.exists(arquivo_txt):
#     print(f"Erro: O arquivo {arquivo_txt} não existe.")
# else:
#     try:
#         # Tentar ler o arquivo como UTF-16
#         try:
#             lines = []
#             with open(arquivo_txt, 'r', encoding='utf-16') as f:
#                 for line in f:
#                     lines.append(line.strip().split(','))
#         except UnicodeError:
#             # Se falhar, tentar ISO-8859-1
#             lines = []
#             with open(arquivo_txt, 'r', encoding='ISO-8859-1') as f:
#                 for line in f:
#                     lines.append(line.strip().split(','))
        
#         # Encontrar o número máximo de colunas
#         max_columns = max(len(line) for line in lines)
        
#         # Adicionar campos vazios para linhas com menos colunas que o máximo
#         for line in lines:
#             while len(line) < max_columns:
#                 line.append('')
        
#         # Transformar a lista de listas em DataFrame
#         df = pd.DataFrame(lines[1:], columns=lines[0])
        
#         # Verificar o DataFrame
#         #print(f"DataFrame lido do arquivo:\n{df.head()}")
        
#         # Escrever o DataFrame no arquivo Excel
#         df.to_excel(arquivo_excel, index=False)
        
#         print(f"Arquivo Excel salvo em: {arquivo_excel}")
            
#     except Exception as e:
#         print(f"Erro ao converter o arquivo: {e}")

# Planilha_OBA = pd.read_excel("arquivotxt.xlsx")

# Ler o arquivo CSV com a codificação correta
try:
    Planilha_Obba = pd.read_csv("e88e679c26037377383d8ad3df789ad0baac5a30.csv", delimiter=",", encoding='ISO-8859-1')
except UnicodeDecodeError:
    # Tente outra codificação se a primeira falhar
    Planilha_Obba = pd.read_csv("e88e679c26037377383d8ad3df789ad0baac5a30.csv", delimiter=",", encoding='Windows-1252')
# Verificar se os dados foram lidos corretamente

#print(Planilha_Obba)

# Ler o arquivo CSV com a codificação correta
try:
    Planilha_Dori = pd.read_csv("BAIXAS DORI.csv", delimiter=";", encoding='ISO-8859-1')
except UnicodeDecodeError:
    # Tente outra codificação se a primeira falhar
    Planilha_Dori = pd.read_csv("BAIXAS DORI.csv", delimiter=";", encoding='Windows-1252')
# Verificar se os dados foram lidos corretamente

Planilha_CACAU = pd.read_excel("BAIXAS CACAU SHOW.xlsx", skiprows=6)

# Selecionar as colunas com os números das notas fiscais
coluna_notas_1 = 'Documento'  # Substitua pelo nome da coluna na primeira planilha
coluna_data_emissao_1 = 'EmissÃ£o'
coluna_status_1 = 'Status'
coluna_finalizacao_1 = 'Data FinalizaÃ§Ã£o'

coluna_notas_2 = 'NF'          # Substitua pelo nome da coluna na segunda planilha
coluna_data_emissao_2 = 'Emissão'
coluna_status_2 = 'Status'
coluna_finalizacao_2 = 'Data de baixa'

coluna_notas_3 = 'Nota Fiscal'    # Substitua pelo nome da coluna na terceira planilha
coluna_data_emissao_3 = 'Data Emissão NF-e'
coluna_status_3 = 'Status'
coluna_finalizacao_3 = 'Data de Entrega'

# Criar DataFrames apenas com a coluna de notas fiscais
notas_df1 = Planilha_Obba[[coluna_notas_1, coluna_data_emissao_1,coluna_status_1,coluna_finalizacao_1]].rename(columns={coluna_notas_1: 'NumeroNotaFiscal', coluna_data_emissao_1: 'DataNotaFiscal', coluna_status_1: 'Status',coluna_finalizacao_1: 'DataEntrega'})
notas_df2 = Planilha_Dori[[coluna_notas_2, coluna_data_emissao_2,coluna_status_2,coluna_finalizacao_2]].rename(columns={coluna_notas_2: 'NumeroNotaFiscal', coluna_data_emissao_2: 'DataNotaFiscal', coluna_status_2: 'Status',coluna_finalizacao_2: 'DataEntrega'})
notas_df3 = Planilha_CACAU[[coluna_notas_3, coluna_data_emissao_3,coluna_status_3,coluna_finalizacao_3]].rename(columns={coluna_notas_3: 'NumeroNotaFiscal', coluna_data_emissao_3: 'DataNotaFiscal', coluna_status_3: 'Status',coluna_finalizacao_3: 'DataEntrega'})

# Converter as colunas de datas para o formato datetime
notas_df1['DataNotaFiscal'] = pd.to_datetime(notas_df1['DataNotaFiscal'], dayfirst=True)
notas_df1['DataEntrega'] = pd.to_datetime(notas_df1['DataEntrega'], dayfirst=True)

notas_df2['DataNotaFiscal'] = pd.to_datetime(notas_df2['DataNotaFiscal'], dayfirst=True)
notas_df2['DataEntrega'] = pd.to_datetime(notas_df2['DataEntrega'], dayfirst=True)

notas_df3['DataNotaFiscal'] = pd.to_datetime(notas_df3['DataNotaFiscal'], dayfirst=True)
notas_df3['DataEntrega'] = pd.to_datetime(notas_df3['DataEntrega'], dayfirst=True)

# Concatenar os DataFrames
df_concatenado = pd.concat([notas_df1, notas_df2, notas_df3], ignore_index=True)
print(df_concatenado)
# # #LOGIN
# pyautogui.press("capslock")  # Desativa o CAPS LOCK se estiver ativado
# # # pyautogui.keyDown('win')
# # # pyautogui.press("m")
# # # pyautogui.keyUp('win')
# # click_image('logo_rodopar_areatrabalho.png')#PC ESCRITORIO
# click_image('logo_rodopar_areatrabalho_casa.png')#PC CASA
# pyautogui.click()
# click_image('conectar_rodopar.png')
# click_image('senha_rodopar_1.png')
# pyautogui.write("17@mudar")
# click_image('ok_primeiro_login.png')
# click_image('sim_primeiro_login.png')
# click_image('segundo_login.png')
# pyautogui.sleep(1)
# pyautogui.write("anascimento")
# pyautogui.press("tab")
# pyautogui.write("990607")
# for i in range(2): 
#     pyautogui.press("enter")
# click_image('escolha_filial.png')
# pyautogui.press("enter")

# click_image('botao_frota.png')
# pyautogui.press('right')
# pyautogui.press('down')
# pyautogui.press('right')
# for i in range(6): 
#     pyautogui.press("down")
# pyautogui.press('right')
# pyautogui.press('enter')
# click_image('filial_manifesto.png')
# pyautogui.write("5")
# pyautogui.press('tab')
# pyautogui.write("1")
# pyautogui.press('tab')
# pyautogui.write("1")
# pyautogui.press('tab')
# click_image('ok_aviso.png')
# click_image('gerar_ocorrencia.png')

i = 0
for linha in df_concatenado.index:
    nf = df_concatenado.loc[linha,'NumeroNotaFiscal'] 
    status = df_concatenado.loc[linha,'Status']
    data_emissao = df_concatenado.loc[linha,'DataNotaFiscal']
    data_emissao = data_emissao.strftime('%d/%m/%Y')  
    data_baixa = df_concatenado.loc[linha,'DataEntrega']   
    if status == 'Entregue' or status == 'Entregue com ocorrência' or status == 'Entregue sem ocorrência':
        data_baixa = data_baixa.strftime('%d/%m/%Y')
        if data_baixa == '10/05/2024':#data_ontem: #rodar com a data da baixa do dia anterior
            print(f"nf:{nf}  status:{status}  data emissao:{data_emissao} data baixa:{data_baixa}")
            i += 1
            
            # pyautogui.sleep(2)
            # click_image('cancelar.png')
            # pyautogui.sleep(2)
            # click_image('digitar_data.png')
            # for i in range(10):
            #     pyautogui.press("backspace")
            # pyautogui.write(str(data_emissao))
            # click_nota('digitar_nota.png')
            # pyautogui.write(str(nf))
            # click_image('atualizar.png')
            # click_image('salvar_filial.png')    
            # pyautogui.write("5")
            # click_image('salvar_ocorrencia.png')
            # pyautogui.write("1")
            # click_image('salvar_observ.png')
            # pyautogui.write("1")
            # click_image('salvar_datachegada.png')
            # for i in range(10):
            #     pyautogui.press("backspace")
            # pyautogui.write(str(data_baixa))
            # pyautogui.write("17:00")
            # pyautogui.press('tab')
            # pyautogui.write(str(data_baixa))
            # pyautogui.write("17:01")
            # pyautogui.press('tab')
            # pyautogui.write(str(data_baixa))
            # pyautogui.write("17:02")
            # pyautogui.press('tab')
            # pyautogui.write("aaa")
            # for i in range(2):
            #     pyautogui.press("tab")
            # pyautogui.write("111")
            # #click_image('efetuar_baixa.png')
            # pyautogui.sleep(1)
            # click_image('cancelar.png')

print(i)
