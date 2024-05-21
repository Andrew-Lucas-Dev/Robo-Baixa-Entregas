import pandas as pd
import os
import pyautogui

# Caminho para o arquivo TXT
arquivo_txt = r'C:\Users\Andrew Lucas\Desktop\automacao\BAIXAS EBBA.txt'
arquivo_excel = r'C:\Users\Andrew Lucas\Desktop\automacao\arquivotxt.xlsx'

# Verificar se o arquivo TXT existe
if not os.path.exists(arquivo_txt):
    print(f"Erro: O arquivo {arquivo_txt} não existe.")
else:
    try:
        # Tentar ler o arquivo como UTF-16
        try:
            lines = []
            with open(arquivo_txt, 'r', encoding='utf-16') as f:
                for line in f:
                    lines.append(line.strip().split(','))
        except UnicodeError:
            # Se falhar, tentar ISO-8859-1
            lines = []
            with open(arquivo_txt, 'r', encoding='ISO-8859-1') as f:
                for line in f:
                    lines.append(line.strip().split(','))
        
        # Encontrar o número máximo de colunas
        max_columns = max(len(line) for line in lines)
        
        # Adicionar campos vazios para linhas com menos colunas que o máximo
        for line in lines:
            while len(line) < max_columns:
                line.append('')
        
        # Transformar a lista de listas em DataFrame
        df = pd.DataFrame(lines[1:], columns=lines[0])
        
        # Verificar o DataFrame
        #print(f"DataFrame lido do arquivo:\n{df.head()}")
        
        # Escrever o DataFrame no arquivo Excel
        df.to_excel(arquivo_excel, index=False)
        
        print(f"Arquivo Excel salvo em: {arquivo_excel}")
            
    except Exception as e:
        print(f"Erro ao converter o arquivo: {e}")


Planilha_Documentaçao = pd.read_excel("arquivotxt.xlsx")
# print(Planilha_Documentaçao)
for linha in Planilha_Documentaçao.index:
    nf = Planilha_Documentaçao.loc[linha,"Documento"] 
    status = Planilha_Documentaçao.loc[linha,"Status"]
    data = Planilha_Documentaçao.loc[linha,'""Data Finalização""']
    print(f"nf:{nf}  status:{status}  data:{data}")
