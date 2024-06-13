import pyautogui
import os
import pandas as pd
import numpy as np

caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)


# Carregar a planilha do Excel
Planilha_2 = pd.read_excel("DETALHES DAS NOTAS.xlsx")

# Remover colunas com todos os valores NaN
Planilha_2 = Planilha_2.dropna(axis=1, how='all')

# Selecionar as colunas específicas
Planilha_2 = Planilha_2[['NF', 'STATUS', 'DT NF', 'CHEGADA', 'FIM DESCARGA', 'ENTREGA']]

# Filtrar linhas onde 'STATUS' é 'ATRASADA' ou 'NO PRAZO'
Planilha_2 = Planilha_2[(Planilha_2['STATUS'] == 'ATRASADA') | (Planilha_2['STATUS'] == 'NO PRAZO')]

Planilha_2['NF'] = Planilha_2['NF'].astype(np.int64)#problema do .0
# Mostrar o DataFrame resultante
print(Planilha_2)








