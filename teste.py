import pandas as pd
from openpyxl import load_workbook
import os

# Carregar os dados dos arquivos Excel
Planilha_1 = pd.read_excel("Pasta1.xlsx")
Planilha_1.fillna(0, inplace=True)  # Preencher células vazias com 0
#print(Planilha_1)

Planilha_2 = pd.read_excel("Pasta2.xlsx")
Planilha_2.fillna(0, inplace=True)  # Preencher células vazias com 0
#print(Planilha_2)

# Juntar os dados das duas planilhas
combined_df = pd.concat([Planilha_1, Planilha_2], ignore_index=True)
combined_df.fillna(0, inplace=True)  # Preencher células vazias com 0

# Identificar linhas duplicadas na coluna 'NF'
duplicates = combined_df[combined_df.duplicated(subset='NF', keep=False)]

# Marcar as duplicatas que atendem aos critérios desejados ('STATUS' igual a 'ENTREGUE' ou 'BAIXADO' igual a 'SIM')
duplicates_to_keep = duplicates[(duplicates['STATUS'] == 'ENTREGUE') & (duplicates['BAIXADO'] == 'SIM')]

# Marcar todas as duplicatas
combined_df['is_duplicate'] = combined_df.duplicated(subset='NF', keep=False)

# Manter apenas as linhas que não são duplicatas ou que atendem aos critérios
filtered_df = combined_df[~((combined_df['is_duplicate'] == True) & ~(combined_df.index.isin(duplicates_to_keep.index)))]

# Remover a coluna de marcação
filtered_df.drop(columns=['is_duplicate'], inplace=True)

filtered_df = filtered_df.drop_duplicates(subset='NF', keep='first')

# Mostrar o DataFrame resultante
#print(filtered_df)

#Impressão dos dados sem duplicatas
for linha in filtered_df.index:
    nf = filtered_df.loc[linha, "NF"]
    data = filtered_df.loc[linha, "DATA"]
    status = filtered_df.loc[linha, "STATUS"]
    baixado = filtered_df.loc[linha, "BAIXADO"]

    if baixado == 0 and status == 'ENTREGUE':
        print('ENTREGUE')
        filtered_df.loc[linha, "BAIXADO"] = 'SIM'  
    elif baixado == 0 and status == 'NÃO ENTREGUE':
        print('NÃO ENTREGUE')

filtered_df = filtered_df[(combined_df['STATUS'] == 'ENTREGUE') & (filtered_df['BAIXADO'] == 'SIM')]
filtered_df = filtered_df.sort_values(by=['DATA', 'NF'], ascending=[True, False])

# Nome do arquivo e da planilha
file_path = "Pasta2.xlsx"
sheet_name = "Sheet1"

# Salva o DataFrame no arquivo Excel, substituindo o arquivo existente, se houver
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"O DataFrame foi salvo na planilha '{sheet_name}' do arquivo {file_path}.")