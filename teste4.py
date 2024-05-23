import pandas as pd



Planilha_1 = pd.read_excel(r"C:\Users\Usuario\OneDrive\Pasta de trabalho 1.xlsx")
#print(Planilha_1)

Planilha_2 = pd.read_excel(r"C:\Users\Usuario\OneDrive\Pasta de trabalho.xlsx")
#print(Planilha_2)

Planilha_3 = pd.read_excel(r"C:\Users\Usuario\OneDrive\Pasta1.xlsx")

for linha in Planilha_3.index:
    nf = Planilha_3.loc[linha,'N° NF'] 
    status = Planilha_3.loc[linha,'Status da entrega']
    data_emissao = Planilha_3.loc[linha,'Data NF']
    data_baixa = str(Planilha_3.loc[linha, 'DATA  ENTREGA'])
    data_baixa = pd.to_datetime(data_baixa, errors='coerce')
    data_baixa = data_baixa.strftime('%d/%m/%Y')
    print(f'nota:{nf} data emissao:{data_emissao} status:{status} data baixa:{data_baixa}')  
    
# combined_df = pd.concat([Planilha_1, Planilha_2])

# # Salva a planilha combinada
# combined_df.to_excel(r'C:\Users\Usuario\Desktop\Robo-Baixa-Entregas\planilha_combinada.xlsx', index=False)
# print(combined_df)

