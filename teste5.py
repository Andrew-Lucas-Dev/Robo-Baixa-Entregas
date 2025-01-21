import sqlite3
import pandas as pd
import re
from datetime import datetime

# def inserir_nota(filial, nota, data_nota, data_chegada, data_entrega, data_descarreg, status, man, num_carga):
#     # Conexão com o banco de dados
#     with sqlite3.connect('banco_dados_entregas.db') as conexao:
#         cursor = conexao.cursor()

#         try:
#             # Verifica se a combinação nota + MDFe já existe
#             cursor.execute('SELECT id FROM notas WHERE nota = ? AND MDFe = ?', (nota, man))
#             resultado = cursor.fetchone()

#             if resultado:
#                 print(f"A nota {nota} já está associada ao MDFe {man}. Registrando novo histórico.")

#             # Tenta inserir uma nova linha no banco de dados
#             cursor.execute(''' 
#             INSERT INTO notas (filial, nota, data_nota_fiscal, data_chegada, data_entrega, data_descarreg, status, baixado, MDFe, num_carga)
#             VALUES (?, ?, ?, ?, ?, ?, ?, 'NAO', ?, ?)''', 
#             (filial, nota, data_nota, data_chegada, data_entrega, data_descarreg, status, man, num_carga))
            
#             conexao.commit()
#             print(f"Nota {nota} processada com sucesso no MDFe {man}.")

#         except sqlite3.IntegrityError:
#             print(f"Erro ao inserir a nota {nota}. Verifique os dados.")
#             # Atualiza o registro existente caso o status seja 'Entregue'
#             if status == 'Entregue':
#                 cursor.execute(''' 
#                 UPDATE notas
#                 SET 
#                     filial = ?, 
#                     data_nota_fiscal = ?, 
#                     data_chegada = ?, 
#                     data_entrega = ?, 
#                     data_descarreg = ?, 
#                     status = ?, 
#                     baixado = 'NAO', 
#                     MDFe = ?, 
#                     num_carga = ?
#                 WHERE nota = ? AND MDFe = ?''', 
#                 (filial, data_nota, data_chegada, data_entrega, data_descarreg, status, man, num_carga, nota, man))
                
#                 conexao.commit()
#                 print(f"Nota {nota} atualizada para status 'Entregue' no MDFe {man}.")


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
        Planilha.loc[Planilha['STATUS'].str.upper() == 'FINALIZADO', 'STATUS'] = 'Entregue'

    # Tentar remover a coluna 'Plano' se existir
    if 'Plano' in Planilha.columns:
        Planilha.drop(columns='Plano', inplace=True)


    return Planilha

# Parâmetros de colunas a remover
colunas_para_remover_cc19 = [
    'Série', 'Cnpj cliente', 'Cliente', 'Cidade', 'Ct-e/OST', 'Peso', 'Qtde', 'Vlr Merc.',
    'Entrega Canhoto Físico'
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
Planilha_CC16 = processar_planilha("Pasta1.xlsx", colunas_para_remover_cc16, 'Plano', 'Plano')
Planilha_CC21 = processar_planilha("Planilha de Baixas CC21.xlsx", colunas_para_remover_cc21, 'N° Carga', 'N° Carga')
# print(Planilha_CC19)

# # # # Juntar as 4 planilhas
combined_df = pd.concat([Planilha_CC19,Planilha_CC15,Planilha_CC16,Planilha_CC21], ignore_index=True)

#-----------------------------------------------------------------------------------------
#           UM A UM


# #Iterar pelas linhas do DataFrame e inserir cada linha no banco de dados
# for index, row in combined_df.iterrows():
#     filial = row['Filial']
#     nota = row['NF']
#     data_nota = row['DATA NOTA FISCAL']
#     data_chegada = row['Data Chegada']
#     data_entrega = row['Data Entrega']
#     data_descarreg = row['Fim Descarreg.']
#     status = row['STATUS']
#     man = row['MDF-e']
#     num_carga = row.get('N° Carga', None)  # Ajuste caso a coluna não exista em todas as planilhas
#     data_ajustada = ajustar_data(data_nota)
#     print(f'nota:{nota} data_nota:{data_ajustada}')
    # try:
    #     inserir_nota(filial, nota, data_nota, data_chegada, data_entrega, data_descarreg, status, man, num_carga)
    # except Exception as e:
    #     print(f"Erro ao processar a linha {index}: {e}")
    #     exit()
#----------------------------------------------------------------------


def ajustar_data(data_str):
    """
    Corrige a data caso o mês e o dia estejam invertidos.
    :param data_str: Data em formato string (DD/MM/YYYY)
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
    if grupo["status"].isin(["Entregue", "entregue"]).all():
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
        # Baixar notas individuais
        # print(f"Baixando notas do manifesto {mdfe} (Filial {filial}) carga:{serie}...")
        for _, linha in grupo.iterrows():
            if linha["status"] in ["Entregue", "entregue"]:
                nf = linha["nota"]
                manifesto = linha["MDFe"]
                data_nota = linha["data_nota_fiscal"]
                data_chegada = linha["data_chegada"]
                data_entrega = linha["data_entrega"]
                data_descarregamento = linha["data_descarreg"]
                carga = linha["num_carga"]
                baixado1 = linha["baixado"]
                nf = int(nf)
                filial = int(filial)
                if carga is None:
                    continue
                print(f"Filial: {filial} Nota: {nf} man:{manifesto} carga:{carga} data nota:{data_nota} cheg:{data_chegada} entre:{data_entrega} desc:{data_descarregamento} baixado:{baixado1}")
                baixado = 'SIM'
                # atualizar_status(nf, baixado,manifesto,filial)

print('------------------------------------')
print('Baixando Manifestos')
# Remover os duplicados
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
        print(f"Filial: {filial} serie:{serie} Manifesto: {man} data nota:{data_nota_fiscal} cheg:{data_chegada} entre:{data_entrega} desc:{data_descarregamento}")
        baixado = 'SIM'
        man = int(man)
        filial = int(filial)
        # atualizar_status_man(filial, man, baixado)

print('fim')