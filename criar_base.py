import sqlite3

# Conexão com o banco de dados (será criado se não existir)
conexao = sqlite3.connect('notas_fiscais.db')
cursor = conexao.cursor()

# Criar a tabela
cursor.execute('''
CREATE TABLE IF NOT EXISTS notas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filial INTEGER NOT NULL,
    nota INTEGER NOT NULL,
    data_nota_fiscal TEXT,
    data_chegada TEXT,
    data_entrega TEXT,
    data_descarreg TEXT,                                 
    status TEXT,
    baixado TEXT,
    MDFe INTEGER NOT NULL,
    num_carga TEXT,
    UNIQUE(nota, MDFe) -- Restringe combinações duplicadas de nota e MDFe
)
''')

print("Banco de dados e tabela criados com sucesso!")

# Fechar a conexão
conexao.commit()
conexao.close()