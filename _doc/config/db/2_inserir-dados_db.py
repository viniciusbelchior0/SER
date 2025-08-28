import pandas as pd
import sqlite3

# Caminho para seu banco de dados
caminho_banco = r"G:\Drives compartilhados\FINANÇAS - Documentos e Arquivos\SER\ser.db"

# Conecta ao banco
conn = sqlite3.connect(caminho_banco)
cursor = conn.cursor()

# Importar pessoa - pessoa
pessoa = pd.read_excel("db/pessoas.xlsx")
for _, row in pessoa.iterrows():
    try:
        cursor.execute("INSERT INTO pessoas (nome_pessoa, tipo_pessoa, identificador_pessoa) VALUES (?, ?, ?)", (row['nome_pessoa'], row['tipo_pessoa'], row['identificador_pessoa']))
    except sqlite3.IntegrityError:
        pass  # Já existe


# Importar recibo - Recibos
recibo = pd.read_excel("db/recibos.xlsx")
recibo["data"] = recibo["data"].astype(str)
for _, row in recibo.iterrows():
    try:
        cursor.execute("INSERT INTO recibos (num_recibo, nome_pessoa, observacao, data, codigo_pagamento, banco, nome_responsavel) VALUES (?, ?, ?, ?, ?, ?, ?)", (row['num_recibo'], row['nome_pessoa'], row['observacao'], row['data'], row['codigo_pagamento'], row['banco'], row['nome_responsavel']))
    except sqlite3.IntegrityError:
        pass  # Já existe

# Importar item recibo - itens recibo
itens_recibo = pd.read_excel("db/itens_recibo.xlsx")
for _, row in itens_recibo.iterrows():
    try:
        cursor.execute("INSERT INTO itens_recibo (num_recibo, numero_item, descricao_produto, valor_unitario, quantidade) VALUES (?, ?, ?, ?, ?)", (row['num_recibo'], row['numero_item'], row['descricao_produto'], row['valor_unitario'], row['quantidade']))
    except sqlite3.IntegrityError:
        pass  # Já existe

conn.commit()
conn.close()
print("Operação Concluída")
