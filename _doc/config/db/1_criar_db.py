import sqlite3

# Caminho para seu banco de dados
caminho_banco = r"G:\Drives compartilhados\FINANÇAS - Documentos e Arquivos\SER\ser.db"

# Conexão com banco
conn = sqlite3.connect(caminho_banco)
cursor = conn.cursor()

# Tabelas
cursor.execute('''
CREATE TABLE IF NOT EXISTS pessoas (
    id_pessoa INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_pessoa TEXT UNIQUE,
    tipo_pessoa TEXT,
    identificador_pessoa TEXT    
)''')

cursor.execute('''
CREATE TABLE IF NOT EXISTS produtos (
    codigo_produto INTEGER PRIMARY KEY AUTOINCREMENT,
    descricao_produto TEXT UNIQUE,
    valor_unitario REAL,
    codigo_receita TEXT,
    rubrica TEXT,
    tipo_receita TEXT
)''')

cursor.execute('''
CREATE TABLE IF NOT EXISTS recibos (
    num_recibo INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_pessoa TEXT,
    observacao TEXT,
    data TEXT,
    codigo_pagamento TEXT,
    banco TEXT,
    nome_responsavel TEXT
)''')

cursor.execute('''
CREATE TABLE IF NOT EXISTS itens_recibo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    num_recibo INTEGER,
    numero_item INTEGER,
    descricao_produto TEXT,
    valor_unitario REAL,
    quantidade INTEGER,
    FOREIGN KEY (num_recibo) REFERENCES recibos(num_recibo)
)''')

cursor.execute('''
CREATE TABLE IF NOT EXISTS responsaveis (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_responsavel TEXT
)''')

conn.commit()
conn.close()
print("Operação Concluída")