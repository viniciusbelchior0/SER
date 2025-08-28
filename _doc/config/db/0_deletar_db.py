import sqlite3

# Caminho para seu banco de dados
caminho_banco = r"G:\Drives compartilhados\FINANÇAS - Documentos e Arquivos\SER\ser.db"

# Conecta ao banco
conn = sqlite3.connect(caminho_banco)
cursor = conn.cursor()

# Lê os dados da tabela 'recibos' usando pandas
try:
    cursor.execute("DROP TABLE pessoas")
except:
    pass

try:
    cursor.execute("DROP TABLE produtos")
except:
    pass

try:
    cursor.execute("DROP TABLE recibos")
except:
    pass

try:
    cursor.execute("DROP TABLE itens_recibo")
except:
    pass

try:
    cursor.execute("DROP TABLE responsaveis")
except:
    pass

# Fecha a conexão
conn.close()
print("Operação Concluída")
