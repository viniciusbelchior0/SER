import sqlite3
import pandas as pd

# Caminho para seu banco de dados
caminho_banco = r"G:\Drives compartilhados\FINANÇAS - Documentos e Arquivos\SER\ser.db"

# Conecta ao banco
conn = sqlite3.connect(caminho_banco)

# Lê os dados da tabela 'recibos' usando pandas
#df = pd.read_sql_query("SELECT * FROM pessoas", conn)
df2 = pd.read_sql_query("SELECT * FROM produtos", conn)
#df3 = pd.read_sql_query("SELECT * FROM recibos", conn)
#df4 = pd.read_sql_query("SELECT * FROM itens_recibo", conn)

# Fecha a conexão
conn.close()

# Exporta para um arquivo Excel
#df.to_excel("pessoas_exportados.xlsx", index=False)
df2.to_excel("produtos_exportados.xlsx", index=False)
#df3.to_excel("recibos_exportados.xlsx", index=False)
#df4.to_excel("itens_exportados.xlsx", index=False)

print("Operação Concluída")
