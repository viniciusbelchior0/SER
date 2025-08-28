import sqlite3

# Caminho para seu banco de dados
caminho_banco = r"G:\Drives compartilhados\FINANÇAS - Documentos e Arquivos\SER\ser.db"

conn = sqlite3.connect(caminho_banco)
cursor = conn.cursor()

nomes = [("Fabiana Aparecida Chechio Martins",), ("José Henrique da Silva Santos",),("Nilson Kendi Ogassahara",),("Pedro Luiz Leandro",)]
cursor.executemany("INSERT INTO responsaveis (nome_responsavel) VALUES (?)", nomes)

conn.commit()
conn.close()
print("Operação Concluída")