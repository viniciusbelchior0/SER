import pandas as pd
import sqlite3

# Banco de Dados
caminho_banco = r"G:\Drives compartilhados\FINANÇAS - Documentos e Arquivos\SER\ser.db"
conn = sqlite3.connect(caminho_banco)
cursor = conn.cursor()

#Planilha - Produtos
sheet_id = "1dzgoydq32wdbK71mVZpyKDl3m1uRgiKdhRwzP_zqFKs"
gid = "0"  # geralmente 0 para a primeira aba

url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
produtos = pd.read_csv(url)
produtos['valor_unitario'] = (
    produtos['valor_unitario']
    .str.replace('R\$', '', regex=True)
    .str.replace('.', '', regex=False)
    .str.replace(',', '.', regex=False)
    .str.strip()
    .astype(float)                  
)

# Opção mais simples - DELETAR DADOS DO BANCO e apenas inserí-los
cursor.execute("DELETE FROM produtos")
    
for _, row in produtos.iterrows():
    try:
        cursor.execute("INSERT INTO produtos (codigo_produto, descricao_produto, valor_unitario, codigo_receita, rubrica, tipo_receita) VALUES (?, ?, ?, ?, ?, ?)", (row['codigo_produto'], row['descricao_produto'], row['valor_unitario'], row['codigo_receita'], row['rubrica'], row['tipo_receita']))
    except sqlite3.IntegrityError:
        pass  

conn.commit()
conn.close()

print("Operação Concluída")
