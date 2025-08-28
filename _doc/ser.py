import tkinter as tk
from tkinter import ttk, messagebox, font
import sqlite3
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
import io
import os
import sys
import re
import smtplib
from email.message import EmailMessage
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import webbrowser
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER
from ttkwidgets.autocomplete import AutocompleteCombobox
import pandas as pd
import string

# Caminho base do Google Drive local
caminho_relativo = r"Drives compartilhados\FINANÇAS - Documentos e Arquivos\SER"
CAMINHO_BASE = None

# Procura a pasta nas unidades disponíveis
for letra in string.ascii_uppercase:
    unidade = f"{letra}:"
    caminho_teste = os.path.join(unidade, caminho_relativo)
    if os.path.exists(caminho_teste):
        CAMINHO_BASE = caminho_teste
        break

if CAMINHO_BASE:
    # Agora usamos o CAMINHO_BASE para montar os outros caminhos
    caminho_banco = os.path.join(CAMINHO_BASE, "ser.db")
    base_recibo_controle = os.path.join(CAMINHO_BASE, "modelos", "modelo_recibo_controle.pdf")
    base_recibo_emissao = os.path.join(CAMINHO_BASE, "modelos", "modelo_recibo_emissao.pdf")

PASTA_PRIMEIRA = os.path.join(CAMINHO_BASE, "controle")
PASTA_SEGUNDA = os.path.join(CAMINHO_BASE, "emissao")

#os.makedirs(PASTA_PRIMEIRA, exist_ok=True)
#os.makedirs(PASTA_SEGUNDA, exist_ok=True)

# Conexão com banco
conn = sqlite3.connect(caminho_banco)
cursor = conn.cursor()

#Registrar Arial
pdfmetrics.registerFont(TTFont('Arial', 'C:/Windows/Fonts/arial.ttf'))
pdfmetrics.registerFont(TTFont('ArialBlack', 'C:/Windows/Fonts/ariblk.ttf'))

def gerar_recibo():
    # Janela principal
    janela = tk.Toplevel()
    janela.title("Gerar Recibo")

    #janela.iconbitmap(caminho_icone)
    janela.geometry("770x420")
    janela.configure(bg="#F7F7F7")

    def carregar_pessoas():
        cursor.execute("SELECT nome_pessoa FROM pessoas ORDER BY nome_pessoa")
        return [row[0] for row in cursor.fetchall()]

    def carregar_produtos():
        cursor.execute("SELECT descricao_produto FROM produtos ORDER BY descricao_produto")
        return [row[0] for row in cursor.fetchall()]
    
    def carregar_responsavel():
        cursor.execute("SELECT nome_responsavel FROM responsaveis ORDER BY nome_responsavel")
        return [row[0] for row in cursor.fetchall()]

    def buscar_valor_padrao(produto):
        cursor.execute("SELECT valor_unitario FROM produtos WHERE descricao_produto = ?", (produto,))
        r = cursor.fetchone()
        return r[0] if r else 0.0

    #Recibo
    tk.Label(janela, text="Gerar Recibo",font=("Segoe UI", 16, "bold"), fg="#000000", bg="#F7F7F7").grid(row=0, column=2, columnspan=2)
    tk.Label(janela, text="",bg="#F7F7F7").grid(row=1, column=0)#linha vazia

    # Nome
    tk.Label(janela, text="Interessado", font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=2, column=0, sticky="ne")
    combo_nome = AutocompleteCombobox(janela, completevalues=carregar_pessoas(), width=65)
    combo_nome.grid(row=2, column=1, columnspan=4, sticky="w", padx=(0, 5))

    def formatar_identificador(identificador: str) -> str:
        identificador = re.sub(r'\D', '', identificador)  # Remove tudo que não for número

        if len(identificador) == 11:  # CPF
            return f"{identificador[:3]}.{identificador[3:6]}.{identificador[6:9]}-{identificador[9:]}"
        elif len(identificador) == 14:  # CNPJ
            return f"{identificador[:2]}.{identificador[2:5]}.{identificador[5:8]}/{identificador[8:12]}-{identificador[12:]}"
        else:
            return identificador  # Sem formatação se incompleto

    def janela_cadastro_interessado():
        janela_cadastro = tk.Toplevel()
        janela_cadastro.title("Cadastrar novo interessado")
        janela_cadastro.geometry("350x175")
        janela_cadastro.configure(bg="#F7F7F7")

        # Nome
        tk.Label(janela_cadastro, text="Nome:",font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        entry_nome = tk.Entry(janela_cadastro, width=30)
        entry_nome.grid(row=0, column=1, padx=10, pady=5)

        # Tipo
        tk.Label(janela_cadastro, text="Tipo:",font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        combo_tipo_pessoa = ttk.Combobox(janela_cadastro, values=['FISICA','JURIDICA'],width=27)
        combo_tipo_pessoa.grid(row=1, column=1, padx=10, pady=5)

        # Identificador (CPF/CNPJ/etc)
        tk.Label(janela_cadastro, text="Identificador:",font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        entry_identificador = tk.Entry(janela_cadastro, width=30)
        entry_identificador.grid(row=2, column=1, padx=10, pady=5)

        def salvar_novo_interessado():
            nome = entry_nome.get().strip()
            tipo = combo_tipo_pessoa.get().strip()
            identificador = entry_identificador.get().strip()

            if not nome:
                messagebox.showerror("Erro", "O campo Nome é obrigatório.")
                return
            
            identificador_formatado = formatar_identificador(identificador)

            if len(re.sub(r'\D', '', identificador)) not in [0,11, 14]:
                messagebox.showerror("Erro", "CPF ou CNPJ inválido. CPF deve conter 11 caracteres e CNPJ deve conter 14 caracteres (Ambos sem pontuação, devem ser inseridos apenas números.)")
                return

            try:
                cursor.execute(
                    "INSERT INTO pessoas (nome_pessoa, tipo_pessoa, identificador_pessoa) VALUES (?, ?, ?)",
                    (nome, tipo, identificador_formatado)
                )
                conn.commit()
                messagebox.showinfo("Sucesso", "Interessado cadastrado com sucesso.")
                combo_nome['values'] = carregar_pessoas()
                combo_nome.set(nome)
                janela_cadastro.destroy()
            except sqlite3.IntegrityError:
                messagebox.showinfo("Atenção", "Interessado já cadastrado.")
            except Exception as e:
                messagebox.showerror("Erro inesperado", str(e))

        botao_salvar = tk.Button(janela_cadastro, text="Salvar", command=salvar_novo_interessado)
        botao_salvar.grid(row=3, column=0, columnspan=2, pady=15)
        janela_cadastro.grab_set()


    #Editar Interessado
    def janela_editar_interessado():
        janela_editar = tk.Toplevel()
        janela_editar.title("Editar Interessado")
        janela_editar.geometry("360x200")
        janela_editar.configure(bg="#F7F7F7")

        tk.Label(janela_editar, text="Interessado:",font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=0, column=0, padx=10, pady=5, sticky="e")

        nomes_editar = carregar_pessoas()  # função que retorna lista de nomes existentes
        combo_nomes_editar = AutocompleteCombobox(janela_editar, completevalues=nomes_editar, width=27)
        combo_nomes_editar.grid(row=0, column=1, padx=10, pady=5)

        # Campos a serem preenchidos
        tk.Label(janela_editar, text="Nome:",font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        entry_nome_editar = tk.Entry(janela_editar, width=30)
        entry_nome_editar.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(janela_editar, text="Tipo:",font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        entry_tipo_editar = ttk.Combobox(janela_editar, values=['FISICA','JURIDICA'],width=27)
        entry_tipo_editar.grid(row=2, column=1, padx=10, pady=5)

        tk.Label(janela_editar, text="Identificador:",font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        entry_identificador_editar = tk.Entry(janela_editar, width=30)
        entry_identificador_editar.grid(row=3, column=1, padx=10, pady=5)

        def carregar_dados(*args):
            nome_selecionado = combo_nomes_editar.get()
            if not nome_selecionado:
                return
            cursor.execute("SELECT nome_pessoa, tipo_pessoa, identificador_pessoa FROM pessoas WHERE nome_pessoa = ?", (nome_selecionado,))
            pessoa = cursor.fetchone()
            if pessoa:
                entry_nome_editar.delete(0, tk.END)
                entry_nome_editar.insert(0, pessoa[0])
                entry_tipo_editar.delete(0, tk.END)
                entry_tipo_editar.insert(0, pessoa[1] or "")
                entry_identificador_editar.delete(0, tk.END)
                entry_identificador_editar.insert(0, pessoa[2] or "")

        def atualizar():
            nome_antigo = combo_nomes_editar.get()
            novo_nome = entry_nome_editar.get().strip()
            novo_tipo = entry_tipo_editar.get().strip()
            novo_identificador = entry_identificador_editar.get().strip()

            if not novo_nome:
                messagebox.showerror("Erro", "O campo Nome é obrigatório.")
                return

            try:
                cursor.execute("""
                    UPDATE pessoas
                    SET nome_pessoa = ?, tipo_pessoa = ?, identificador_pessoa = ?
                    WHERE nome_pessoa = ?
                """, (novo_nome, novo_tipo, novo_identificador, nome_antigo))
                conn.commit()
                messagebox.showinfo("Sucesso", "Cadastro atualizado com sucesso.")
                janela_editar.destroy()
            except Exception as e:
                messagebox.showerror("Erro ao atualizar", str(e))

        combo_nomes_editar.bind("<<ComboboxSelected>>", carregar_dados)
        btn_atualizar = tk.Button(janela_editar, text="Atualizar", command=atualizar)
        btn_atualizar.grid(row=4, column=0, columnspan=2, pady=15)

        janela_editar.grab_set()

    # Botões Editar e Cadastrar Interessado
    frame_botoes = ttk.Frame(janela)
    frame_botoes.grid(row=2, column=4, sticky="w", padx=3)

    btn_cadastrar = ttk.Button(frame_botoes, text="Cadastrar", command=janela_cadastro_interessado)
    btn_cadastrar.pack(side="left", padx=(0, 3))

    btn_editar = ttk.Button(frame_botoes, text="Editar", command=janela_editar_interessado)
    btn_editar.pack(side="right")

    # Botões editar e cadastrar antigos
    #tk.Button(janela, text="Cadastrar", command= janela_cadastro_interessado, width=8).grid(row=2, column=5)
    #tk.Button(janela, text="Editar", width=8, command=janela_editar_interessado).grid(row=2, column=6)

    # Tipo Pagamento
    tk.Label(janela, text="Pagamento", font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=3, column=0, sticky="w")
    combo_tipo_pagamento = ttk.Combobox(janela, values=['Transferência Bancária', 'Moeda','Cheque'], width=20)
    combo_tipo_pagamento.grid(row=3, column=1, sticky="w", padx=(0,10))

    # Cod. Pagamento
    tk.Label(janela, text= "Cod. Pagamento", font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=3, column=2,sticky="w", padx=(5,2))
    entry_codigo_pagamento = tk.Entry(janela, width=15)
    entry_codigo_pagamento.grid(row=3, column=3, sticky="w")

    # Banco
    tk.Label(janela, text="Banco", font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=4, column=0, sticky="e", padx=(5,2))
    entry_banco = tk.Entry(janela, width=22)
    entry_banco.grid(row=4, column=1, sticky="w")

    # Responsável pela geração do Recibo
    tk.Label(janela, text="Responsável", font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=5, column=0, sticky="ne")
    combo_responsavel = AutocompleteCombobox(janela, completevalues=carregar_responsavel(), width=30)
    combo_responsavel.grid(row=5, column=1)

    # Imprimir data
    var_imprimir_data = tk.BooleanVar()
    var_imprimir_data.set(True)
    check_data = tk.Checkbutton(janela, text="Imprimir Data?", variable=var_imprimir_data, fg="#606161", bg="#F7F7F7", font=("Segoe UI", 10, "bold"))
    check_data.grid(row=5, column=2)

    #Linhas vazias
    tk.Label(janela, text="",bg="#F7F7F7").grid(row=6,column=0)

    # Tabela de produtos
    tk.Label(janela, text="Descrição", font=("Segoe UI", 10, "bold"), fg="#000000" ,bg="#F7F7F7").grid(row=7, column=0, columnspan=3)
    tk.Label(janela, text="Quantidade", font=("Segoe UI", 10, "bold"), fg="#000000" ,bg="#F7F7F7").grid(row=7, column=3)
    tk.Label(janela, text="Valor Unitário", font=("Segoe UI", 10, "bold"), fg="#000000" ,bg="#F7F7F7").grid(row=7, column=4)
    tk.Label(janela, text="Valor Total", font=("Segoe UI", 10, "bold"), fg="#000000" ,bg="#F7F7F7").grid(row=7, column=5)

    combos_produto = []
    entry_qtd = []
    entry_val_unit = []
    label_total = []

    def atualizar_total(index):
        try:
            qtd = int(entry_qtd[index].get())
            val = float(entry_val_unit[index].get().replace(",", "."))
            total = qtd * val
            label_total[index]['text'] = f"R$ {total:.2f}"
        except:
            label_total[index]['text'] = "R$ 0.00"

    def ao_selecionar_produto(index):
        prod = combos_produto[index].get()
        valor = buscar_valor_padrao(prod)
        entry_val_unit[index].delete(0, tk.END)
        entry_val_unit[index].insert(0, f"{valor:.2f}")
        atualizar_total(index)

    produtos = carregar_produtos()
    for i in range(5):
        cb = AutocompleteCombobox(janela, completevalues=produtos, width=42)
        cb.grid(row=8+i, column=0, columnspan=3)
        cb.bind("<<ComboboxSelected>>", lambda e, idx=i: ao_selecionar_produto(idx))
        combos_produto.append(cb)

        qtd = tk.Entry(janela, width=5)
        qtd.grid(row=8+i, column=3)
        qtd.insert(0, "1")
        qtd.bind("<KeyRelease>", lambda e, idx=i: atualizar_total(idx))
        entry_qtd.append(qtd)

        val = tk.Entry(janela, width=10)
        val.grid(row=8+i, column=4)
        val.bind("<KeyRelease>", lambda e, idx=i: atualizar_total(idx))
        entry_val_unit.append(val)

        lbl = tk.Label(janela, text="R$ 0.00" ,bg="#F7F7F7")
        lbl.grid(row=8+i, column=5)
        label_total.append(lbl)

    # Observações
    def limitar_caracteres(texto):
        return len(texto) <= 215

    vcmd = (janela.register(limitar_caracteres), '%P')

    tk.Label(janela, text="Observações", font=("Segoe UI", 10, "bold"), fg="#606161" ,bg="#F7F7F7").grid(row=15, column=0)
    entry_obs = tk.Entry(janela, width=85, validate='key', validatecommand=vcmd)
    entry_obs.grid(row=15, column=1, columnspan=6, sticky="we", padx=5) ##avaliar wingspan com base nesse

    #Linha Vazia
    tk.Label(janela,text="",bg="#F7F7F7").grid(row=13,column=0)

    def criar_overlay_controle(dados):
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)

        estilo_paragrafo = ParagraphStyle(
            name="Observacao",
            fontName="Arial",
            fontSize=8,
            leading=10,
            alignment=0
        )

        def draw_if_exists(func, *args, value=None):
            """Chama o método do canvas apenas se value não for None ou vazio."""
            if value not in [None, '', 0, 0.0]:
                func(*args, str(value))

        def draw_currency(x, y, val):
            """Desenha moeda formatada se val existir."""
            if val not in [None, '', 0, 0.0]:
                texto = f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                c.drawString(x, y, texto)

        # 1ª VIA
        c.setFont("ArialBlack", 24)
        draw_if_exists(c.drawString, 463, 775, value=dados.get('numero_recibo'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 100, 742, value=dados.get('nome'))
        draw_if_exists(c.drawString, 100, 732, value=dados.get('identificador_nome'))
        draw_if_exists(c.drawString, 490, 742, value=dados.get('tipo_pagamento'))
        draw_if_exists(c.drawString, 490, 729, value=dados.get('codigo_pagamento'))
        draw_if_exists(c.drawString, 490, 716, value=dados.get('banco'))

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 665 - (i - 0.59) * 28, value=dados.get(f'rubrica_item_{i}'))

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 659 - (i - 0.59) * 28, value=dados.get(f'tipo_receita_item_{i}'))    

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 654 - (i - 0.59) * 28, value=dados.get(f'cod_item_{i}'))

        c.setFont("Arial", 9)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 112, 653 - (i - 0.75) * 28, value=dados.get(f'produto_{i}'))
            draw_currency(374, 653 - (i - 0.75) * 28, dados.get(f'val_{i}'))
            draw_if_exists(c.drawString, 451, 653 - (i - 0.75) * 28, value=dados.get(f'qtd_{i}'))

        c.setFont("ArialBlack", 8)
        for i in range(1, 6):
            draw_currency(500, 653 - (i - 0.75) * 28, dados.get(f'total_{i}'))

        c.setFont("ArialBlack", 11)
        draw_currency(455, 492, dados.get('subtotal'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 39, 450, value=dados.get('data'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 39, 440, value=dados.get('responsavel'))

        # Observação 1ª via
        obs_texto = dados.get('observacao', '')
        if obs_texto.strip():
            paragrafo = Paragraph(f"Observação: {obs_texto}", estilo_paragrafo)
            obs_1via = Frame(x1=40, y1=475, width=300, height=42, showBoundary=False)
            obs_1via.addFromList([paragrafo], c)

        # 2ª VIA
        c.setFont("ArialBlack", 24)
        draw_if_exists(c.drawString, 463, 357, value=dados.get('numero_recibo'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 100, 328, value=dados.get('nome'))
        draw_if_exists(c.drawString, 100, 318, value=dados.get('identificador_nome'))
        draw_if_exists(c.drawString, 490, 328, value=dados.get('tipo_pagamento'))
        draw_if_exists(c.drawString, 490, 315, value=dados.get('codigo_pagamento'))
        draw_if_exists(c.drawString, 490, 302, value=dados.get('banco'))

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 254 - (i - 0.55) * 28, value=dados.get(f'rubrica_item_{i}'))
        
        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 248 - (i - 0.55) * 28, value=dados.get(f'tipo_receita_item_{i}'))

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 243 - (i - 0.55) * 28, value=dados.get(f'cod_item_{i}'))

        c.setFont("Arial", 9)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 112, 239 - (i - 0.75) * 28, value=dados.get(f'produto_{i}'))
            draw_currency(374, 239 - (i - 0.75) * 28, dados.get(f'val_{i}'))
            draw_if_exists(c.drawString, 457, 239 - (i - 0.75) * 28, value=dados.get(f'qtd_{i}'))

        c.setFont("ArialBlack", 8)
        for i in range(1, 6):
            draw_currency(500, 239 - (i - 0.75) * 28, dados.get(f'total_{i}'))

        c.setFont("ArialBlack", 11)
        draw_currency(455, 73, dados.get('subtotal'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 39, 35, value=dados.get('data'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 39, 25, value=dados.get('responsavel'))

        # Observação 2ª via
        if obs_texto.strip():
            obs_2via = Frame(x1=40, y1=64, width=300, height=42, showBoundary=False)
            obs_2via.addFromList([paragrafo], c)

        c.save()
        buffer.seek(0)
        return buffer

    def criar_overlay_emissao(dados):
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)

        estilo_paragrafo = ParagraphStyle(
            name="Observacao",
            fontName="Arial",
            fontSize=8,
            leading=10,
            alignment=0
        )

        def draw_if_exists(func, *args, value=None):
            """Chama o método do canvas apenas se value não for None ou vazio."""
            if value not in [None, '', 0, 0.0]:
                func(*args, str(value))

        def draw_currency(x, y, val):
            """Desenha moeda formatada se val existir."""
            if val not in [None, '', 0, 0.0]:
                texto = f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                c.drawString(x, y, texto)

        # 1ª VIA
        c.setFont("ArialBlack", 24)
        draw_if_exists(c.drawString, 463, 775, value=dados.get('numero_recibo'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 100, 742, value=dados.get('nome'))
        draw_if_exists(c.drawString, 100, 732, value=dados.get('identificador_nome'))
        draw_if_exists(c.drawString, 490, 742, value=dados.get('tipo_pagamento'))
        draw_if_exists(c.drawString, 490, 729, value=dados.get('codigo_pagamento'))
        draw_if_exists(c.drawString, 490, 716, value=dados.get('banco'))

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 665 - (i - 0.59) * 28, value=dados.get(f'rubrica_item_{i}'))

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 659 - (i - 0.59) * 28, value=dados.get(f'tipo_receita_item_{i}'))    

        c.setFont("ArialBlack", 6)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 35, 654 - (i - 0.59) * 28, value=dados.get(f'cod_item_{i}'))

        c.setFont("Arial", 9)
        for i in range(1, 6):
            draw_if_exists(c.drawString, 112, 653 - (i - 0.75) * 28, value=dados.get(f'produto_{i}'))
            draw_currency(374, 653 - (i - 0.75) * 28, dados.get(f'val_{i}'))
            draw_if_exists(c.drawString, 451, 653 - (i - 0.75) * 28, value=dados.get(f'qtd_{i}'))

        c.setFont("ArialBlack", 8)
        for i in range(1, 6):
            draw_currency(492, 650 - (i - 0.75) * 28, dados.get(f'total_{i}'))

        c.setFont("ArialBlack", 11)
        draw_currency(455, 492, dados.get('subtotal'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 39, 450, value=dados.get('data'))

        c.setFont("Arial", 9)
        draw_if_exists(c.drawString, 39, 440, value=dados.get('responsavel'))

        # Observação 1ª via
        obs_texto = dados.get('observacao', '')
        if obs_texto.strip():
            paragrafo = Paragraph(f"Observação: {obs_texto}", estilo_paragrafo)
            obs_1via = Frame(x1=40, y1=475, width=300, height=42, showBoundary=False)
            obs_1via.addFromList([paragrafo], c)

        c.save()
        buffer.seek(0)
        return buffer

    def aplicar_overlay_controle(dados, base_pdf, saida):
        overlay_buffer = criar_overlay_controle(dados)

        base = PdfReader(base_pdf)
        overlay = PdfReader(overlay_buffer)
        writer = PdfWriter()

        base_page = base.pages[0]
        overlay_page = overlay.pages[0]

        # Mescla a camada de texto (overlay) na página base
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)

        with open(saida, "wb") as f_out:
            writer.write(f_out)

        print(f"PDF final gerado: {saida}")

    def aplicar_overlay_emissao(dados, base_pdf, saida):
        overlay_buffer = criar_overlay_emissao(dados)

        base = PdfReader(base_pdf)
        overlay = PdfReader(overlay_buffer)
        writer = PdfWriter()

        base_page = base.pages[0]
        overlay_page = overlay.pages[0]

        # Mescla a camada de texto (overlay) na página base
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)

        with open(saida, "wb") as f_out:
            writer.write(f_out)

        print(f"PDF final gerado: {saida}")

    def salvar_recibos(dados, primeira_via, segunda_via, base_recibo_controle, base_recibo_emissao):
        """
        Gera o PDF a partir de dados e salva nas pastas primeira_via e segunda_via
        do Google Drive sincronizado localmente.
        """

        # Gera o nome do arquivo (ajuste conforme sua preferência)
        #numero = dados.get("numero_recibo","sem_numero")
        #nome_arquivo = f"recibo_{numero}.pdf"

        # Caminhos completos
        controle = os.path.join(primeira_via,f"recibo_{dados['numero_recibo']}.pdf")
        emissao = os.path.join(segunda_via,f"recibo_{dados['numero_recibo']}.pdf")

        #Arquivos
        #base_recibo_controle = os.path.join(CAMINHO_BASE, "modelos", "modelo_base_controle.pdf")
        #base_recibo_emissao = os.path.join(CAMINHO_BASE, "primeira_via", "modelo_base_emissao.pdf")
        base_recibo_controle = base_recibo_controle
        base_recibo_emissao = base_recibo_emissao

        # Gera e salva os dois arquivos
        aplicar_overlay_controle(dados, base_pdf=base_recibo_controle, saida=controle)
        aplicar_overlay_emissao(dados, base_pdf=base_recibo_emissao, saida=emissao)

    #Obter código da receita do produto
    def obter_codigo_receita(descricao_produto, cursor):
        cursor.execute("SELECT codigo_receita FROM produtos WHERE descricao_produto = ?", (descricao_produto,))
        codigo_receita = cursor.fetchone()
        return codigo_receita[0] if codigo_receita else None
    
    def obter_tipo_receita(descricao_produto, cursor):
        cursor.execute("SELECT tipo_receita FROM produtos WHERE descricao_produto = ?", (descricao_produto,))
        tipo_receita = cursor.fetchone()
        return tipo_receita[0] if tipo_receita else None
    
    def obter_rubrica_item(descricao_produto, cursor):
        cursor.execute("SELECT rubrica FROM produtos WHERE descricao_produto = ?", (descricao_produto,))
        rubrica_item = cursor.fetchone()
        return rubrica_item[0] if rubrica_item else None

    def obter_identificador_nome(nome_pessoa, cursor):
        cursor.execute("SELECT identificador_pessoa FROM pessoas WHERE nome_pessoa = ?", (nome_pessoa,))
        identificador_nome = cursor.fetchone()
        return identificador_nome[0] if identificador_nome else None

    def limitar_caracteres(texto):
        return len(texto) <= 215

    vcmd = (janela.register(limitar_caracteres), '%P')

    def carregar_lista_emails():
        #Planilha - Produtos
        sheet_id = "1LP35c4ikhT56y_deoUyZPYVSqDAZ19mGGNuetNvOjCk"
        gid = "0"  # geralmente 0 para a primeira aba

        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        df = pd.read_csv(url)

        lista_emails = df['lista_emails'].dropna().tolist()
        return lista_emails

    #Enviar email
    def enviar_email(dados,segunda_via):
        remetente = "ser.stf.fcav@gmail.com"
        senha_app = "ltlxboevgtdxyean"
        dados = dados

        # Nome e caminho completo do PDF
        nome_pdf = f"recibo_{dados['numero_recibo']}.pdf"
        caminho_pdf = os.path.join(segunda_via, nome_pdf)

        #Lista de emails
        lista_emails = carregar_lista_emails()

        # Criar e configurar a mensagem
        msg = EmailMessage()
        msg['Subject'] = f"SER | Novo recibo gerado - Recibo Nº {dados['numero_recibo']}"
        msg['From'] = remetente
        msg['To'] = lista_emails #"vinicius.belchior@unesp.br","fabiana.chechio@unesp.br","nilson.kendi@unesp.br","pedro.leandro@unesp.br","jose.s.santos@unesp.br"
        msg.set_content(f"{dados['responsavel']}\n \nNº do Recibo: {dados['numero_recibo']}\nInteressado: {dados['nome']}\nValor Total: R$ {dados['subtotal']}")

        # Anexar o PDF corretamente
        try:
            with open(caminho_pdf, 'rb') as f:
                msg.add_attachment(
                    f.read(),
                    maintype='application',
                    subtype='pdf',
                    filename=nome_pdf
                )
        except FileNotFoundError:
            print(f"Arquivo PDF não encontrado: {caminho_pdf}")
            return

        # Enviar e-mail
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(remetente, senha_app)
                smtp.send_message(msg)
            print("Email enviado com sucesso.")
        except Exception as e:
            print("Erro ao enviar email:", e)

    # Salvar dados
    def salvar():
        nome_pessoa = combo_nome.get()
        observacao = entry_obs.get()
        tipo_pagamento = combo_tipo_pagamento.get()
        codigo_pagamento = entry_codigo_pagamento.get()
        banco = entry_banco.get()
        responsavel = combo_responsavel.get()
        data = datetime.now().strftime("%d/%m/%Y")
        data_iso = datetime.now().strftime("%Y-%m-%d")

        #Realizar formatação da data
        data_dt = datetime.strptime(data, "%d/%m/%Y")
        meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

        if var_imprimir_data.get() == True:
            data_impressao = f"Jaboticabal, {data_dt.day} de {meses[data_dt.month - 1].capitalize()} de {data_dt.year}"
        else:
            data_impressao = ""

        if not nome_pessoa:
            messagebox.showwarning("Atenção", "Selecione ou cadastre um nome.")
            return
        
        if not tipo_pagamento:
            messagebox.showwarning("Atenção", "Selecione o tipo de pagamento.")
            return
        
        if not responsavel:
            messagebox.showwarning("Atenção", "Selecione o responsável pela emissão do recibo.")
            return

        cursor.execute("INSERT INTO recibos (nome_pessoa, observacao, data) VALUES (?, ?, ?)", (nome_pessoa, observacao, data_iso))
        conn.commit()
        recibo_id = cursor.lastrowid

        dados_recibo = {"numero_recibo": recibo_id,
                "nome":nome_pessoa,
                "observacao":observacao,
                "data": data_impressao,
                "tipo_pagamento": tipo_pagamento,
                "codigo_pagamento": codigo_pagamento,
                "banco": banco,
                "responsavel": f"Recibo emitido por: {responsavel}"}
        
        identificador_nome = obter_identificador_nome(nome_pessoa, cursor)
        dados_recibo["identificador_nome"] = identificador_nome

        if not any(cb.get().strip() for cb in combos_produto):
            messagebox.showwarning("Atenção", "Selecione ao menos um produto.")
            return
        
        total_geral = 0

        for i in range(5):
            produto = combos_produto[i].get()
            qtd = entry_qtd[i].get()
            val = entry_val_unit[i].get()

            if produto and qtd and val:
                try:
                    qtd = int(qtd)
                    val = float(val.replace(",", "."))

                    total_item = val * qtd
                    total_geral += total_item
                    codigo_receita = obter_codigo_receita(produto, cursor)
                    tipo_receita = obter_tipo_receita(produto, cursor)
                    rubrica_item = obter_rubrica_item(produto, cursor)

                    cursor.execute('''
                        INSERT INTO itens_recibo (num_recibo, numero_item, descricao_produto, valor_unitario, quantidade)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (recibo_id, i+1, produto, val, qtd))

                    #Inserir limite de caracteres para o produto aqui


                    # Dicionário
                    dados_recibo[f"produto_{i+1}"] = produto
                    dados_recibo[f"qtd_{i+1}"] = qtd
                    dados_recibo[f"val_{i+1}"] = val
                    dados_recibo[f"total_{i+1}"] = total_item
                    dados_recibo[f"cod_item_{i+1}"] = codigo_receita
                    dados_recibo[f"tipo_receita_item_{i+1}"] = tipo_receita
                    dados_recibo[f"rubrica_item_{i+1}"] = rubrica_item

                except ValueError:
                    continue

        dados_recibo["subtotal"] = total_geral

        def enviar_email_tentativa():
            try:
                enviar_email(dados_recibo, PASTA_SEGUNDA)
            except:
                pass

        conn.commit()
        salvar_recibos(dados_recibo,PASTA_PRIMEIRA,PASTA_SEGUNDA,base_recibo_controle,base_recibo_emissao)
        messagebox.showinfo("Sucesso", f"Recibo {recibo_id} gerado. Aguarde cerca de 10 segundos para o envio do email. O cadastro de um novo recibo estará liberado após esse período.")
        janela.after(10000, enviar_email_tentativa)
        limpar()

    def limpar():
        combo_nome.set('')
        combo_tipo_pagamento.set('')
        entry_codigo_pagamento.delete(0, tk.END)
        entry_banco.delete(0, tk.END)
        entry_obs.delete(0, tk.END)
        for i in range(5):
            combos_produto[i].set('')
            entry_qtd[i].delete(0, tk.END)
            entry_qtd[i].insert(0, "1")
            entry_val_unit[i].delete(0, tk.END)
            label_total[i]['text'] = "R$ 0.00"


    #Salvar Registro
    btn_gerar = tk.Button(
        janela,
        text="Gerar Recibo",
        bg="#0093DD",
        fg="#FFFFFF",
        font=("Segoe UI", 10, "bold"),
        activebackground="#007BB5",
        activeforeground="#FFFFFF",
        relief="flat",
        command=salvar
    )
    btn_gerar.grid(row=17, column=2, columnspan=2, pady=(20, 10), ipadx=10, ipady=5)

    # Efeito hover no botão
    def on_enter(e):
        btn_gerar.config(bg="#026999")
    def on_leave(e):
        btn_gerar.config(bg="#007BB5")

    btn_gerar.bind("<Enter>", on_enter)
    btn_gerar.bind("<Leave>", on_leave)
    
    #tk.Button(janela, text="Gerar Recibo", command=salvar).grid(row=15, column=3, pady=10)

def gerar_relatorio():
    # Janela principal
    janela_relatorio = tk.Toplevel()
    janela_relatorio.title("Gerar Relatório")
    janela_relatorio.geometry("580x280")
    janela_relatorio.configure(bg="#F7F7F7")

    # Estilo visual
    fonte_titulo = ("Segoe UI", 16, "bold")
    fonte_label = ("Segoe UI", 10)
    cor_titulo = "#000000"
    cor_fundo = "#F7F7F7"
    cor_botao = "#0093DD"
    cor_botao_hover = "#007BB5"
    cor_texto_botao = "#FFFFFF"

    def carregar_pessoas_relatorio():
        cursor.execute("SELECT nome_pessoa FROM pessoas ORDER BY nome_pessoa")
        return [row[0] for row in cursor.fetchall()]

    def carregar_produtos_relatorio():
        cursor.execute("SELECT descricao_produto FROM produtos ORDER BY descricao_produto")
        return [row[0] for row in cursor.fetchall()]
    
    # Cabeçalho
    tk.Label(
        janela_relatorio, text="Gerar Relatório",
        font=fonte_titulo, fg=cor_titulo, bg=cor_fundo
    ).grid(row=0, column=0, columnspan=5, pady=(15, 5))

    # Campo Interessado
    tk.Label(
        janela_relatorio, text="Interessado:",
        font=fonte_label, fg="#000000", bg=cor_fundo
    ).grid(row=1, column=0, sticky="e", padx=(10, 5), pady=5)
    combo_nome_relatorio = AutocompleteCombobox(
        janela_relatorio, completevalues=carregar_pessoas_relatorio(), width=60
    )
    combo_nome_relatorio.grid(row=1, column=1, columnspan=4, pady=5, sticky="w")

    # Campo Produto
    tk.Label(
        janela_relatorio, text="Produto:",
        font=fonte_label, fg="#000000", bg=cor_fundo
    ).grid(row=2, column=0, sticky="e", padx=(10, 5), pady=5)
    combo_produtos_relatorio = AutocompleteCombobox(
        janela_relatorio, completevalues=carregar_produtos_relatorio(), width=60
    )
    combo_produtos_relatorio.grid(row=2, column=1, columnspan=4, pady=5, sticky="w")

    
    def formatar_data_inicial(event):
        texto = entry_datainicial_relatorio.get().replace("/", "")  # remove barras anteriores
        novo_texto = ''

        for i, c in enumerate(texto):
            if not c.isdigit():
                continue
            if i == 2 or i == 4:
                novo_texto += '/'
            if i < 8:
                novo_texto += c

        entry_datainicial_relatorio.delete(0, tk.END)
        entry_datainicial_relatorio.insert(0, novo_texto)

    def formatar_data_final(event):
        texto = entry_datafinal_relatorio.get().replace("/", "")  # remove barras anteriores
        novo_texto = ''

        for i, c in enumerate(texto):
            if not c.isdigit():
                continue
            if i == 2 or i == 4:
                novo_texto += '/'
            if i < 8:
                novo_texto += c

        entry_datafinal_relatorio.delete(0, tk.END)
        entry_datafinal_relatorio.insert(0, novo_texto)

    def validar_data_inicial():
        data = entry_datainicial_relatorio.get()
        dt = datetime.strptime(data, "%d/%m/%Y")
        data_sql = dt.strftime("%Y-%m-%d")
        return data_sql

    def validar_data_final():
        data = entry_datafinal_relatorio.get()
        dt = datetime.strptime(data, "%d/%m/%Y")
        data_sql = dt.strftime("%Y-%m-%d")
        return data_sql

    
    # Data Inicial
    tk.Label(
        janela_relatorio, text="Data Inicial:",
        font=fonte_label, fg="#000000", bg=cor_fundo
    ).grid(row=3, column=0, sticky="e", padx=(10, 5), pady=5)
    entry_datainicial_relatorio = ttk.Entry(janela_relatorio, width=20)
    entry_datainicial_relatorio.grid(row=3, column=1, pady=5, sticky="w")
    entry_datainicial_relatorio.bind("<KeyRelease>", formatar_data_inicial)

    # Data Final
    tk.Label(
        janela_relatorio, text="Data Final:",
        font=fonte_label, fg="#000000", bg=cor_fundo
    ).grid(row=3, column=3, sticky="e", padx=(10, 5), pady=5)
    entry_datafinal_relatorio = ttk.Entry(janela_relatorio, width=20)
    entry_datafinal_relatorio.grid(row=3, column=4, pady=5, sticky="w")
    entry_datafinal_relatorio.bind("<KeyRelease>", formatar_data_final)


    #Filtrar os dados e gerar relatório
    def gerar_pdf():
        nome_pessoa_relatorio = combo_nome_relatorio.get()
        descricao_produto_relatorio = combo_produtos_relatorio.get()
        data_inicial_relatorio = entry_datainicial_relatorio.get()
        data_final_relatorio = entry_datafinal_relatorio.get()

        #Transformar as datas
        data_inicial_relatorio = validar_data_inicial()
        data_final_relatorio = validar_data_final()

        erros = []

        if not data_inicial_relatorio:
            erros.append("Data inicial é obrigatória.")
        if not data_final_relatorio:
            erros.append("Data final é obrigatória.")
        if not (nome_pessoa_relatorio or descricao_produto_relatorio):
            erros.append("Informe ao menos o nome ou o produto.")

        if erros:
            messagebox.showerror("Campos Obrigatórios", "\n".join(erros))
            return
        
        query = """ SELECT r.nome_pessoa,i.num_recibo,r.data , i.descricao_produto, (i.valor_unitario * i.quantidade) AS valor_total, r.observacao
                    FROM itens_recibo as i
                    INNER JOIN recibos as r
                            ON i.num_recibo = r.num_recibo
                    WHERE r.data >= ? AND r.data <= ?
                    """
        params = [data_inicial_relatorio, data_final_relatorio]

        if nome_pessoa_relatorio:
            query += "AND r.nome_pessoa = ?"
            params.append(nome_pessoa_relatorio)

        if descricao_produto_relatorio:
            query += "AND i.descricao_produto = ?"
            params.append(descricao_produto_relatorio)

        query += "ORDER BY r.data"

        cursor.execute(query, tuple(params))
        dados = cursor.fetchall()
        conn.close

        if not dados:
            print("Nenhum registro obtido.")
            return
        
        # Criar PDF
        pdf_path = "relatorio_recibos.pdf"
        doc = SimpleDocTemplate(pdf_path, pagesize=A4)

        styles = getSampleStyleSheet()
        style_normal = styles['Normal']
        elementos = []

        estilo_titulo_recibo = ParagraphStyle(
            name="TituloRelatorioRecibo",
            fontName="ArialBlack",
            fontSize=14,
            alignment=TA_CENTER,
            textColor=HexColor("#000000")
        )

        estilo_informacoes_recibo = ParagraphStyle(
            name="InformacoesRelatorioRecibo",
            fontName="Arial",
            fontSize=10,
            alignment=TA_CENTER,
            textColor=HexColor("#000000")
        )

        estilo_informacoes_recibo_negrito = ParagraphStyle(
            name="InformacoesRelatorioRecibo",
            fontName="ArialBlack",
            fontSize=10,
            alignment=TA_CENTER,
            textColor=HexColor("#000000")
        )


        elementos.append(Paragraph(f"SISTEMA DE EMISSÃO DE RECIBOS", estilo_titulo_recibo))
        elementos.append(Paragraph(f"Seção Técnica de Finanças - FCAV/Unesp", estilo_titulo_recibo))
        elementos.append(Paragraph(f"Relatório - Recibos Gerados", estilo_titulo_recibo))
        elementos.append(Paragraph("-", estilo_informacoes_recibo))
        elementos.append(Paragraph("Filtros aplicados:", estilo_informacoes_recibo_negrito))
        elementos.append(Paragraph("-", estilo_informacoes_recibo))
        elementos.append(Paragraph(f"Interessado: {nome_pessoa_relatorio}", estilo_informacoes_recibo))
        elementos.append(Paragraph(f"Produto: {descricao_produto_relatorio}", estilo_informacoes_recibo))
        elementos.append(Paragraph("-", estilo_informacoes_recibo))
        elementos.append(Paragraph(f"Data Inicial: {datetime.strptime(data_inicial_relatorio, "%Y-%m-%d").strftime("%d/%m/%Y")} - Data Final: {datetime.strptime(data_final_relatorio, "%Y-%m-%d").strftime("%d/%m/%Y")}", estilo_informacoes_recibo))
        elementos.append(Spacer(1, 12))

        def formatar_data(data_str):
            if not data_str or str(data_str).lower() in ["nat", "none", ""]:
                return ""
            try:
                return datetime.strptime(str(data_str), "%Y-%m-%d").strftime("%d/%m/%Y")
            except ValueError:
                return str(data_str)

        # Cabeçalho da tabela
        dados_tabela = [['Interessado', 'Recibo', 'Data', 'Produto', 'Total', 'Obs.']] + [
            [Paragraph(str(nome_pessoa), style_normal),
            num_recibo,
            formatar_data(data),
            Paragraph(str(descricao_produto), style_normal), 
            f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), 
            Paragraph(str(observacao or ""), style_normal)] for nome_pessoa, num_recibo, data, descricao_produto, valor_total, observacao in dados
        ]

        # Tabela
        tabela = Table(dados_tabela, colWidths=[120, 45, 55, 110, 75, 140])
        tabela.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Arial'),
            ('FONTNAME', (4, 1), (4, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),

            # Alinhamentos
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),   # Interessado
            ('ALIGN', (1, 1), (1, -1), 'CENTER'), # Recibo
            ('ALIGN', (2, 1), (2, -1), 'CENTER'), # Data
            ('ALIGN', (3, 1), (3, -1), 'LEFT'),   # Produto
            ('ALIGN', (4, 1), (6, -1), 'RIGHT'),  # Valor, Qtd, Total
            ('ALIGN', (7, 1), (7, -1), 'LEFT'),   # Observação

            ('ALIGN', (0, 0), (-1, 0), 'CENTER')  # Cabeçalho
        ]))

        elementos.append(tabela)
        doc.build(elementos)

        # 3. Abrir o PDF automaticamente
        webbrowser.open_new(pdf_path)


    btn_gerar = tk.Button(
        janela_relatorio,
        text="Gerar Relatório",
        bg=cor_botao,
        fg=cor_texto_botao,
        font=("Segoe UI", 10, "bold"),
        activebackground=cor_botao_hover,
        activeforeground=cor_texto_botao,
        relief="flat",
        command=gerar_pdf
    )
    btn_gerar.grid(row=4, column=0, columnspan=5, pady=(20, 10), ipadx=10, ipady=5)

    # Efeito hover no botão
    def on_enter(e):
        btn_gerar.config(bg=cor_botao_hover)
    def on_leave(e):
        btn_gerar.config(bg=cor_botao)

    btn_gerar.bind("<Enter>", on_enter)
    btn_gerar.bind("<Leave>", on_leave)

def atualizar_produtos():

    # Janela principal
    janela_atprodutos = tk.Toplevel()
    janela_atprodutos.title("Atualizar Produtos")
    janela_atprodutos.geometry("520x300")
    janela_atprodutos.configure(bg="#F7F7F7")  # Fundo claro
    janela_atprodutos.resizable(False, False)

    def ajustar_produtos():
        #msg = tk.Label(janela_atprodutos, text="Aguarde a atualização da tabela de produtos...", fg="red", font=("Segoe UI", 12, "bold"))
        #msg.grid(row=1, column=0, pady=10)

        # Agendar para remover a mensagem após 10 segundos
        #janela_atprodutos.after(10000, msg.destroy)

        # Ajustar tabela
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
        #cursor.execute("DELETE FROM produtos")
        try:
            cursor.execute("DROP TABLE produtos")
        except:
            pass

        cursor.execute('''
        CREATE TABLE IF NOT EXISTS produtos (
            codigo_produto INTEGER PRIMARY KEY AUTOINCREMENT,
            descricao_produto TEXT UNIQUE,
            valor_unitario REAL,
            codigo_receita TEXT,
            rubrica TEXT,
            tipo_receita TEXT
        )''')

        for _, row in produtos.iterrows():
            try:
                cursor.execute("INSERT INTO produtos (codigo_produto, descricao_produto, valor_unitario, codigo_receita, rubrica, tipo_receita) VALUES (?, ?, ?, ?, ?, ?)", (row['codigo_produto'], row['descricao_produto'], row['valor_unitario'], row['codigo_receita'], row['rubrica'], row['tipo_receita']))
            except sqlite3.IntegrityError:
                pass  

        conn.commit()
        conn.close()

    # Fonte e cores padrão
    titulo_font = ("Segoe UI", 16, "bold")
    label_font = ("Segoe UI", 11)
    cor_titulo = "#000000"
    cor_texto = "#333333"

    # Título
    tk.Label(janela_atprodutos, text="Atualizar Produtos", font=titulo_font, fg=cor_titulo, bg="#F7F7F7").pack(pady=(20, 10))

    # Instruções em um frame
    frame_instrucao = tk.Frame(janela_atprodutos, bg="#F7F7F7")
    frame_instrucao.pack(pady=(0, 20), padx=20, fill="x")

    instrucoes = [
        "1. Acesse o Drive FINANÇAS > SER > tabelas > produtos",
        "2. Faça as alterações necessárias nos campos",
        "3. Espere as alterações serem salvas na planilha",
        "4. Clique no botão 'Atualizar' abaixo"
    ]

    for instrucao in instrucoes:
        tk.Label(frame_instrucao, text=instrucao, font=label_font, fg=cor_texto, bg="#F7F7F7", anchor="w", justify="left").pack(fill="x", pady=2)

    # Espaço antes do botão
    tk.Label(janela_atprodutos, text=" ", bg="#F7F7F7").pack()

    # Botão centralizado
    tk.Button(janela_atprodutos, text="Atualizar", font=("Segoe UI", 12, "bold"), bg="#99ADA8", fg="white", activebackground="#99ADA8",
              activeforeground="white", command=ajustar_produtos, padx=20, pady=5).pack(pady=10)

def atualizar_responsaveis():
    janela_atresponsaveis = tk.Toplevel()
    janela_atresponsaveis.title("Atualizar Responsáveis")
    janela_atresponsaveis.geometry("520x300")
    janela_atresponsaveis.configure(bg="#F7F7F7")  # Fundo claro
    janela_atresponsaveis.resizable(False, False)

    def ajustar_responsaveis():
        #msg = tk.Label(janela_atresponsaveis, text="Aguarde a atualização da tabela de responsáveis...", fg="red", font=("Segoe UI", 12, "bold"))
        #msg.grid(row=1, column=0, pady=10)

        # Agendar para remover a mensagem após 10 segundos
        #janela_atresponsaveis.after(10000, msg.destroy)

        #Atualizar tabela
        conn = sqlite3.connect(caminho_banco)
        cursor = conn.cursor()

        #Planilha - Produtos
        sheet_id = "1gQWvjnjXnFbbqIviV3Uq16H_oVnqdH4-BL6f8w0NRDQ"
        gid = "0"  # geralmente 0 para a primeira aba

        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        responsaveis = pd.read_csv(url)

        # Opção mais simples - DELETAR DADOS DO BANCO e apenas inserí-los
        #cursor.execute("DELETE FROM produtos")
        try:
            cursor.execute("DROP TABLE responsaveis")
        except:
            pass

        cursor.execute('''
        CREATE TABLE IF NOT EXISTS responsaveis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_responsavel TEXT
        )''')

        for _, row in responsaveis.iterrows():
            try:
                cursor.execute("INSERT INTO responsaveis (id, nome_responsavel) VALUES (?, ?)", (row['id'], row['nome_responsavel']))
            except sqlite3.IntegrityError:
                pass

        conn.commit()
        conn.close()

    # Fonte e cores padrão
    titulo_font = ("Segoe UI", 16, "bold")
    label_font = ("Segoe UI", 11)
    cor_titulo = "#000000"
    cor_texto = "#333333"

    # Título
    tk.Label(janela_atresponsaveis, text="Atualizar Responsáveis", font=titulo_font, fg=cor_titulo, bg="#F7F7F7").pack(pady=(20, 10))

    # Instruções em um frame
    frame_instrucao = tk.Frame(janela_atresponsaveis, bg="#F7F7F7")
    frame_instrucao.pack(pady=(0, 20), padx=20, fill="x")

    instrucoes = [
        "1. Acesse o Drive FINANÇAS > SER > tabelas > responsaveis",
        "2. Faça as alterações necessárias no campo (nome_responsavel)",
        "3. Espere as alterações serem salvas na planilha",
        "4. Clique no botão 'Atualizar' abaixo."
    ]

    for instrucao in instrucoes:
        tk.Label(frame_instrucao, text=instrucao, font=label_font, fg=cor_texto, bg="#F7F7F7", anchor="w", justify="left").pack(fill="x", pady=2)

    # Espaço antes do botão
    tk.Label(janela_atresponsaveis, text=" ", bg="#F7F7F7").pack()

    # Botão centralizado
    tk.Button(janela_atresponsaveis, text="Atualizar", font=("Segoe UI", 12, "bold"), bg="#99ADA8", fg="white", activebackground="#99ADA8",
              activeforeground="white", command=ajustar_responsaveis, padx=20, pady=5).pack(pady=10)

def comando_vazio():
    return None

def menu_principal():
    janela_menu = tk.Tk()
    janela_menu.title("Sistema de Emissão de Recibos")
    janela_menu.geometry("640x560")
    janela_menu.configure(bg="#F7F7F7") #"#F7F7F7"
    janela_menu.resizable(False, False)

    def encerrar():
        print("Encerrando o sistema.")
        janela_menu.quit()
        janela_menu.destroy()

    janela_menu.protocol("WM_DELETE_WINDOW", encerrar)

    # Fonts customizadas
    titulo_font = font.Font(family="Segoe UI", size=17, weight="bold")
    subtitulo_font = font.Font(family="Segoe UI", size=11)
    botao_font = font.Font(family="Segoe UI", size=10, weight="bold")

    # Título
    tk.Label(
        janela_menu,
        text="Sistema de Emissão de Recibos (SER)",
        font=titulo_font,
        fg="#000000",
        bg="#F7F7F7"
    ).pack(pady=(20, 5))

    tk.Label(
        janela_menu,
        text="Faculdade de Ciências Agrárias e Veterinárias - FCAV/Unesp - Câmpus de Jaboticabal",
        font=subtitulo_font,
        fg="#000000",
        bg="#F7F7F7"
    ).pack(pady=(0, 20))

    # Frame central para botões
    frame_botoes = tk.Frame(janela_menu, bg="#F7F7F7")
    frame_botoes.pack(pady=10)

    # Função para criar botão moderno
    def criar_botao_moderno(texto, comando, cor="#2F8AC6"):
        canvas = tk.Canvas(frame_botoes, width=220, height=50, bg="#F7F7F7", highlightthickness=0)
        canvas.pack(pady=10)

        # Desenhar retângulo com cantos arredondados (simples)
        x0, y0, x1, y1 = 2, 2, 218, 48
        canvas.create_rectangle(x0, y0, x1, y1, fill=cor, outline=cor, width=2)

        # Texto do botão
        canvas.create_text(110, 25, text=texto, fill="white", font=botao_font)

        # Função de clique
        def on_click(event):
            comando()

        canvas.bind("<Button-1>", on_click)
        return canvas

    # Botões principais
    criar_botao_moderno("Gerar Recibo", gerar_recibo)
    criar_botao_moderno("Gerar Relatório", gerar_relatorio)
    criar_botao_moderno("Atualizar Produtos", atualizar_produtos, cor="#99ADA8")
    criar_botao_moderno("Atualizar Responsáveis", atualizar_responsaveis, cor="#99ADA8")
    criar_botao_moderno("", comando_vazio, cor="#F7F7F7")
    criar_botao_moderno("Sair", encerrar, cor="#E06777")

    janela_menu.mainloop()

#Executar o código:
menu_principal()