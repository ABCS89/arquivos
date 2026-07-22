import pandas as pd
from docx import Document
import os
from datetime import datetime
from num2words import num2words
import re

# ========= CONFIGURAÇÕES ========= #

arquivo_excel = os.path.join('../template/devedores.xlsx')
template_path = os.path.join('../template/template_refis.docx')
pasta_saida = os.path.join('../output/refis')

os.makedirs(pasta_saida, exist_ok=True)

# ========= FUNÇÕES ========= #

def limpar_nome_arquivo(nome):
    """Remove caracteres inválidos para nome de arquivo"""
    nome = str(nome).strip()
    nome = re.sub(r'[\\/*?:"<>|]', '', nome)
    return nome

def formatar_valor(valor):
    """Formata valor para padrão brasileiro"""
    return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def valor_por_extenso(valor):
    """Gera valor por extenso correto (sem maiúsculo)"""
    valor_float = float(valor)
    reais = int(valor_float)
    centavos = int(round((valor_float - reais) * 100))

    if centavos > 0:
        extenso = f"{num2words(reais, lang='pt_BR')} reais e {num2words(centavos, lang='pt_BR')} centavos"
    else:
        extenso = f"{num2words(reais, lang='pt_BR')} reais"

    return extenso

def substituir_texto(doc, substituicoes):
    """Substitui placeholders no Word (inclusive dentro de runs)"""
    for p in doc.paragraphs:
        for chave, valor in substituicoes.items():
            if chave in p.text:
                for run in p.runs:
                    run.text = run.text.replace(chave, str(valor))

# ========= DATA ATUAL ========= #

hoje = datetime.now()

dia = hoje.day
mes = hoje.strftime('%B')  # nome do mês
ano = hoje.year

# corrigir mês para português manualmente
meses_pt = {
    "January": "janeiro", "February": "fevereiro", "March": "março",
    "April": "abril", "May": "maio", "June": "junho",
    "July": "julho", "August": "agosto", "September": "setembro",
    "October": "outubro", "November": "novembro", "December": "dezembro"
}
mes = meses_pt.get(mes, mes)

# ========= LEITURA DO EXCEL ========= #

df = pd.read_excel(arquivo_excel, sheet_name='Desligados')

# ========= LOOP PRINCIPAL ========= #

for _, row in df.iterrows():

    nome = row['Nome']

    # 🔹 Nome formatado
    nome_cap = str(nome).title()
    nome_upper = str(nome).upper()

    # 🔹 Nome do arquivo (bonito)
    nome_arquivo = limpar_nome_arquivo(nome_cap)

    # 🔹 Valores
    valor = row['Saldo (Atualizado)']
    valor_formatado = formatar_valor(valor)
    valor_extenso = valor_por_extenso(valor)

    # 🔹 Referência (ex: 07/2026)
    referencia = row['Mês/Ano']

    # 🔹 Data limite (se quiser usar vencimento)
    data_limite = row.get('Data de Vencimento', '')

    # ========= ABRIR TEMPLATE ========= #

    doc = Document(template_path)

    # ========= SUBSTITUIÇÕES ========= #

    substituicoes = {
        '{{nome}}': nome_cap,
        '{{nome_upper}}': nome_upper,
        '{{valor}}': valor_formatado,
        '{{valor_extenso}}': valor_extenso,
        '{{referencia}}': referencia,
        '{{data_limite}}': data_limite,
        '{{dia}}': dia,
        '{{mes}}': mes,
        '{{ano}}': ano,
    }

    substituir_texto(doc, substituicoes)

    # ========= SALVAR ========= #

    caminho_saida = os.path.join(pasta_saida, f"{nome_arquivo}.docx")
    doc.save(caminho_saida)

print("✅ Cartas geradas com sucesso!")