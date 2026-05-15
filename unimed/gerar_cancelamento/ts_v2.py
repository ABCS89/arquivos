import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import calendar
import os

def formata_competencia(valor):
    if pd.isna(valor):
        return ""

    # se já for texto (ex: Mar/26)
    if isinstance(valor, str):
        return valor.strip()

    # se for data (Timestamp)
    try:
        data = pd.to_datetime(valor, errors="coerce")

        if pd.isna(data):
            return ""

        meses = [
            "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"
        ]

        mes = meses[data.month - 1]
        ano = str(data.year)[-2:]

        return f"{mes}/{ano}"

    except:
        return str(valor)

def limpa(valor):
    return "" if pd.isna(valor) else str(valor).strip()

# =========================
# FUNÇÕES DE DATA
# =========================
def dia_limite_pagamento(ano, mes):
    data = datetime(ano, mes, 25)

    # 5 = sábado | 6 = domingo
    while data.weekday() >= 5:
        data = data.replace(day=data.day - 1)

    return data.day

def ultimo_dia_util_mes(ano, mes):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data = datetime(ano, mes, ultimo_dia)

    while data.weekday() >= 5:
        data = data.replace(day=data.day - 1)

    return data.day

def mes_por_extenso(mes):
    meses = [
        "janeiro", "fevereiro", "março", "abril",
        "maio", "junho", "julho", "agosto",
        "setembro", "outubro", "novembro", "dezembro"
    ]
    return meses[mes - 1]

# =========================
# CARREGAR ARQUIVOS
# =========================
df_base = pd.read_excel("../template/teste.ods", engine="odf")
df_inadimplentes = pd.read_excel("../template/devedores.xlsx", sheet_name="Inadimplentes")
df_cancelados = pd.read_excel("../template/devedores.xlsx", sheet_name="Cancelados")

print("COLUNAS BASE:")
print(df_base.columns.tolist())

print("\nCOLUNAS DIVIDAS:")
print(df_dividas.columns.tolist())

# =========================
# LIMPEZA DE DADOS (🔥 IMPORTANTE)
# =========================

# limpar nomes das colunas
df_inadimplentes.columns = df_inadimplentes.columns.str.strip()
df_cancelados.columns = df_cancelados.columns.str.strip()

df_inadimplentes["Funcional"] = (
    pd.to_numeric(df_inadimplentes["Funcional"], errors="coerce")
    .astype("Int64")
    .astype(str)
)

df_cancelados["Funcional"] = (
    pd.to_numeric(df_cancelados["Funcional"], errors="coerce")
    .astype("Int64")
    .astype(str)
)


# padronizar chaves (funcional)
df_base["Nro Funcional"] = (
    pd.to_numeric(df_base["Nro Funcional"], errors="coerce")
    .astype("Int64")
    .astype(str)
)

df_dividas["Funcional"] = (
    pd.to_numeric(df_dividas["Funcional"], errors="coerce")
    .astype("Int64")
    .astype(str)
)

# =========================
# DATA ATUAL
# =========================
hoje = datetime.now()
dia = hoje.day
mes = hoje.month
ano = hoje.year

mes_extenso = mes_por_extenso(mes)
dia_limite = dia_limite_pagamento(ano, mes)


# =========================
# TEMPLATE
# =========================
template_aviso_path = "../template/template_aviso.docx"
template_cancelado_path = "../template/template_cancelado.docx"


# =========================
# LOOP PRINCIPAL
# =========================

condicao = str(row.get("condição", "")).lower().strip()

if condicao not in ["aviso", "cancelado"]:
    continue

if condicao == "aviso":
    df_dividas = df_inadimplentes
    template_path = template_aviso_path
    nome_saida = f"{nome} - aviso.docx"

elif condicao == "cancelado":
    df_dividas = df_cancelados
    template_path = template_cancelado_path
    nome_saida = f"{nome} - cancelado.docx"



for _, row in df_base.iterrows():

    if str(row.get("condição", "")).lower().strip() != "aviso":
        continue

    nome = row.get("Funcionário")
    if pd.isna(row.get("Nro Funcional")):
        continue

    matricula = str(row.get("Nro Funcional")).strip()
   

    # =========================
    # FILTRO CORRIGIDO
    # =========================

    print(f"\n🔎 Procurando funcional: {matricula}")

    print("BASE:", df_base["Nro Funcional"].unique()[:5])
    print("DIVIDAS:", df_dividas["Funcional"].unique()[:5])

    df_func = df_dividas[df_dividas["Funcional"] == matricula]

    if df_func.empty:
        print(f"⚠️ Sem débitos para: {nome} ({matricula})")
        continue

    endereco = limpa(row.get("endereço"))
    bairro = limpa(row.get("bairro"))
    complemento = limpa(row.get("complemento"))

    linha_endereco = " – ".join(
        [x for x in [endereco, bairro, complemento] if x]
    )

    tabela = []

    for _, d in df_func.iterrows():

        # tratar valores nulos
        principal = float(d.get("Principal (Saldo)", 0) or 0)
        total = float(d.get("Saldo (Atualizado)", 0) or 0)
        encargos = total - principal

        # tratar datas com segurança
        try:
            data_venc = pd.to_datetime(
                d.get("Data de Vencimento"),
                dayfirst=True,
                errors="coerce"
            )

            vencimento = data_venc.strftime("%d/%m/%Y") if pd.notna(data_venc) else ""

        except:
            vencimento = ""

        tabela.append({
            "competencia": formata_competencia(d.get("Mês/Ano")),
            "vencimento": vencimento,
            "principal": f"R$ {principal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "encargos": f"R$ {encargos:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "total": f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        })


    contexto = {
        "dia": dia,
        "mes": mes_extenso,
        "ano": ano,
        # "ultimo_dia_util": ultimo_dia,
        "dia_limite": dia_limite,
        "nome_cap": str(nome).upper(),
        "nome_upper": str(nome).upper(),
        "linha_endereco": linha_endereco,
        "CEP": limpa(row.get("CEP")),
        "cidade": limpa(row.get("cidade")),
        "uf": limpa(row.get("uf")),
        "tabela": tabela
        }

    # 🔥 recriar template a cada loop (evita sobrescrita bugada)
    template = DocxTemplate(template_path)
    template.render(contexto)
    template.save(nome_saida)

    nome_arquivo = f"{nome} - aviso de cancelamento.docx"
    template.save(nome_arquivo)

    # =========================
    # CARTA DE CANCELADO
    # =========================
    template_cancelado_path = "../template/template_cancelado.docx"

    template_cancelado = DocxTemplate(template_cancelado_path)
    template_cancelado.render(contexto)

    nome_cancelado = f"{nome} - cancelado.docx"
    template_cancelado.save(nome_cancelado)


print(d.get("Mês/Ano"), type(d.get("Mês/Ano")))
print("Arquivos gerados com sucesso!")