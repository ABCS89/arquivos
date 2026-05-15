import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import calendar
import os

# =========================
# FUNÇÕES DE DATA
# =========================
def ultimo_dia_util_mes(ano, mes):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data = datetime(ano, mes, ultimo_dia)

    while data.weekday() >= 5:  # 5 = sábado, 6 = domingo
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
df_dividas = pd.read_excel("../template/devedores.xlsx", sheet_name="Inadimplentes")

# =========================
# DATA ATUAL
# =========================
hoje = datetime.now()
dia = hoje.day
mes = hoje.month
ano = hoje.year

mes_extenso = mes_por_extenso(mes)
ultimo_dia = ultimo_dia_util_mes(ano, mes)

# =========================
# TEMPLATE
# =========================
template = DocxTemplate("../template/template_aviso_cancelamento.docx")

# =========================
# LOOP PRINCIPAL
# =========================
for _, row in df_base.iterrows():

    if str(row.get("condição", "")).lower() != "não enviar":
        continue

    nome = row.get("nome")
    matricula = row.get("Nro Funcional")  # ajuste se necessário

    # filtrar dívidas do funcionário
    df_func = df_dividas[df_dividas["Funcional"] == matricula]

    if df_func.empty:
        continue

    tabela = []

    for _, d in df_func.iterrows():

        principal = float(d["Principal (Saldo)"])
        total = float(d["Saldo (Atualizado)"])
        encargos = total - principal

        tabela.append({
            "competencia": d["Mês/Ano"],
            "vencimento": pd.to_datetime(d["Data de Vencimento"]).strftime("%d/%m/%Y"),
            "principal": f"R$ {principal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "encargos": f"R$ {encargos:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "total": f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        })

    contexto = {
        "dia": dia,
        "mes": mes_extenso,
        "ano": ano,
        "ultimo_dia_util": ultimo_dia,
        "nome_cap": str(nome).upper(),
        "nome_upper": str(nome).upper(),
        "endereco": row.get("endereco", ""),
        "cep": row.get("cep", ""),
        "cidade": row.get("cidade", ""),
        "uf": row.get("uf", ""),
        "tabela": tabela
    }

    template.render(contexto)

    nome_arquivo = f"{nome} - aviso de cancelamento.docx"
    template.save(nome_arquivo)

print("Arquivos gerados com sucesso!")