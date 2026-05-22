import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import calendar
import os

# =========================
# FUNÇÕES AUXILIARES
# =========================

output_dir = "../output/cancelados"
os.makedirs(output_dir, exist_ok=True)

def ultimo_dia_util_mes(ano, mes):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data = datetime(ano, mes, ultimo_dia)

    while data.weekday() >= 5:  # sábado ou domingo
        data = data.replace(day=data.day - 1)

    return data.day


def formata_competencia(valor):
    if pd.isna(valor):
        return ""

    if isinstance(valor, str):
        return valor.strip()

    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return ""

    meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
             "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

    return f"{meses[data.month - 1]}/{str(data.year)[-2:]}"


def limpa(valor):
    return "" if pd.isna(valor) else str(valor).strip()


# =========================
# FUNÇÕES DE DATA
# =========================

def dia_limite_pagamento(ano, mes):
    ultimo_dia_util = ultimo_dia_util_mes(ano, mes)
    data = datetime(ano, mes, 25)
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
# CARREGAR DADOS
# =========================
df_base = pd.read_excel("../template/teste.ods", engine="odf")

df_inadimplentes = pd.read_excel(
    "../template/devedores.xlsx", sheet_name="Inadimplentes"
)
df_cancelados = pd.read_excel(
    "../template/devedores.xlsx", sheet_name="Cancelados"
)

# =========================
# LIMPEZA
# =========================
df_inadimplentes.columns = df_inadimplentes.columns.str.strip()
df_cancelados.columns = df_cancelados.columns.str.strip()

df_base.columns = df_base.columns.str.strip()

# padronizar chave
df_base["Nro Funcional"] = (
    pd.to_numeric(df_base["Nro Funcional"], errors="coerce")
    .astype("Int64")
    .astype(str)
)

for df in [df_inadimplentes, df_cancelados]:
    df["Funcional"] = (
        pd.to_numeric(df["Funcional"], errors="coerce")
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
ultimo_dia_util = ultimo_dia_util_mes(ano, mes)
ultimo_dia_do_mes = calendar.monthrange(ano, mes)[1]

# =========================
# TEMPLATES
# =========================
template_aviso = "../template/template_aviso.docx"
template_cancelado = "../template/template_cancelado.docx"

# =========================
# LOOP PRINCIPAL
# =========================
for _, row in df_base.iterrows():

    condicao = str(row.get("condição", "")).lower().strip()

    if condicao not in ["aviso", "cancelado"]:
        continue

    nome = limpa(row.get("Funcionário"))

    if pd.isna(row.get("Nro Funcional")):
        continue

    matricula = str(row.get("Nro Funcional")).strip()

    # escolher base e template
    if condicao == "aviso":
        df_dividas = df_inadimplentes
        template_path = template_aviso
        nome_saida = f"{nome} - aviso.docx"
    else:
        df_dividas = df_cancelados
        template_path = template_cancelado
        nome_saida = f"{nome} - cancelado.docx"

    print(f"\n🔎 {nome} | {condicao} | {matricula}")

    df_func = df_dividas[df_dividas["Funcional"] == matricula]

    if df_func.empty:
        print(f"⚠️ Sem dados para: {nome}")
        continue

    # endereço
    linha_endereco = " – ".join(
        [x for x in [
            limpa(row.get("endereço")),
            limpa(row.get("bairro")),
            limpa(row.get("complemento"))
        ] if x]
    )

    tabela = []

    for _, d in df_func.iterrows():

        principal = float(d.get("Principal (Saldo)", 0) or 0)
        total = float(d.get("Saldo (Atualizado)", 0) or 0)
        encargos = total - principal

        data_venc = pd.to_datetime(
            d.get("Data de Vencimento"),
            dayfirst=True,
            errors="coerce"
        )

        vencimento = data_venc.strftime("%d/%m/%Y") if pd.notna(data_venc) else ""

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
        "dia_limite": dia_limite,
        "ultimo_dia_util": ultimo_dia_util,  # 🔥 ESSENCIAL
        "ultimo_dia_do_mes": ultimo_dia_do_mes,  # 🔥 NOVO
        "nome_upper": nome.upper(),
        "nome_cap": str(nome).title(),
        "linha_endereco": linha_endereco,
        "CEP": limpa(row.get("CEP")),
        "cidade": limpa(row.get("cidade")),
        "uf": limpa(row.get("uf")),
        "tabela": tabela
    }

    doc = DocxTemplate(template_path)
    doc.render(contexto)
    caminho_saida = os.path.join(output_dir, nome_saida)
    doc.save(caminho_saida)

print("\n✅ Arquivos gerados com sucesso!")