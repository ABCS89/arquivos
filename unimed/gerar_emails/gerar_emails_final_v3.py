import pandas as pd
from datetime import datetime, timedelta
import calendar
import os

# -------------------------
# FUNÇÕES DE DATA
# -------------------------

def ultimo_dia_util_mes(ano, mes):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data = datetime(ano, mes, ultimo_dia)

    while data.weekday() >= 5:
        data -= timedelta(days=1)

    return data.strftime("%d/%m/%Y")


def mes_referencia():
    hoje = datetime.today()
    primeiro_dia_mes = hoje.replace(day=1)
    mes_anterior = primeiro_dia_mes - timedelta(days=1)

    meses = [
        "janeiro","fevereiro","março","abril","maio","junho",
        "julho","agosto","setembro","outubro","novembro","dezembro"
    ]

    nome_mes = meses[mes_anterior.month - 1]
    return f"{nome_mes}/{mes_anterior.year}"


def formata_competencia(valor):
    if pd.isna(valor):
        return ""

    try:
        data = pd.to_datetime(valor, errors="coerce")

        if pd.isna(data):
            return str(valor)

        meses = ["Jan","Fev","Mar","Abr","Mai","Jun",
                 "Jul","Ago","Set","Out","Nov","Dez"]

        return f"{meses[data.month - 1]}/{str(data.year)[-2:]}"

    except:
        return str(valor)


# -------------------------
# TABELA MARKDOWN
# -------------------------

def gerar_tabela_markdown(tabela):
    if not tabela:
        return "Sem débitos."

    linhas = []
    linhas.append("| Competência | Vencimento | Principal | Encargos | Total |")
    linhas.append("|-------------|------------|-----------|----------|-------|")

    for item in tabela:
        linhas.append(
            f"| {item['competencia']} | {item['vencimento']} | {item['principal']} | {item['encargos']} | {item['total']} |"
        )

    return "\n".join(linhas)


# -------------------------
# CAMINHOS
# -------------------------

pasta = "../template"
arquivo_ods = os.path.join(pasta, "teste.ods")
arquivo_dividas = os.path.join(pasta, "devedores.xlsx")

# Templates
with open(os.path.join(pasta, "email_desligados.txt"), "r", encoding="utf-8") as f:
    template_desligado = f.read()

with open(os.path.join(pasta, "email_normal.txt"), "r", encoding="utf-8") as f:
    template_normal = f.read()

with open(os.path.join(pasta, "email_aviso.txt"), "r", encoding="utf-8") as f:
    template_aviso = f.read()

with open(os.path.join(pasta, "email_cancelado.txt"), "r", encoding="utf-8") as f:
    template_cancelado = f.read()


# -------------------------
# DADOS
# -------------------------

df = pd.read_excel(arquivo_ods, engine="odf")
df_dividas = pd.read_excel(arquivo_dividas, sheet_name="Inadimplentes")

# Padronizar chave
df["Nro Funcional"] = pd.to_numeric(df["Nro Funcional"], errors="coerce").astype("Int64").astype(str)
df_dividas["Funcional"] = pd.to_numeric(df_dividas["Funcional"], errors="coerce").astype("Int64").astype(str)

hoje = datetime.today()
data_envio = hoje.strftime("%d/%m/%Y")
data_vencimento = ultimo_dia_util_mes(hoje.year, hoje.month)
referencia = mes_referencia()


# -------------------------
# LISTAS DE SAÍDA
# -------------------------

emails_desligados = []
emails_normais = []
emails_aviso = []
emails_cancelados = []


# -------------------------
# PROCESSAMENTO
# -------------------------

for _, row in df.iterrows():
    condição = str(row.get("condição", "")).strip().lower()

    if condição == "desligado":
        tipo = "desligado"
    elif condição == "aviso":
        tipo = "aviso"
    elif condição == "cancelado":
        tipo = "cancelado"
    elif condição == "" or condição == "nan":
        tipo = "normal"
    elif "não enviar" in condição or "nao enviar" in condição:
        continue
    else:
        continue

    nome = row["Funcionário"]
    mail = row["mail"]
    total = row.get("Total", "")

    matricula = str(row.get("Nro Funcional")).strip()

    # -------------------------
    # GERAR TABELA (SÓ PARA AVISO)
    # -------------------------

    tabela_md = ""

    if tipo in ["aviso", "cancelado"]:
        df_func = df_dividas[df_dividas["Funcional"] == matricula]
        total_geral = pd.to_numeric(df_func["Saldo (Atualizado)"], errors="coerce").fillna(0).sum()
        valor_total_formatado = f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        tabela = []

        for _, d in df_func.iterrows():
            principal = float(d.get("Principal (Saldo)", 0) or 0)
            total_div = float(d.get("Saldo (Atualizado)", 0) or 0)
            encargos = total_div - principal

            data_venc = pd.to_datetime(d.get("Data de Vencimento"), dayfirst=True, errors="coerce")
            vencimento = data_venc.strftime("%d/%m/%Y") if pd.notna(data_venc) else ""

            tabela.append({
                "competencia": formata_competencia(d.get("Mês/Ano")),
                "vencimento": vencimento,
                "principal": f"R$ {principal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "encargos": f"R$ {encargos:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "total": f"R$ {total_div:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            })

        tabela_md = gerar_tabela_markdown(tabela)

    # -------------------------
    # TEMPLATE
    # -------------------------

    if tipo == "desligado":
        template = template_desligado
    elif tipo == "aviso":
        template = template_aviso
    elif tipo == "cancelado":
        template = template_cancelado
    else:
        template = template_normal

    # -------------------------
    # DEFINIR VALOR TOTAL
    # -------------------------

    if tipo in ["aviso", "cancelado"]:
        valor_total_final = valor_total_formatado
    else:
        valor_total_final = total

    # -------------------------
    # TEMPLATE
    # -------------------------

    corpo = template.format(
        nome=nome,
        valor_total=valor_total_final,
        valores=valor_total_final,
        data_final_mes=data_vencimento,
        referencia=referencia,
        data=data_envio,
        tabela=tabela_md
    )
    assunto = f"Boleto do Plano de Saúde Unimed – Referente a {referencia}"

    bloco = {
        "nome": nome,
        "email": mail,
        "assunto": assunto,
        "mensagem": corpo
    }

    if tipo == "desligado":
        emails_desligados.append(bloco)
    elif tipo == "aviso":
        emails_aviso.append(bloco)
    elif tipo == "cancelado":
        emails_cancelados.append(bloco)
    else:
        emails_normais.append(bloco)


# -------------------------
# SALVAR
# -------------------------

def salvar(nome_arquivo, titulo, lista):
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        f.write(f"# {titulo}\n\n")

        for email in lista:
            f.write("---\n\n")
            f.write(f"## 👤 {email['nome']}\n\n")
            f.write(f"> {email['email']}  \n")
            f.write(f"> {email['assunto']}  \n\n")

            f.write("### ✉️ Mensagem:\n\n")
            f.write(f"{email['mensagem']}\n\n")


salvar("emails_normais.md", "📧 Emails Normais", emails_normais)
salvar("emails_desligados.md", "📧 Emails de Desligados", emails_desligados)
salvar("emails_aviso.md", "📧 Emails de Aviso de Cancelamento", emails_aviso)
salvar("emails_cancelados.md", "📧 Emails de Cancelados", emails_cancelados)

print("Arquivos gerados com sucesso!")