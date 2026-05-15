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

    # Se sábado ou domingo → volta até sexta
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


# -------------------------
# CAMINHOS
# -------------------------

pasta = "../template"
arquivo_ods = os.path.join(pasta, "teste.ods")

# Templates
with open(os.path.join(pasta, "email_desligados.txt"), "r", encoding="utf-8") as f:
    template_desligado = f.read()

with open(os.path.join(pasta, "email_normal.txt"), "r", encoding="utf-8") as f:
    template_normal = f.read()
with open(os.path.join(pasta, "email_aviso.txt"), "r", encoding="utf-8") as f:
    template_aviso = f.read()

with open(os.path.join(pasta, "email_cancelado.txt"), "r", encoding="utf-8") as f:
    template_cancelado = f.read()

hoje = datetime.today()
data_envio = hoje.strftime("%d/%m/%Y")

# -------------------------
# DADOS
# -------------------------

df = pd.read_excel(arquivo_ods, engine="odf")

hoje = datetime.today()
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
        continue        # qualquer coisa diferente também ignora

    Funcionário = row["Funcionário"]
    mail = row["mail"]
    Total = row.get("Total", "")

    # Escolher template
    if tipo == "desligado":
        template = template_desligado
    elif tipo == "aviso":
        template = template_aviso
    elif tipo == "cancelado":
        template = template_cancelado
    else:
        template = template_normal

    # Ajustar referência
    texto = template

    corpo = texto.format(
        nome=Funcionário,
        valor_total=Total,
        valores=Total,
        data_final_mes=data_vencimento,
        referencia=referencia,
        data=data_envio
    )

    assunto = f"Boleto do Plano de Saúde Unimed – Referente a {referencia}"

    bloco = {
    "nome": Funcionário,
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
# SALVAR ARQUIVOS
# -------------------------

separador = "\n-------------------------\n"

with open("emails_normais.md", "w", encoding="utf-8") as f:
    f.write("# 📧 Emails Normais\n\n")

    for email in emails_normais:
        f.write("---\n\n")
        f.write(f"## 👤 {email['nome']}\n\n")
        f.write(f"> {email['email']}  \n")
        f.write(f"> {email['assunto']}  \n")
        # f.write(f"> **Tipo:** Normal\n\n")

        f.write("### ✉️ Mensagem:\n\n")
        f.write(f"{email['mensagem']}\n\n")
        
with open("emails_desligados.md", "w", encoding="utf-8") as f:
    f.write("# 📧 Emails de Desligados\n\n")

    for email in emails_desligados:
        f.write("---\n\n")
        f.write(f"## 👤 {email['nome']}\n\n")
        f.write(f"> {email['email']}  \n")
        f.write(f"> {email['assunto']}  \n")
        # f.write(f"- Matrícula: {email['matricula']}\n")
        # f.write(f"- Tipo: Desligado\n\n")

        f.write("### ✉️ Conteúdo do Email:\n")
        f.write(f"{email['mensagem']}\n\n")

with open("emails_aviso.md", "w", encoding="utf-8") as f:
    f.write("# 📧 Emails de Aviso de Cancelamento\n\n")

    for email in emails_aviso:
        f.write("---\n\n")
        f.write(f"## 👤 {email['nome']}\n\n")
        f.write(f"> {email['email']}  \n")
        f.write(f"> {email['assunto']}  \n")

        f.write("### ✉️ Mensagem:\n\n")
        f.write(f"{email['mensagem']}\n\n")


with open("emails_cancelados.md", "w", encoding="utf-8") as f:
    f.write("# 📧 Emails de Cancelados\n\n")

    for email in emails_cancelados:
        f.write("---\n\n")
        f.write(f"## 👤 {email['nome']}\n\n")
        f.write(f"> {email['email']}  \n")
        f.write(f"> {email['assunto']}  \n")

        f.write("### ✉️ Mensagem:\n\n")
        f.write(f"{email['mensagem']}\n\n")

        
print("Arquivos gerados com sucesso!")