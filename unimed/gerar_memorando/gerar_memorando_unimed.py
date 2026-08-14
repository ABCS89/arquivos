"""
gerar_memorando_unimed.py

Gera o memorando único (uma lista com todos os servidores) do Plano de
Saúde. O template_memorando.docx usa um loop Jinja ({% for pessoa in
pessoas %}) pra listar cada pessoa, mais dois campos de cabeçalho:

    {{data_atual}}       -> data de hoje, por extenso ("1 de agosto de 2026")
    {{data_vencimento}}  -> último dia útil do mês, no formato DD/MM/AAAA,
                             pulando fins de semana E feriados nacionais/SP.

⚠️ Limitação do cálculo de dia útil: a biblioteca "holidays" cobre feriados
NACIONAIS e ESTADUAIS (SP), mas não sabe de feriados/pontos facultativos
MUNICIPAIS de Piracicaba (ex: aniversário da cidade). Se o último dia útil
"oficial" do mês cair num feriado municipal, o script não vai saber e vai
usar esse dia mesmo assim — vale uma conferência visual nesses meses.
"""

from datetime import datetime, timedelta
from pathlib import Path

import holidays
import pandas as pd
from docxtpl import DocxTemplate

# =========================================================================
# CAMINHOS
# =========================================================================

BASE_DIR = Path(__file__).resolve().parent.parent
PASTA_TEMPLATE = BASE_DIR / "template"
PASTA_SAIDA = BASE_DIR / "output" / "memorando"

ARQUIVO_BASE = PASTA_TEMPLATE / "teste.ods"
TEMPLATE_MEMORANDO = PASTA_TEMPLATE / "template_memorando.docx"

# =========================================================================
# CONSTANTES
# =========================================================================

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

FERIADOS_SP = holidays.Brazil(subdiv="SP")


# =========================================================================
# FUNÇÕES AUXILIARES
# =========================================================================

def data_por_extenso(data):
    return f"{data.day} de {MESES_PT[data.month]} de {data.year}"


def ultimo_dia_util_do_mes(ano, mes):
    """
    Último dia do mês, recuando enquanto cair em sábado, domingo ou
    feriado nacional/estadual (SP). Retorna um objeto date.
    """
    if mes == 12:
        primeiro_dia_prox_mes = datetime(ano + 1, 1, 1)
    else:
        primeiro_dia_prox_mes = datetime(ano, mes + 1, 1)

    data = primeiro_dia_prox_mes - timedelta(days=1)

    while data.weekday() >= 5 or data.date() in FERIADOS_SP:
        data -= timedelta(days=1)

    return data


# =========================================================================
# PROGRAMA PRINCIPAL
# =========================================================================

def main():
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    hoje = datetime.now()
    data_atual = data_por_extenso(hoje)

    vencimento = ultimo_dia_util_do_mes(hoje.year, hoje.month)
    data_vencimento = vencimento.strftime("%d/%m/%Y")

    print(f"Data atual (cabeçalho): {data_atual}")
    print(f"Data de vencimento das guias: {data_vencimento}")

    df = pd.read_excel(ARQUIVO_BASE, engine="odf")

    pessoas = []
    for i, row in df.iterrows():
        valores = []
        if float(row["Mensalidade"]) != 0:
            valores.append(f"Mensalidade: R$ {row['Mensalidade']:.2f}")
        if float(row["Coparticipação"]) != 0:
            valores.append(f"Coparticipação: R$ {row['Coparticipação']:.2f}")

        pessoas.append({
            "guia": i + 1,
            "nome": str(row["Funcionário"]).title(),
            "cpf": row["cpf"],
            "data_nascimento": pd.to_datetime(row["data_nascimento"]).strftime("%d/%m/%Y"),
            "endereco": str(row["endereço"]).title(),
            "bairro": str(row["bairro"]).title(),
            "cidade": str(row["cidade"]).title(),
            "uf": "SP",
            "valores": "\n".join(valores),
            "total": f"R$ {float(row['Total']):.2f}",
        })

    doc = DocxTemplate(TEMPLATE_MEMORANDO)
    doc.render({
        "pessoas": pessoas,
        "data_atual": data_atual,
        "data_vencimento": data_vencimento,
    })
    doc.save(PASTA_SAIDA / "memorando_final.docx")

    print(f"\n✅ Memorando gerado com {len(pessoas)} pessoas.")


if __name__ == "__main__":
    main()
