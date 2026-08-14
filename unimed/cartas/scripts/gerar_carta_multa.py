"""
gerar_carta_multa.py

Gera uma carta específica (baseada no template_base, com um parágrafo a
mais) para servidores que:
  1. Recebem a carta "base" normal (condição em branco na teste.ods); e
  2. Têm, na aba "Inadimplentes" da devedores.xlsx, uma parcela com
     vencimento no MÊS ANTERIOR ao atual, cujo Saldo (Atualizado) é menor
     que uma mensalidade cheia (R$ 333,49 fixo) — ou seja, a mensalidade
     já foi paga, mas a multa de atraso ainda está em aberto.

Essa carta é um documento À PARTE da carta base mensal normal (não
substitui — a pessoa pode receber as duas).

Estrutura de pastas (mesma raiz dos outros scripts):

    <raiz>/
      scripts/gerar_carta_multa.py
      template/teste.ods
      template/devedores.xlsx
      template/template_base_multa.docx
      pdfs/                          <- mesmos PDFs de comprovação de e-mail
      output/multa/
"""

import calendar
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import PyPDF2
from docxtpl import DocxTemplate
from num2words import num2words

# =========================================================================
# CAMINHOS
# =========================================================================

BASE_DIR = Path(__file__).resolve().parent.parent

PASTA_TEMPLATE = BASE_DIR / "template"
PASTA_PDFS = BASE_DIR / "pdfs"
PASTA_SAIDA = BASE_DIR / "output" / "multa"

ARQUIVO_BASE = PASTA_TEMPLATE / "teste.ods"
ARQUIVO_DEVEDORES = PASTA_TEMPLATE / "devedores.xlsx"
TEMPLATE_MULTA = PASTA_TEMPLATE / "template_base_multa.docx"

# =========================================================================
# CONSTANTES
# =========================================================================

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

VALOR_UMA_MENSALIDADE = 333.49  # fixo, conforme definido


# =========================================================================
# FUNÇÕES AUXILIARES (mesmas do gerar_cartas_mensais.py)
# =========================================================================

def limpa(valor):
    return "" if pd.isna(valor) else str(valor).strip()


def normalizar_nome(nome):
    nome = str(nome).lower()
    substituicoes = {
        "[áàãâä]": "a", "[éèêë]": "e", "[íìîï]": "i",
        "[óòõôö]": "o", "[úùûü]": "u", "[ç]": "c",
    }
    for padrao, letra in substituicoes.items():
        nome = re.sub(padrao, letra, nome)
    return re.sub(r"[^a-z0-9]", "", nome)


def capitalizar_nome(nome):
    return " ".join(palavra.capitalize() for palavra in str(nome).lower().split())


def formatar_valor_br(valor):
    return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def valor_por_extenso(valor):
    valor_float = float(valor)
    reais = int(valor_float)
    centavos = int(round((valor_float - reais) * 100))
    texto = f"{num2words(reais, lang='pt_BR')} reais"
    if centavos > 0:
        texto += f" e {num2words(centavos, lang='pt_BR')} centavos"
    return texto


def extrair_texto_pdf(caminho_pdf):
    texto = ""
    try:
        with open(caminho_pdf, "rb") as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)
            for pagina in leitor.pages:
                texto += pagina.extract_text() or ""
    except Exception as erro:
        print(f"  ⚠️ Erro ao ler PDF {caminho_pdf}: {erro}")
    return texto


def extrair_data_email_do_pdf(texto_pdf):
    match = re.search(r"Data\s+(\d{2}/\d{2}/\d{4})", texto_pdf)
    if not match:
        return {"dia_email": "dia", "mes_email": "mês", "ano_email": "ano"}
    data_obj = datetime.strptime(match.group(1), "%d/%m/%Y")
    return {
        "dia_email": data_obj.day,
        "mes_email": MESES_PT[data_obj.month],
        "ano_email": data_obj.year,
    }


def mapear_pdfs_por_funcional(pasta_pdfs, df_base):
    mapa = {}
    if not pasta_pdfs.exists():
        print(f"⚠️ Pasta de PDFs não encontrada: {pasta_pdfs}")
        return mapa
    arquivos_pdf = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith(".pdf")]
    for nome_arquivo in arquivos_pdf:
        nome_extraido = os.path.splitext(nome_arquivo)[0].split(" - ")[0]
        nome_normalizado = normalizar_nome(nome_extraido)
        for _, linha in df_base.iterrows():
            nome_planilha = normalizar_nome(linha["Funcionário"])
            if nome_normalizado in nome_planilha or nome_planilha in nome_normalizado:
                mapa[linha["Nro Funcional"]] = nome_arquivo
                break
    return mapa


def mes_anterior(hoje):
    """Devolve (ano, mes) do mês anterior ao de 'hoje'."""
    if hoje.month == 1:
        return hoje.year - 1, 12
    return hoje.year, hoje.month - 1


# =========================================================================
# PROGRAMA PRINCIPAL
# =========================================================================

def main():
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    hoje = datetime.now()
    contexto_data_atual = {
        "dia": hoje.day,
        "mes": MESES_PT[hoje.month],
        "ano": hoje.year,
    }

    ano_anterior, mes_anterior_num = mes_anterior(hoje)
    print(f"Verificando parcelas vencidas em {MESES_PT[mes_anterior_num]}/{ano_anterior}...")

    last_day = calendar.monthrange(hoje.year, hoje.month)[1]
    due_date = datetime(hoje.year, hoje.month, last_day)
    while due_date.weekday() >= 5:
        due_date = due_date.replace(day=due_date.day - 1)
    contexto_vencimento = {
        "ultimo_dia_util": due_date.day,
        "mes_vencimento": MESES_PT[hoje.month],
        "ano_vencimento": hoje.year,
    }

    # --- dados base ---
    df_base = pd.read_excel(ARQUIVO_BASE, engine="odf")
    df_base.columns = df_base.columns.str.strip()
    df_base["Nro Funcional"] = pd.to_numeric(df_base["Nro Funcional"], errors="coerce").astype("Int64")

    # só quem recebe a carta base normal (condição em branco)
    df_base_normal = df_base[
        df_base["condição"].isna() | (df_base["condição"].astype(str).str.strip() == "")
    ]

    # --- dívidas ---
    df_inadimplentes = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Inadimplentes")
    df_inadimplentes.columns = df_inadimplentes.columns.str.strip()
    df_inadimplentes["Funcional"] = pd.to_numeric(df_inadimplentes["Funcional"], errors="coerce").astype("Int64")

    # filtra: vencimento no mês anterior + saldo menor que 1 mensalidade
    candidatos = df_inadimplentes[
        (df_inadimplentes["Data de Vencimento"].dt.month == mes_anterior_num)
        & (df_inadimplentes["Data de Vencimento"].dt.year == ano_anterior)
        & (df_inadimplentes["Saldo (Atualizado)"] < VALOR_UMA_MENSALIDADE)
    ]

    if candidatos.empty:
        print("Nenhuma pessoa se enquadra no critério esse mês.")
        return

    print(f"Mapeando PDFs de comprovação de e-mail...")
    pdf_map = mapear_pdfs_por_funcional(PASTA_PDFS, df_base)

    gerados = 0
    for _, divida in candidatos.iterrows():
        funcional = int(divida["Funcional"])
        linha_pessoa = df_base_normal[df_base_normal["Nro Funcional"] == funcional]

        if linha_pessoa.empty:
            print(f"  ⚠️ {divida['Nome']} (Funcional {funcional}) tem multa em aberto, "
                  f"mas não está na lista de carta base normal (condição diferente ou não encontrada) — pulando.")
            continue

        linha = linha_pessoa.iloc[0]
        funcionario_raw = linha["Funcionário"]

        if funcional in pdf_map:
            texto_pdf = extrair_texto_pdf(PASTA_PDFS / pdf_map[funcional])
            contexto_email = extrair_data_email_do_pdf(texto_pdf)
        else:
            print(f"  ⚠️ Sem PDF de comprovação para: {funcionario_raw}")
            contexto_email = {"dia_email": "dia", "mes_email": "mês", "ano_email": "ano"}

        endereco_completo = limpa(linha.get("endereço"))
        complemento = limpa(linha.get("complemento"))
        bairro = limpa(linha.get("bairro"))
        if complemento:
            endereco_completo += f" – {complemento}"
        if bairro:
            endereco_completo += f" – {bairro}"

        valor_multa = float(divida["Saldo (Atualizado)"])

        contexto = {
            "nome_cap": capitalizar_nome(funcionario_raw),
            "nome_upper": str(funcionario_raw).upper(),
            "linha_endereco": endereco_completo,
            "CEP": limpa(linha.get("CEP")),
            "cidade": limpa(linha.get("cidade")),
            "valor": formatar_valor_br(linha["Total"]),
            "valor_extenso": valor_por_extenso(linha["Total"]),
            "valor_multa": formatar_valor_br(valor_multa),
            "valor_multa_extenso": valor_por_extenso(valor_multa),
            "email": limpa(linha.get("mail")) or "mail",
            **contexto_data_atual,
            **contexto_vencimento,
            **contexto_email,
        }

        doc = DocxTemplate(TEMPLATE_MULTA)
        doc.render(contexto)
        doc.save(PASTA_SAIDA / f"{funcionario_raw}.docx")
        gerados += 1
        print(f"  ✅ {funcionario_raw} — multa de R$ {formatar_valor_br(valor_multa)}")

    print(f"\n✅ Concluído: {gerados} carta(s) de aviso de multa gerada(s).")


if __name__ == "__main__":
    main()
