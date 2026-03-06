import pdfplumber
import pandas as pd
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta

MES_REFERENCIA = 2
ANO_REFERENCIA = 2026

def extrair_texto_pdf(caminho):
    texto = []
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)

# ---------------------------------------------------
# FREQUENCIA SECRETARIA
# ---------------------------------------------------

def parse_frequencia(texto):

    dados = []
    linhas = texto.split("\n")

    padrao = re.compile(r'(\d{2}\.\d{3}-\d)\s+(.+?)\s+(Férias regulamentares|Frequência normal)\s*(\d{2}/\d{2}/\d{4})?\s*([\d,]+)?', re.I)

    for linha in linhas:
        m = padrao.search(linha)

        if not m:
            continue

        funcional = m.group(1)
        nome = m.group(2).strip()
        ocorrencia = m.group(3).strip()

        data = m.group(4)
        qtd = m.group(5)

        if ocorrencia.lower() == "frequência normal":
            continue

        if data:
            data = datetime.strptime(data, "%d/%m/%Y")

        if qtd:
            qtd = float(qtd.replace(",", "."))

        dados.append({
            "funcional": funcional,
            "nome": nome,
            "ocorrencia": ocorrencia.upper(),
            "data": data,
            "qtd_secretaria": qtd
        })

    return pd.DataFrame(dados)

# ---------------------------------------------------
# RELATORIO RH
# ---------------------------------------------------

def parse_rh(texto):

    dados = []
    linhas = texto.split("\n")

    padrao = re.compile(
        r'(\d{2}\.\d{3}-\d)\s+(.+?)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(Férias Regulamentares)',
        re.I
    )

    for linha in linhas:

        m = padrao.search(linha)

        if not m:
            continue

        funcional = m.group(1)
        nome = m.group(2).strip()
        data_inicio = datetime.strptime(m.group(3), "%d/%m/%Y")
        data_fim = datetime.strptime(m.group(4), "%d/%m/%Y")
        qtd = int(m.group(5))
        ocorrencia = m.group(6).upper()

        dados.append({
            "funcional": funcional,
            "nome": nome,
            "ocorrencia": ocorrencia,
            "data_inicio": data_inicio,
            "data_fim": data_fim,
            "qtd_rh": qtd
        })

    return pd.DataFrame(dados)

# ---------------------------------------------------
# CALCULAR DIAS DENTRO DO MES
# ---------------------------------------------------

def dias_no_mes(data_inicio, data_fim):

    inicio_mes = datetime(ANO_REFERENCIA, MES_REFERENCIA, 1)
    fim_mes = inicio_mes + relativedelta(months=1) - relativedelta(days=1)

    inicio = max(data_inicio, inicio_mes)
    fim = min(data_fim, fim_mes)

    if inicio > fim:
        return 0

    return (fim - inicio).days + 1

# ---------------------------------------------------
# COMPARAÇÃO
# ---------------------------------------------------

def comparar(df_freq, df_rh):

    resultados = []

    for _, rh in df_rh.iterrows():

        funcional = rh["funcional"]

        dias_mes = dias_no_mes(
            rh["data_inicio"],
            rh["data_fim"]
        )

        freq = df_freq[df_freq["funcional"] == funcional]

        if not freq.empty:
            qtd_secretaria = freq.iloc[0]["qtd_secretaria"]
        else:
            qtd_secretaria = 0

        status = "OK"

        if dias_mes != qtd_secretaria:
            status = "DIVERGENTE"

        resultados.append({
            "funcional": funcional,
            "nome": rh["nome"],
            "ocorrencia": rh["ocorrencia"],
            "dias_rh_no_mes": dias_mes,
            "dias_secretaria": qtd_secretaria,
            "status": status
        })

    return pd.DataFrame(resultados)

# ---------------------------------------------------
# EXECUÇÃO
# ---------------------------------------------------

from pathlib import Path

def main():

    pasta = Path(".")

    freq_pdf = list(pasta.glob("*frequencia_secretaria*.pdf"))[0]
    rh_pdf = list(pasta.glob("*relatorio_rh*.pdf"))[0]

    texto_freq = extrair_texto_pdf(freq_pdf)
    texto_rh = extrair_texto_pdf(rh_pdf)

    df_freq = parse_frequencia(texto_freq)
    df_rh = parse_rh(texto_rh)

    resultado = comparar(df_freq, df_rh)

    resultado.to_excel("comparacao_frequencia.xlsx", index=False)

    print("\nConferência finalizada!")
    print(resultado)


if __name__ == "__main__":
    main()