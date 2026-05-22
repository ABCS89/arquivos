import pdfplumber
import pandas as pd
import re
import odf
import unicodedata

import unicodedata

def linha_util_texto(linha_txt):
    if not linha_txt:
        return False
    
    linha_upper = linha_txt.upper()
    
    if any(x in linha_upper for x in IGNORAR_LINHAS):
        return False
    
    if not PADRAO_DATA.search(linha_txt):
        return False
    
    return True

    
def normalizar_texto(txt):
    if not txt:
        return ""
    
    txt = txt.upper()
    
    # remove acentos
    txt = unicodedata.normalize('NFKD', txt)
    txt = txt.encode('ASCII', 'ignore').decode('ASCII')
    
    # remove espaços duplicados
    txt = " ".join(txt.split())
    
    return txt

# =========================
# PADRÕES
# =========================

MAPA_OCORRENCIAS = {
    "ABONO": ["ABONO"],
    
    "FÉRIAS REGULAMENTARES": [
        "FÉRIAS REGULAMENTARES",
        "FERIAS REGULAMENTARES"
    ],
    
    "DOAÇÃO DE SANGUE": [
        "DOAÇÃO DE SANGUE",
        "DOACAO DE SANGUE"
    ],
    
    "TRATAMENTO DE SAÚDE": [
        "TRATAMENTO DE SAÚDE",
        "TRATAMENTO DE SAUDE",
        "TRATAMENTO DE"
    ],
    
    "AUXÍLIO DOENÇA": [
        "AUXÍLIO DOENÇA",
        "AUXILIO DOENCA"
    ],
    
    "LICENÇA MÉDICA": [
        "LICENÇA MÉDICA",
        "LICENCA MEDICA"
    ],
    
    "AFASTAMENTO SEM VENCIMENTOS": [
        "AFASTAMENTO SEM VENCIMENTOS",
        "AFASTAMENTO SEM VENCIMENTO",
        "AFASTAMENTO S/VENCIMENTOS",
        "AFASTAMENTO SEM VENC",
        "SEM VENCIMENTOS"
    ],

    "FALTA": [
        "FALTA",
        "FALTAS EFETIVOS",
        "FALT"
    ],

    "CEDIDO SEM ÔNUS PARA CEDENTE": [
        "CEDIDO SEM ÔNUS PARA CEDENTE",
        "CEDIDO SEM ÔNUS",
        "CEDIDO"
    ],

    "AFASTAMENTO POR MANDADO JUDICIAL": [
        "AFASTAMENTO POR MANDADO JUDICIAL",
        "MANDADO JUDICIAL"
    ],

    "DOENÇA EM PESSOA DA FAMÍLIA": [
        "DOENÇA EM PESSOA DA FAMÍLIA",
        "DOENÇA EM PESSOA"
    ],

    "NOJO": [
        "NOJO"
    ],

    "AGUARDANDO PERÍCIA SEMPEM": [
        "AGUARDANDO PERÍCIA SEMPEM",
        "AGUARDANDO PERÍCIA",
        "PERICIA",
        "PERÍCIA"
    ],

    "Férias prêmio": [
        "Férias prêmio",
        "PREMIO",
        "PRÊMIO"
    ],

    "Licença maternidade prorrogação": [
        "Licença maternidade prorrogação"
    ],

    "Licença maternidade": [
        "Licença maternidade"
    ]


}


PADRAO_FUNCIONAL = re.compile(r"\d{2}\.\d{3}-\d")
PADRAO_DATA = re.compile(r"\d{2}/\d{2}/\d{4}")

IGNORAR_LINHAS = [
    "REFERENTE",
    "DATA IMPRESS",
    "PÁGINA",
    "SECRETARIA",
    "DIVISÃO",
]

# =========================
# FUNÇÕES AUXILIARES
# =========================
def limpar_texto(txt):
    return " ".join(txt.split()) if txt else ""


def linha_util(linha):
    if not linha:
        return False
    
    linha_upper = linha.upper()
    
    if any(x in linha_upper for x in IGNORAR_LINHAS):
        return False
    
    if not PADRAO_FUNCIONAL.search(linha):
        return False
    
    if not PADRAO_DATA.search(linha):
        return False
    
    return True


def extrair_ocorrencia(texto):
    texto = normalizar_texto(texto)

    for padrao, variacoes in MAPA_OCORRENCIAS.items():
        for v in variacoes:
            if normalizar_texto(v) in texto:
                return padrao

    return "NÃO IDENTIFICADO"

    
# =========================
# EXTRAÇÃO - SECRETARIA
# =========================
def extrair_secretaria(pdf_path):
    dados = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            
            tabelas = pagina.extract_tables()
            
            if tabelas:
                for tabela in tabelas:
                    ultimo_funcional = None
                    
                    for linha in tabela:
                        if not linha:
                            continue
        
                        linha_txt = " ".join([str(x) for x in linha if x])
        
                        if not linha_util_texto(linha_txt):
                            continue
        
                        funcional = None
                        data = None
        
                        for item in linha:
                            if not item:
                                continue
        
                            # 🔥 extrai funcional
                            m_func = PADRAO_FUNCIONAL.search(item)
                            if m_func:
                                funcional = m_func.group()
                                ultimo_funcional = funcional
        
                            # 🔥 extrai data
                            m_data = PADRAO_DATA.search(item)
                            if m_data:
                                data = m_data.group()
        
                        # 🔥 fallback de funcional (linhas quebradas no PDF)
                        if not funcional:
                            funcional = ultimo_funcional
        
                        # 🔥 validação mínima
                        if not funcional or not data:
                            continue
        
                        ocorrencia = extrair_ocorrencia(linha_txt)
        
                        dados.append({
                            "funcional": funcional,
                            "data": data,
                            "ocorrencia": ocorrencia,
                            "origem": "secretaria"
                        })


            else:
                texto = pagina.extract_text()
                if not texto:
                    continue


                for linha in texto.split("\n"):
                
                    if not linha:
                        continue

                    linha_upper = linha.upper()

                    if any(x in linha_upper for x in IGNORAR_LINHAS):
                        continue

                    m_func = PADRAO_FUNCIONAL.search(linha)
                    m_data = PADRAO_DATA.search(linha)

                    if m_func:
                        ultimo_funcional = m_func.group()

                    if not m_data:
                        continue

                    if not ultimo_funcional:
                        continue

                    funcional = ultimo_funcional
                    data = m_data.group()
                    ocorrencia = extrair_ocorrencia(linha)

                    dados.append({
                        "funcional": funcional,
                        "data": data,
                        "ocorrencia": ocorrencia,
                        "origem": "secretaria"
                    })
                            
                                
    return pd.DataFrame(dados)


# =========================
# EXTRAÇÃO - SISTEMA (ODS)
# =========================
def extrair_sistema_ods(arquivo_ods):
    
    df = pd.read_excel(arquivo_ods, engine="odf")
    
    df.columns = df.columns.str.upper().str.strip()
    
    df = df.rename(columns={
        "FUNCIONÁRIO": "funcional",
        "DATA INICIAL": "data",
        "DESCRIÇÃO": "ocorrencia"
    })
    
    df = df.dropna(subset=["funcional", "data"])
    
    df["funcional"] = df["funcional"].astype(str).str.strip()
    
    # 🔥 CORREÇÃO PRINCIPAL
    df["data"] = pd.to_datetime(df["data"], errors="coerce").dt.strftime("%d/%m/%Y")
    
    df["ocorrencia"] = df["ocorrencia"].astype(str).str.strip()
    
    df["origem"] = "sistema"

    print(df.columns)
    
    return df[["funcional", "data", "ocorrencia", "origem"]]

# =========================
# NORMALIZAÇÃO
# =========================
def normalizar(df):
    
    df["funcional"] = df["funcional"].astype(str).str.strip()
    df["data"] = df["data"].astype(str).str.strip()
    
    # padroniza ocorrência
    df["ocorrencia"] = (
        df["ocorrencia"]
        .astype(str)
        .str.upper()
        .str.strip()
    )
    
    # 🔥 chave de comparação
    df["chave"] = df["funcional"] + "_" + df["data"]
    
    print(df.columns)

    return df
# =========================
# COMPARAÇÃO
# =========================
def comparar(df_sec, df_sis):
    
    merge = pd.merge(
        df_sec,
        df_sis,
        on="chave",
        how="outer",
        suffixes=("_sec", "_sis"),
        indicator=True
    )
    
    def classificar(row):
        if row["_merge"] == "left_only":
            return "SÓ_SECRETARIA"
        elif row["_merge"] == "right_only":
            return "SÓ_SISTEMA"
        else:
            if row["ocorrencia_sec"] != row["ocorrencia_sis"]:
                return "DIVERGENCIA_OCORRENCIA"
            return "OK"
    
    merge["status"] = merge.apply(classificar, axis=1)
    
    print(merge.columns)

    return merge

# =========================
# EXECUÇÃO
# =========================
def main():
    arquivo_secretaria = "116 - secretaria.pdf"
    arquivo_sistema = "116 - sistema.ods"
    
    print("Lendo secretaria...")
    df_sec = extrair_secretaria(arquivo_secretaria)
    
    print("Lendo sistema...")
    df_sis = extrair_sistema_ods(arquivo_sistema)
    
    print("Normalizando...")
    df_sec = normalizar(df_sec)
    df_sis = normalizar(df_sis)
    
    print("Comparando...")
    df_resultado = comparar(df_sec, df_sis)
    
    df_divergencias = df_resultado[
        df_resultado["status"] != "OK"
    ].copy()
    
    
    df_divergencias = df_divergencias[[
        "funcional_sec",
        "data_sec",
        "ocorrencia_sec",
        "ocorrencia_sis",
        "status"
    ]].sort_values(by=["funcional_sec", "data_sec"])

    print("Salvando Excel...")
    with pd.ExcelWriter("resultado_comparacao.xlsx") as writer:
        df_sec.to_excel(writer, sheet_name="Secretaria", index=False)
        df_sis.to_excel(writer, sheet_name="Sistema", index=False)
        df_divergencias.to_excel(writer, sheet_name="DIVERGENCIAS", index=False)
    
    print("✔ Finalizado!")


if __name__ == "__main__":
    main()