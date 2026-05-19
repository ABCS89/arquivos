import pdfplumber
import pandas as pd
import re
import odf

# =========================
# PADRÕES
# =========================

OCORRENCIAS_VALIDAS = [
    "ABONO",
    "FÉRIAS REGULAMENTARES",
    "DOAÇÃO DE SANGUE",
    "TRATAMENTO DE SAÚDE",
    "AUXÍLIO DOENÇA",
    "LICENÇA MÉDICA",
    "Aguardando perícia sempem",
]

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
    texto = texto.upper()

    for ocorrencia in OCORRENCIAS_VALIDAS:
        if ocorrencia in texto:
            return ocorrencia

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
                    for linha in tabela:
                        if not linha:
                            continue
                        
                        linha_txt = " ".join([str(x) for x in linha if x])
                        
                        if not linha_util(linha_txt):
                            continue
                        
                        funcional = None
                        data = None
                        
                        for item in linha:
                            if not item:
                                continue
                            
                            if not funcional:
                                m = PADRAO_FUNCIONAL.search(item)
                                if m:
                                    funcional = m.group()
                            
                            if not data:
                                m = PADRAO_DATA.search(item)
                                if m:
                                    data = m.group()
                        
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
                    
                    if not linha_util(linha):
                        continue
                    
                    m_func = PADRAO_FUNCIONAL.search(linha)
                    m_data = PADRAO_DATA.search(linha)

                    if not m_func or not m_data:
                        continue

                    funcional = m_func.group()
                    data = m_data.group()
                    
                    ocorrencia = extrair_ocorrencia(linha_txt)
                    
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