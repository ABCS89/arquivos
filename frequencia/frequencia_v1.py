import pdfplumber
import pandas as pd
import re

# =========================
# PADRÕES
# =========================
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
    
    # ignora lixo conhecido
    if any(x in linha_upper for x in IGNORAR_LINHAS):
        return False
    
    # precisa ter funcional e data
    if not PADRAO_FUNCIONAL.search(linha):
        return False
    
    if not PADRAO_DATA.search(linha):
        return False
    
    return True


# =========================
# EXTRAÇÃO - SECRETARIA
# =========================
def extrair_secretaria(pdf_path):
    dados = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            
            # TENTA EXTRAIR COMO TABELA PRIMEIRO
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
                                match = PADRAO_FUNCIONAL.search(item)
                                if match:
                                    funcional = match.group()
                            
                            if not data:
                                match = PADRAO_DATA.search(item)
                                if match:
                                    data = match.group()
                        
                        ocorrencia = limpar_texto(linha_txt)
                        
                        dados.append({
                            "funcional": funcional,
                            "data": data,
                            "ocorrencia": ocorrencia,
                            "origem": "secretaria"
                        })
            
            # FALLBACK → TEXTO
            else:
                texto = pagina.extract_text()
                if not texto:
                    continue
                
                for linha in texto.split("\n"):
                    
                    if not linha_util(linha):
                        continue
                    
                    funcional = PADRAO_FUNCIONAL.search(linha).group()
                    data = PADRAO_DATA.search(linha).group()
                    
                    ocorrencia = limpar_texto(linha)
                    
                    dados.append({
                        "funcional": funcional,
                        "data": data,
                        "ocorrencia": ocorrencia,
                        "origem": "secretaria"
                    })
    
    return pd.DataFrame(dados)


# =========================
# EXTRAÇÃO - SISTEMA
# =========================
def extrair_sistema(pdf_path):
    dados = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            
            for linha in texto.split("\n"):
                
                if not linha_util(linha):
                    continue
                
                funcional = PADRAO_FUNCIONAL.search(linha).group()
                data = PADRAO_DATA.search(linha).group()
                
                ocorrencia = limpar_texto(linha)
                
                dados.append({
                    "funcional": funcional,
                    "data": data,
                    "ocorrencia": ocorrencia,
                    "origem": "sistema"
                })
    
    return pd.DataFrame(dados)


# =========================
# NORMALIZAÇÃO
# =========================
def normalizar(df):
    df["chave"] = (
        df["funcional"].astype(str).str.strip() + "_" +
        df["data"].astype(str).str.strip() + "_" +
        df["ocorrencia"].str.upper().str.strip()
    )
    return df


# =========================
# COMPARAÇÃO
# =========================
def comparar(df_sec, df_sis):
    
    sec_keys = set(df_sec["chave"])
    sis_keys = set(df_sis["chave"])
    
    somente_sec = sec_keys - sis_keys
    somente_sis = sis_keys - sec_keys
    iguais = sec_keys & sis_keys
    
    df_sec["status"] = df_sec["chave"].apply(
        lambda x: "OK" if x in iguais else "SÓ_SECRETARIA"
    )
    
    df_sis["status"] = df_sis["chave"].apply(
        lambda x: "OK" if x in iguais else "SÓ_SISTEMA"
    )
    
    return df_sec, df_sis


# =========================
# EXECUÇÃO
# =========================
def main():
    arquivo_secretaria = "116 - secretaria.pdf"
    arquivo_sistema = "116 - sistema.pdf"
    
    print("Lendo secretaria...")
    df_sec = extrair_secretaria(arquivo_secretaria)
    
    print("Lendo sistema...")
    df_sis = extrair_sistema(arquivo_sistema)
    
    print("Normalizando...")
    df_sec = normalizar(df_sec)
    df_sis = normalizar(df_sis)
    
    print("Comparando...")
    df_sec, df_sis = comparar(df_sec, df_sis)
    
    print("Salvando Excel...")
    with pd.ExcelWriter("resultado_comparacao.xlsx") as writer:
        df_sec.to_excel(writer, sheet_name="Secretaria", index=False)
        df_sis.to_excel(writer, sheet_name="Sistema", index=False)
    
    print("✔ Finalizado!")


if __name__ == "__main__":
    main()