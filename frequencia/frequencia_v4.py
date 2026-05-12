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
                        
                        ocorrencia = linha_txt
                        
                        # remove funcional e data principal
                        if funcional:
                            ocorrencia = ocorrencia.replace(funcional, "")
                        if data:
                            ocorrencia = ocorrencia.replace(data, "")
                        
                        # 🔥 REMOVE TODAS AS DATAS RESTANTES
                        ocorrencia = re.sub(PADRAO_DATA, "", ocorrencia)
                        
                        # 🔥 REMOVE NÚMEROS SOLTOS (dias, quantidade, etc.)
                        ocorrencia = re.sub(r"\b\d+\b", "", ocorrencia)
                        
                        # limpeza final
                        ocorrencia = limpar_texto(ocorrencia)
                        
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
                    
                    match_funcional = PADRAO_FUNCIONAL.search(linha)
                    match_data = PADRAO_DATA.search(linha)

                    if not match_funcional or not match_data:
                        continue

                    funcional = match_funcional.group()
                    data = match_data.group()
                    
                    ocorrencia = linha
                    
                    if funcional:
                        ocorrencia = ocorrencia.replace(funcional, "")
                    if data:
                        ocorrencia = ocorrencia.replace(data, "")
                    
                    # 🔥 LIMPEZA INTELIGENTE
                    ocorrencia = re.sub(PADRAO_DATA, "", ocorrencia)
                    ocorrencia = re.sub(r"\b\d+\b", "", ocorrencia)
                    
                    ocorrencia = limpar_texto(ocorrencia)
                    
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
                
                match_funcional = PADRAO_FUNCIONAL.search(linha)
                match_data = PADRAO_DATA.search(linha)

                if not match_funcional or not match_data:
                    continue

                funcional = match_funcional.group()
                data = match_data.group()
                
                ocorrencia = linha
                
                if funcional:
                    ocorrencia = ocorrencia.replace(funcional, "")
                if data:
                    ocorrencia = ocorrencia.replace(data, "")
                
                # 🔥 REMOVE DATAS INTERNAS (PERÍODOS)
                ocorrencia = re.sub(PADRAO_DATA, "", ocorrencia)
                
                # 🔥 REMOVE NÚMEROS SOLTOS
                ocorrencia = re.sub(r"\b\d+\b", "", ocorrencia)
                
                # limpeza final
                ocorrencia = limpar_texto(ocorrencia)
                
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
    df["ocorrencia"] = (
        df["ocorrencia"]
        .astype(str)
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    
    # REMOVE nomes próprios (opcional mas recomendado)
    # isso evita erro por variação de nome
    df["ocorrencia_limpa"] = (
        df["ocorrencia"]
        .str.replace(r"[A-ZÀ-Ú]+(?: [A-ZÀ-Ú]+)+", "", regex=True)
        .str.strip()
    )
    
    # chave baseada no que REALMENTE importa
    df["chave"] = (
        df["funcional"].astype(str).str.strip() + "_" +
        df["ocorrencia_limpa"]
    )
    
    df = df.drop_duplicates(subset=["chave"])
    
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
    
    # =========================
    # CRIAR ABA DE DIVERGÊNCIAS
    # =========================
    
    divergencias_sec = df_sec[df_sec["status"] == "SÓ_SECRETARIA"].copy()
    divergencias_sis = df_sis[df_sis["status"] == "SÓ_SISTEMA"].copy()
    
    divergencias_sec["tipo_erro"] = "FALTA NO SISTEMA"
    divergencias_sis["tipo_erro"] = "FALTA NA SECRETARIA"
    
    df_divergencias = pd.concat([divergencias_sec, divergencias_sis])
    
    # organiza melhor
    df_divergencias = df_divergencias[[
        "funcional", "data", "ocorrencia", "tipo_erro"
    ]].sort_values(by=["funcional", "data", "tipo_erro"])
    
    print("Salvando Excel...")
    with pd.ExcelWriter("resultado_comparacao.xlsx") as writer:
        df_sec.to_excel(writer, sheet_name="Secretaria", index=False)
        df_sis.to_excel(writer, sheet_name="Sistema", index=False)
        df_divergencias.to_excel(writer, sheet_name="DIVERGENCIAS", index=False)
    
    print("✔ Finalizado com aba de divergências!")

if __name__ == "__main__":
    main()