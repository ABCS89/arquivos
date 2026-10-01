import os
import unicodedata
import pandas as pd

def norm(text):
    if pd.isna(text):
        return ""
    t = str(text).strip().upper()
    t = unicodedata.normalize('NFKD', t).encode('ASCII', 'ignore').decode('ASCII')
    return " ".join(t.split())

def testar_cruzamento():
    inicial_dir = os.path.join("data", "entrega_inicial")
    arquivos = [f for f in sorted(os.listdir(inicial_dir)) if f.endswith(('.ods', '.xls', '.xlsx'))]
    
    registros_iniciais = []
    
    for fname in arquivos:
        fpath = os.path.join(inicial_dir, fname)
        ext = os.path.splitext(fname)[1].lower()
        engine = "odf" if ext == ".ods" else ("openpyxl" if ext == ".xlsx" else "xlrd")
        xl = pd.ExcelFile(fpath, engine=engine)
        for s in xl.sheet_names:
            df = xl.parse(s, header=None)
            header_idx = None
            for idx in range(min(15, len(df))):
                row_vals = [str(x).strip().upper() for x in df.iloc[idx].values if pd.notna(x)]
                if "CPF" in row_vals:
                    header_idx = idx
                    break
            
            if header_idx is not None and header_idx + 1 < len(df):
                cpf_col = None
                nome_col = None
                for col_idx in range(df.shape[1]):
                    val = str(df.iloc[header_idx, col_idx]).strip().upper()
                    if val == "CPF":
                        cpf_col = col_idx
                    elif val in ["NOME", "NOME DO FUNCIONÁRIO", "FUNCIONÁRIO"]:
                        nome_col = col_idx
                
                if cpf_col is not None and nome_col is not None:
                    dados = df.iloc[header_idx + 1:].copy()
                    for _, row in dados.iterrows():
                        cpf = str(row[cpf_col]).strip() if pd.notna(row[cpf_col]) else ""
                        nome = str(row[nome_col]).strip() if pd.notna(row[nome_col]) else ""
                        if nome and nome.upper() != "NOME":
                            registros_iniciais.append({
                                "arquivo_origem": fname,
                                "aba_origem": s,
                                "cpf": cpf,
                                "nome": nome,
                                "nome_norm": norm(nome)
                            })
                            
    df_base_inicial = pd.DataFrame(registros_iniciais)
    print(f"Total de registros carregados da entrega inicial: {len(df_base_inicial)}")
    
    # Agora carrega a lista de escolhas
    novos_path = os.path.join("data", "novos_e_alteracoes", "Escolhas das empresas de Vale Alimentação.ods")
    df_novos = pd.read_excel(novos_path, engine="odf")
    validos = df_novos[df_novos['Nome'].notna() & (df_novos['Nome'] != 'Nome')].copy()
    validos['nome_norm'] = validos['Nome'].apply(norm)
    
    # Cruzamento
    cruzados = []
    novos_absolutos = []
    
    # Cria dict para busca rápida
    nomes_base = df_base_inicial.groupby('nome_norm').first().to_dict(orient='index')
    
    for idx, r in validos.iterrows():
        nn = r['nome_norm']
        if nn in nomes_base:
            match = nomes_base[nn]
            cruzados.append({
                "nome": r['Nome'],
                "cargo": r['Cargo'],
                "secretaria": r['SECRETARIA'],
                "inicio": r['Início'],
                "empresa_escolhida": r['Empresa do V.A:'],
                "cpf_encontrado": match['cpf'],
                "origem_anterior": f"{match['arquivo_origem']} ({match['aba_origem']})"
            })
        else:
            novos_absolutos.append({
                "nome": r['Nome'],
                "cargo": r['Cargo'],
                "secretaria": r['SECRETARIA'],
                "inicio": r['Início'],
                "empresa_escolhida": r['Empresa do V.A:']
            })
            
    print(f"\nTotal na planilha de Escolhas: {len(validos)}")
    print(f"Servidores que JÁ CONSTAVAM na base de Abril (Troca de Cargo / Reemissão): {len(cruzados)}")
    print(f"Servidores que NÃO CONSTAVAM na base inicial (Admissões Novas): {len(novos_absolutos)}")
    
    print("\nExemplo de servidores encontrados na base inicial (Trocas de cargo):")
    for item in cruzados[:5]:
        print(f"  {item['nome']} | CPF: {item['cpf_encontrado']} | Cargo Atual: {item['cargo']} | Antigo: {item['origem_anterior']}")
        
    print("\nExemplo de servidores NÃO encontrados (Novos servidores):")
    for item in novos_absolutos[:5]:
        print(f"  {item['nome']} | Cargo: {item['cargo']} | Sec: {item['secretaria']} | Início: {item['inicio']}")

if __name__ == "__main__":
    testar_cruzamento()
