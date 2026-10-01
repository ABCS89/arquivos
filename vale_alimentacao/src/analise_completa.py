import os
import pandas as pd

def analisar_arquivos():
    inicial_dir = os.path.join("data", "entrega_inicial")
    print("=" * 60)
    print("RESUMO DA ENTREGA INICIAL (ABRIL)")
    print("=" * 60)
    
    total_geral_servidores = 0
    arquivos = [f for f in sorted(os.listdir(inicial_dir)) if f.endswith(('.ods', '.xls', '.xlsx'))]
    
    detalhes_arquivos = []
    
    for fname in arquivos:
        fpath = os.path.join(inicial_dir, fname)
        ext = os.path.splitext(fname)[1].lower()
        engine = "odf" if ext == ".ods" else ("openpyxl" if ext == ".xlsx" else "xlrd")
        try:
            xl = pd.ExcelFile(fpath, engine=engine)
            total_arquivo = 0
            abas_info = []
            for s in xl.sheet_names:
                df = xl.parse(s, header=None)
                # Verifica onde está a linha de cabeçalho com CPF
                header_idx = None
                for idx in range(min(15, len(df))):
                    row_vals = [str(x).strip().upper() for x in df.iloc[idx].values if pd.notna(x)]
                    if "CPF" in row_vals:
                        header_idx = idx
                        break
                
                if header_idx is not None and header_idx + 1 < len(df):
                    dados = df.iloc[header_idx + 1:]
                    # considera linha válida se tiver algum valor não nulo nas primeiras 2 colunas
                    linhas_validas = dados[dados.iloc[:, 0].notna() | dados.iloc[:, 1].notna()]
                    qtd = len(linhas_validas)
                else:
                    qtd = 0
                
                total_arquivo += qtd
                abas_info.append((s, qtd))
                
            total_geral_servidores += total_arquivo
            detalhes_arquivos.append({
                "arquivo": fname,
                "qtd_abas": len(xl.sheet_names),
                "total_servidores": total_arquivo,
                "abas": abas_info
            })
            print(f"- {fname}: {len(xl.sheet_names)} abas | Total de registros: {total_arquivo}")
        except Exception as e:
            print(f"- {fname}: ERRO ({e})")
            
    print(f"\nTOTAL GERAL ESTIMADO DE SERVIDORES NA ENTREGA INICIAL: {total_geral_servidores}")
    
    print("\n" + "=" * 60)
    print("ARQUIVO DE NOVOS E ALTERAÇÕES")
    print("=" * 60)
    novos_path = os.path.join("data", "novos_e_alteracoes", "Escolhas das empresas de Vale Alimentação.ods")
    if os.path.exists(novos_path):
        df_novos = pd.read_excel(novos_path, engine="odf")
        # Filtrar linhas válidas de nomes
        validos = df_novos[df_novos['Nome'].notna() & (df_novos['Nome'] != 'Nome')]
        print(f"Total de registros de servidores na planilha de escolhas: {len(validos)}")
        print("\nDistribuição por Empresa escolhida:")
        print(validos['Empresa do V.A:'].value_counts(dropna=False))
        print("\nDistribuição por Cargo (top 10):")
        print(validos['Cargo'].value_counts(dropna=False).head(10))

if __name__ == "__main__":
    analisar_arquivos()
