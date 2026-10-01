import os
import pandas as pd

def inspecionar():
    print("=== ARQUIVO DE NOVOS E ALTERAÇÕES ===")
    novos_path = os.path.join("data", "novos_e_alteracoes", "Escolhas das empresas de Vale Alimentação.ods")
    if os.path.exists(novos_path):
        xl = pd.ExcelFile(novos_path, engine="odf")
        print(f"Abas encontradas: {xl.sheet_names}")
        for s in xl.sheet_names:
            df = xl.parse(s)
            print(f"\n--- Aba: '{s}' | Linhas: {len(df)}, Colunas: {len(df.columns)} ---")
            print("Colunas:", list(df.columns))
            if not df.empty:
                print("Exemplo 1ª linha:")
                for col in df.columns:
                    print(f"  {col}: {df.iloc[0][col]}")
    else:
        print(f"Arquivo não encontrado: {novos_path}")

    print("\n" + "="*50)
    print("=== ARQUIVOS DE ENTREGA INICIAL ===")
    inicial_dir = os.path.join("data", "entrega_inicial")
    if not os.path.exists(inicial_dir):
        print("Diretório não existe.")
        return

    arquivos = [f for f in os.listdir(inicial_dir) if f.endswith(('.ods', '.xls', '.xlsx'))]
    for fname in sorted(arquivos):
        fpath = os.path.join(inicial_dir, fname)
        ext = os.path.splitext(fname)[1].lower()
        engine = "odf" if ext == ".ods" else ("openpyxl" if ext == ".xlsx" else "xlrd")
        try:
            xl = pd.ExcelFile(fpath, engine=engine)
            print(f"\nArquivo: {fname}")
            print(f"  Total de abas: {len(xl.sheet_names)}")
            print(f"  Nomes das abas (primeiras 5): {xl.sheet_names[:5]}")
            s0 = xl.sheet_names[0]
            df0 = xl.parse(s0, nrows=2)
            print(f"  Aba de exemplo ('{s0}') colunas: {list(df0.columns)}")
            if not df0.empty:
                ex = {k: v for k, v in df0.iloc[0].to_dict().items() if pd.notna(v)}
                print(f"  Amostra 1º registro: {ex}")
        except Exception as e:
            print(f"  Erro ao ler {fname}: {e}")

if __name__ == "__main__":
    inspecionar()
