import pdfplumber
import pandas as pd

pdf_path = "frequencia.pdf"
output_path = "frequencia_funcionarios.xlsx"

todas_linhas = []

with pdfplumber.open(pdf_path) as pdf:
    for num_pagina, pagina in enumerate(pdf.pages, start=1):
        # Tenta extrair as tabelas com detecção de texto mais tolerante
        tabelas = pagina.extract_tables({
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "intersection_tolerance": 8,
            "snap_tolerance": 3,
            "join_tolerance": 3,
            "edge_min_length": 3,
        })

        for tabela in tabelas:
            for linha in tabela:
                # Remove linhas vazias e cabeçalhos repetidos
                if not linha or all(c is None or str(c).strip() == "" for c in linha):
                    continue
                if any("Nro Funcional" in str(c) for c in linha):
                    continue
                todas_linhas.append(linha)

# Cria DataFrame consolidado
df = pd.DataFrame(todas_linhas)

# Ajuste opcional: renomear colunas se a estrutura for consistente
colunas_padrao = ["Nro Funcional", "Nome", "Ocorrência", "Data", "Quantidade"]
if len(df.columns) >= 5:
    df.columns = colunas_padrao + list(df.columns[len(colunas_padrao):])

# Remove duplicatas e espaços extras
df = df.drop_duplicates().fillna("").applymap(lambda x: str(x).strip())

# Salva em Excel
df.to_excel(output_path, index=False)
print(f"✅ Arquivo Excel criado com sucesso: {output_path}")
