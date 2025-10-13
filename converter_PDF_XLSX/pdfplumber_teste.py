import pdfplumber
import pandas as pd
from openpyxl import load_workbook

pdf_path = "cns_relatorio_ocorrencia_geral.pdf"
output_path = "relatorio_ocorrencia_geral.xlsx"

dados = []
colunas = ["Divisão", "Funcionário", "Pessoa", "Data Inicial", "Data Final", "Qtde Dias", "Descrição"]

with pdfplumber.open(pdf_path) as pdf:
    for pagina in pdf.pages:
        # Tenta detectar tabelas na página
        tabela = pagina.extract_table()
        if tabela:
            for linha in tabela:
                if linha and not "Divisão" in linha[0]:
                    dados.append(linha)

# Cria um DataFrame
df = pd.DataFrame(dados, columns=colunas)

# Remove linhas vazias
df = df.dropna(how="all")
df = df.fillna("")

# Salva em Excel
df.to_excel(output_path, index=False)

# Ajusta largura automática das colunas
wb = load_workbook(output_path)
ws = wb.active
for col in ws.columns:
    max_len = max(len(str(c.value or "")) for c in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 2
wb.save(output_path)

print(f"✅ Arquivo Excel criado com sucesso: {output_path}")
