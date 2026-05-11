import pandas as pd
from docxtpl import DocxTemplate

# Arquivos
arquivo_excel = "../template/teste.ods"
template_word = "../template/template_lista.docx"
saida = "lista_gerada.docx"

# Lê o .ods
df = pd.read_excel(arquivo_excel, engine="odf")

# Aqui você define o nome da coluna que tem os funcionários
# (ajuste conforme sua planilha)
coluna_nome = "Funcionário"

# Cria lista de pessoas
pessoas = []

for _, row in df.iterrows():
    if pd.notna(row[coluna_nome]):
        pessoas.append({
            "nome": str(row[coluna_nome]).strip().upper()
        })

# Carrega template
doc = DocxTemplate(template_word)

# Contexto
contexto = {
    "pessoas": pessoas
}

# Renderiza
doc.render(contexto)

# Salva
doc.save(saida)

print("Documento gerado com sucesso!")