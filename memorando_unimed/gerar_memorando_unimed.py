import pandas as pd
from docxtpl import DocxTemplate

arquivo_excel = "dados.ods"
template_word = "memorando_template.docx"

df = pd.read_excel(arquivo_excel, engine="odf")

pessoas = []

for i, row in df.iterrows():

    valores = []

    if float(row["Mensalidade"]) != 0:
        valores.append(f"Mensalidade: R$ {row['Mensalidade']:.2f}")

    if float(row["Coparticipação"]) != 0:
        valores.append(f"Coparticipação: R$ {row['Coparticipação']:.2f}")

    texto_valores = "\n".join(valores)

    pessoa = {
        "guia": i + 1,
        "nome": str(row["Funcionário"]).title(),
        "cpf": row["cpf"],
        "data_nascimento": pd.to_datetime(row["data_nascimento"]).strftime("%d/%m/%Y"),
        "endereco": str(row["Endereço"]).title(),
        "bairro" : str(row["Bairro"]).title(),
        "cidade": str(row["cidade"]).title(),   # pode fixar ou puxar de outro campo
        "uf": "SP",
        "valores": texto_valores,
        "total": f"R$ {float(row['Total']):.2f}"
    }

    pessoas.append(pessoa)

doc = DocxTemplate(template_word)

contexto = {
    "pessoas": pessoas
}

doc.render(contexto)

doc.save("memorando_final.docx")

print("Memorando gerado com sucesso!")