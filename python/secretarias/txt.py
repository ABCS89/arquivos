import PyPDF2
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Função para extrair texto de um arquivo PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text() or ""
    return text

# Função para converter o texto extraído em um DataFrame
def convert_text_to_dataframe(text):
    lines = text.splitlines()
    data = [line.split() for line in lines if line.strip() != '']
    if len(data) > 1:
        df = pd.DataFrame(data[1:], columns=data[0])
    else:
        df = pd.DataFrame(columns=data[0])
    return df

# Função para salvar as diferenças em um arquivo Excel
def save_to_excel(dataframe, excel_path):
    dataframe.to_excel(excel_path, index=False)

# Função para salvar as diferenças em um arquivo PDF
def save_to_pdf(dataframe, pdf_path):
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    text = c.beginText(40, height - 40)
    text.setFont("Helvetica", 12)
    
    for i, row in dataframe.iterrows():
        line = ", ".join(map(str, row.values))
        text.textLine(line)
    
    c.drawText(text)
    c.showPage()
    c.save()

# Função principal que coordena a execução do script
def main():
    # Caminhos dos arquivos PDF
    pdf_path_na = "NAA_financas_frequencia.pdf"
    pdf_path_drh = "DRH_financas.pdf"
    
    # Extraindo texto dos PDFs
    text_na = extract_text_from_pdf(pdf_path_na)
    text_drh = extract_text_from_pdf(pdf_path_drh)
    
    # Convertendo texto em DataFrames
    df_na = convert_text_to_dataframe(text_na)
    df_drh = convert_text_to_dataframe(text_drh)
    
    # Encontrando diferenças
    differences = pd.merge(df_na, df_drh, how='outer', indicator=True)
    differences = differences[differences['_merge'] != 'both']
    
    # Salvando as diferenças em Excel e PDF
    save_to_excel(differences, "differences.xlsx")
    save_to_pdf(differences, "differences.pdf")

if __name__ == "__main__":
    main()
