import pdfplumber

with pdfplumber.open("frequencia.pdf") as pdf:
    page = pdf.pages[0]
    text = page.extract_text()
    print(text[:2000])  # mostra só os primeiros 2000 caracteres
