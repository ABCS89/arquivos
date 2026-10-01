import pdfplumber
import os

pdf_path = os.path.join("data", "Funcionário por Cargo - 09-2026.pdf")

with pdfplumber.open(pdf_path) as pdf:
    print(f"Total de páginas no PDF: {len(pdf.pages)}")
    p0 = pdf.pages[0]
    text0 = p0.extract_text()
    print("--- AMOSTRA PÁGINA 1 (Primeiras 30 linhas) ---")
    for line in text0.split("\n")[:30]:
        print(line)
