import os
import re
import pdfplumber
from datetime import datetime

def extrair_funcionarios_pdf(caminho_pdf):
    """
    Extrai servidores do relatório oficial 'Funcionário por Cargo' em PDF.
    Retorna lista de dicionários com:
    - nro_funcional_original
    - funcional (inteiro)
    - nome
    - admissao (string dd/mm/aaaa)
    - admissao_dt (objeto datetime date)
    - cargo
    """
    if not os.path.exists(caminho_pdf):
        raise FileNotFoundError(f"Arquivo PDF não encontrado: {caminho_pdf}")

    padrao_func = re.compile(r'^(\d{1,3}\.\d{3}-[\dwW])\s+(.+?)\s+(\d{2}/\d{2}/\d{4})\b')
    
    servidores = []
    cargo_atual = "NÃO INFORMADO"

    with pdfplumber.open(caminho_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                line = line.strip()
                if line.startswith('Cargo:'):
                    m_cargo = re.search(r'Cargo:\s*(.+?)(?:\s+CBO:|$)', line)
                    if m_cargo:
                        cargo_atual = m_cargo.group(1).strip()
                else:
                    m_func = padrao_func.match(line)
                    if m_func:
                        nro_func = m_func.group(1)
                        nome = m_func.group(2).strip()
                        admissao_str = m_func.group(3)
                        func_digitos = re.sub(r'\D', '', nro_func)
                        funcional_int = int(func_digitos) if func_digitos.isdigit() else nro_func
                        
                        try:
                            dt = datetime.strptime(admissao_str, "%d/%m/%Y").date()
                        except ValueError:
                            dt = None

                        servidores.append({
                            "nro_funcional_original": nro_func,
                            "funcional": funcional_int,
                            "nome": nome,
                            "admissao": admissao_str,
                            "admissao_dt": dt,
                            "cargo": cargo_atual
                        })

    return servidores
