# -*- coding: utf-8 -*-
"""
Extrai os registros do relatório "Frequência dos Funcionários" gerado
pela SECRETARIA. Este PDF não vem em formato de tabela real: cada
funcionário aparece em uma linha "MATRICULA NOME OCORRENCIA [DATA QTDE]"
e, se tiver mais de uma ocorrência no mês, as ocorrências extras vêm em
linhas seguintes só com "OCORRENCIA DATA QTDE" (sem repetir matrícula/nome).

Estratégia: como os nomes vêm 100% em CAIXA ALTA e as ocorrências vêm em
Formato Título (só a 1ª letra maiúscula), separamos nome de ocorrência
procurando o primeiro token que não é totalmente maiúsculo.
"""
import re
import pdfplumber

MATRICULA_RE = re.compile(r"^\d{2}\.\d{3}-\d\b")
DATE_QTY_TAIL_RE = re.compile(r"(\d{2}/\d{2}/\d{4})\s+([\d\.,]+)\s*$")

# linhas de cabeçalho/rodapé/assinatura a ignorar
IGNORAR_SUBSTRINGS = [
    "PREFEITURA DO MUNICÍPIO",
    "Frequência dos Funcionários",
    "Referente:",
    "Nro Funcional Nome",
    "Pág.",
    "Página:",
    "Peça do processo",
    "Assinaturas do documento",
    "Código para verificação",
    "Este documento foi assinado",
    "Emitido por",
    "Assinatura ICP",
    "Assinatura do Sistema",
    "válido até",
    "Para verificar a autenticidade",
    "aponte a câmera",
    "gerada automaticamente",
    "materializada por",
    "sempapel.piracicaba",
    "CPF:",
    "NÚCLEO DE APOIO",
]


def extrair_cabecalho(caminho_pdf):
    """Extrai código da secretaria, nome da secretaria e mês/ano de referência do cabeçalho da página 1."""
    codigo_sec, nome_sec, mes_ano_ref = None, None, None
    with pdfplumber.open(caminho_pdf) as pdf:
        if not pdf.pages:
            return None, None, None
        texto = pdf.pages[0].extract_text() or ""
        for linha in texto.split("\n")[:20]:
            linha_limpa = linha.strip()

            # 1. Mês/Ano de referência (ex: Referente: Julho/2026)
            m_ref = re.search(r"Referente:\s*([A-Za-zçÇ]+)/(\d{4})", linha_limpa)
            if m_ref and not mes_ano_ref:
                mes_str, ano_str = m_ref.groups()
                meses = {
                    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
                    "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
                    "outubro": 10, "novembro": 11, "dezembro": 12
                }
                num_mes = meses.get(mes_str.lower())
                if num_mes:
                    mes_ano_ref = (int(ano_str), num_mes)

            # 2. Código e Nome da Secretaria (ex: 116 GUARDA CIVIL DO MUNICIPIO DE PIRACICABA)
            m_sec = re.match(r"^(\d{3})\s+([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ][A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s,\.\-/]+)$", linha_limpa)
            if m_sec and not codigo_sec:
                codigo_sec = m_sec.group(1)
                nome_sec = m_sec.group(2).strip()

    return codigo_sec, nome_sec, mes_ano_ref


def _eh_ignoravel(linha):
    linha_limpa = linha.strip()
    if not linha_limpa:
        return True
    if re.match(r"^\d+/\d+$", linha_limpa):  # "1/7", "2/7"...
        return True
    if re.match(r"^\d{2,4}-\d{2}", linha_limpa):  # "00-00", "00-00-000", "10-81"...
        return True
    if re.match(r"^\d{3}\s+[A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ]", linha_limpa):  # cabeçalho com código e nome da secretaria
        return True
    if re.match(r"^\d{2,4}[-\d]+\s+[A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ]", linha_limpa):  # subdivisão de departamento
        return True
    if "http://" in linha_limpa or "https://" in linha_limpa:
        return True
    for s in IGNORAR_SUBSTRINGS:
        if s in linha_limpa:
            return True
    return False


def _eh_token_maiusculo(tok):
    """True se o token é 'nome' (todo maiúsculo, ignorando pontuação)."""
    letras = [c for c in tok if c.isalpha()]
    if not letras:
        return True  # token sem letras (ex.: pontuação) não quebra o nome
    return all(c.isupper() for c in letras)


def _separar_nome_ocorrencia(texto):
    """Dado 'ACYR CARDOSO ... Frequência normal' ou 'ADILSON ... Abono',
    retorna (nome, ocorrencia)."""
    tokens = texto.split()
    corte = len(tokens)
    for i, tok in enumerate(tokens):
        if not _eh_token_maiusculo(tok):
            corte = i
            break
    nome = " ".join(tokens[:corte]).strip()
    ocorrencia = " ".join(tokens[corte:]).strip()
    return nome, ocorrencia


def extrair(caminho_pdf):
    """Retorna lista de dicts: matricula, nome, ocorrencia, data, qtde_dias."""
    registros = []
    matricula_atual = None
    nome_atual = None
    nao_reconhecidas = []

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            for linha in texto.split("\n"):
                linha = linha.strip()
                if _eh_ignoravel(linha):
                    continue

                m_data_qty = DATE_QTY_TAIL_RE.search(linha)
                data, qtde = (None, None)
                corpo = linha
                if m_data_qty:
                    data, qtde_str = m_data_qty.groups()
                    corpo = linha[: m_data_qty.start()].strip()
                    try:
                        qtde = float(qtde_str.replace(".", "").replace(",", "."))
                    except ValueError:
                        qtde = None

                if MATRICULA_RE.match(linha):
                    matricula_atual = linha[:8]
                    resto = linha[8:].strip()
                    if m_data_qty:
                        resto_corpo = corpo[8:].strip()
                    else:
                        resto_corpo = resto
                    nome_atual, ocorrencia = _separar_nome_ocorrencia(resto_corpo)
                    registros.append({
                        "matricula": matricula_atual,
                        "nome": nome_atual,
                        "ocorrencia": ocorrencia,
                        "data": data,
                        "qtde_dias": qtde,
                    })
                elif m_data_qty and corpo and corpo[:1].isalpha() and corpo[:1].isupper():
                    # linha de continuação (ocorrência extra do último funcionário)
                    # só aceitamos se tiver de fato "data + quantidade" no final,
                    # senão é ruído de marca d'água/rodapé espelhado no PDF
                    if matricula_atual is None:
                        nao_reconhecidas.append(linha)
                        continue
                    ocorrencia = corpo.strip()
                    registros.append({
                        "matricula": matricula_atual,
                        "nome": nome_atual,
                        "ocorrencia": ocorrencia,
                        "data": data,
                        "qtde_dias": qtde,
                    })
                else:
                    nao_reconhecidas.append(linha)

    return registros, nao_reconhecidas


if __name__ == "__main__":
    import sys
    import json
    regs, nao_rec = extrair(sys.argv[1])
    print(f"{len(regs)} registros extraídos.")
    if nao_rec:
        print(f"{len(nao_rec)} linha(s) não reconhecida(s) (normalmente ruído de "
              f"marca d'água/rodapé — revisar só se parecer dado de verdade):")
        for l in nao_rec:
            print("   ", repr(l))
    print(json.dumps(regs[:8], ensure_ascii=False, indent=2))
