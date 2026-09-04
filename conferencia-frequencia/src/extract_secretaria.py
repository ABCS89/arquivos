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
    "GUARDA CIVIL DO MUNICÍPIO",
    "NÚCLEO DE APOIO",
]


def _eh_ignoravel(linha):
    if not linha.strip():
        return True
    if re.match(r"^\d+/\d+$", linha.strip()):  # "1/7", "2/7"...
        return True
    if re.match(r"^\d{2}-\d{2}", linha.strip()):  # "00-00", "00-00-000"...
        return True
    for s in IGNORAR_SUBSTRINGS:
        if s in linha:
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
