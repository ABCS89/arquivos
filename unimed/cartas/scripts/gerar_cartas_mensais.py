"""
gerar_cartas_mensais.py

Script único de geração das cartas mensais de notificação do Plano de Saúde
dos Servidores Públicos de Piracicaba.

Consolida o que antes eram dois scripts separados:
  - unimed_final_v2.py  (cartas "base" e "desligado")
  - gerar_cancelamento.py (cartas "aviso" e "cancelado")

A condição de cada servidor (coluna "condição" na planilha teste.ods) decide
automaticamente qual template usar e qual carta gerar:

    condição vazia/normal -> template_base.docx      -> output/cartas/
    condição = desligado  -> template_desligado.docx -> output/cartas/
    condição = aviso      -> template_aviso.docx     -> output/cancelados/
    condição = cancelado  -> template_cancelado.docx -> output/cancelados/

Todos os 4 templates agora usam o mesmo padrão de placeholder {{variavel}}
(docxtpl), então não existe mais a manipulação manual de parágrafo/run que
o script antigo precisava para os templates de colchete [placeholder].

Estrutura de pastas esperada (igual à de antes, só que agora ancorada na
localização deste próprio arquivo, então funciona de onde você rodar):

    <raiz>/
      scripts/gerar_cartas_mensais.py   <- este arquivo
      template/teste.ods
      template/devedores.xlsx
      template/template_base.docx
      template/template_desligado.docx
      template/template_aviso.docx
      template/template_cancelado.docx
      pdfs/                              <- PDFs de comprovação de envio de e-mail
      output/cartas/
      output/cancelados/
"""

import calendar
import os
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import PyPDF2
from docxtpl import DocxTemplate
from num2words import num2words

# =========================================================================
# CAMINHOS (ancorados na pasta deste script, não na pasta de execução)
# =========================================================================

BASE_DIR = Path(__file__).resolve().parent.parent

PASTA_TEMPLATE = BASE_DIR / "template"
PASTA_PDFS = BASE_DIR / "pdfs"
PASTA_SAIDA_CARTAS = BASE_DIR / "output" / "cartas"
PASTA_SAIDA_CANCELADOS = BASE_DIR / "output" / "cancelados"

ARQUIVO_BASE = PASTA_TEMPLATE / "teste.ods"
ARQUIVO_DEVEDORES = PASTA_TEMPLATE / "devedores.xlsx"

TEMPLATE_BASE = PASTA_TEMPLATE / "template_base.docx"
TEMPLATE_DESLIGADO = PASTA_TEMPLATE / "template_desligado.docx"
TEMPLATE_AVISO = PASTA_TEMPLATE / "template_aviso.docx"
TEMPLATE_CANCELADO = PASTA_TEMPLATE / "template_cancelado.docx"

# =========================================================================
# CONSTANTES
# =========================================================================

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


# =========================================================================
# FUNÇÕES AUXILIARES (antes duplicadas em vários scripts, agora só aqui)
# =========================================================================

def limpa(valor):
    """Converte NaN em string vazia; senão retorna str().strip()."""
    return "" if pd.isna(valor) else str(valor).strip()


def normalizar_nome(nome):
    """Remove acentos/pontuação e deixa em minúsculo, pra comparar nomes."""
    nome = str(nome).lower()
    substituicoes = {
        "[áàãâä]": "a", "[éèêë]": "e", "[íìîï]": "i",
        "[óòõôö]": "o", "[úùûü]": "u", "[ç]": "c",
    }
    for padrao, letra in substituicoes.items():
        nome = re.sub(padrao, letra, nome)
    nome = re.sub(r"[^a-z0-9]", "", nome)
    return nome


def capitalizar_nome(nome):
    return " ".join(palavra.capitalize() for palavra in str(nome).lower().split())


def formatar_valor_br(valor):
    """1234.5 -> '1.234,50' (sem o prefixo R$, que já vem no template)."""
    return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def valor_por_extenso(valor):
    """1234.56 -> 'mil duzentos e trinta e quatro reais e cinquenta e seis centavos'"""
    valor_float = float(valor)
    reais = int(valor_float)
    centavos = int(round((valor_float - reais) * 100))

    texto = f"{num2words(reais, lang='pt_BR')} reais"
    if centavos > 0:
        texto += f" e {num2words(centavos, lang='pt_BR')} centavos"
    return texto


def ultimo_dia_util_mes(ano, mes):
    """Retorna o dia (int) do último dia útil (não sábado/domingo) do mês."""
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data = datetime(ano, mes, ultimo_dia)
    while data.weekday() >= 5:
        data -= timedelta(days=1)
    return data.day


def formata_competencia(valor):
    """Valor de 'Mês/Ano' -> 'Jun/26'."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, str):
        return valor.strip()
    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return ""
    meses_abrev = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
                   "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    return f"{meses_abrev[data.month - 1]}/{str(data.year)[-2:]}"


def extrair_texto_pdf(caminho_pdf):
    texto = ""
    try:
        with open(caminho_pdf, "rb") as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)
            for pagina in leitor.pages:
                texto += pagina.extract_text() or ""
    except Exception as erro:
        print(f"  ⚠️ Erro ao ler PDF {caminho_pdf}: {erro}")
    return texto


def extrair_data_email_do_pdf(texto_pdf):
    """
    Procura 'Data DD/MM/AAAA' no texto do PDF (comprovante de envio de e-mail)
    e devolve um dict pronto pra jogar no contexto do docxtpl.
    Se não achar, devolve placeholders textuais pra ficar óbvio no Word
    que aquela pessoa precisa de conferência manual.
    """
    match = re.search(r"Data\s+(\d{2}/\d{2}/\d{4})", texto_pdf)

    if not match:
        return {
            "dia_email": "dia", "mes_email": "mês", "ano_email": "ano",
        }

    data_obj = datetime.strptime(match.group(1), "%d/%m/%Y")
    return {
        "dia_email": data_obj.day,
        "mes_email": MESES_PT[data_obj.month],
        "ano_email": data_obj.year,
    }


def mapear_pdfs_por_funcional(pasta_pdfs, df_base):
    """
    Varre a pasta de PDFs e casa cada arquivo com um 'Nro Funcional' da
    planilha base, comparando o nome (antes do ' - ') com o nome do
    funcionário, de forma normalizada (sem acento/case).
    """
    mapa = {}

    if not pasta_pdfs.exists():
        print(f"⚠️ Pasta de PDFs não encontrada: {pasta_pdfs}")
        return mapa

    arquivos_pdf = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith(".pdf")]

    for nome_arquivo in arquivos_pdf:
        nome_extraido = os.path.splitext(nome_arquivo)[0].split(" - ")[0]
        nome_normalizado = normalizar_nome(nome_extraido)

        for _, linha in df_base.iterrows():
            nome_planilha = normalizar_nome(linha["Funcionário"])
            if nome_normalizado in nome_planilha or nome_planilha in nome_normalizado:
                mapa[linha["Nro Funcional"]] = nome_arquivo
                break

    return mapa


# =========================================================================
# GERAÇÃO: cartas "base" e "desligado" (sem tabela de dívidas, com data de e-mail)
# =========================================================================

def gerar_carta_base_ou_desligado(linha, template_path, contexto_data_atual,
                                    contexto_vencimento, pdf_map, pasta_saida):
    nro_funcional = linha["Nro Funcional"]
    funcionario_raw = linha["Funcionário"]

    if nro_funcional in pdf_map:
        caminho_pdf = PASTA_PDFS / pdf_map[nro_funcional]
        texto_pdf = extrair_texto_pdf(caminho_pdf)
        contexto_email = extrair_data_email_do_pdf(texto_pdf)
    else:
        print(f"  ⚠️ Sem PDF de comprovação para: {funcionario_raw}")
        contexto_email = {"dia_email": "dia", "mes_email": "mês", "ano_email": "ano"}

    endereco_completo = limpa(linha.get("endereço"))
    complemento = limpa(linha.get("complemento"))
    bairro = limpa(linha.get("bairro"))
    if complemento:
        endereco_completo += f" – {complemento}"
    if bairro:
        endereco_completo += f" – {bairro}"

    contexto = {
        "nome_cap": capitalizar_nome(funcionario_raw),
        "nome_upper": str(funcionario_raw).upper(),
        "linha_endereco": endereco_completo,
        "CEP": limpa(linha.get("CEP")),
        "cidade": limpa(linha.get("cidade")),
        "valor": formatar_valor_br(linha["Total"]),
        "valor_extenso": valor_por_extenso(linha["Total"]),
        "email": limpa(linha.get("mail")) or "mail",
        **contexto_data_atual,
        **contexto_vencimento,
        **contexto_email,
    }

    doc = DocxTemplate(template_path)
    doc.render(contexto)

    nome_saida = f"{funcionario_raw}.docx"
    doc.save(pasta_saida / nome_saida)


# =========================================================================
# GERAÇÃO: cartas "aviso" e "cancelado" (com tabela de dívidas)
# =========================================================================

def gerar_carta_aviso_ou_cancelado(linha, condicao, template_path, df_dividas,
                                     contexto_data_atual, contexto_vencimento,
                                     pasta_saida):
    nome = limpa(linha.get("Funcionário"))
    matricula = limpa(linha.get("Nro Funcional"))

    df_pessoa = df_dividas[df_dividas["Funcional"] == matricula]
    if df_pessoa.empty:
        print(f"  ⚠️ Sem dívidas cadastradas para: {nome} ({condicao}) — carta NÃO gerada")
        return False

    linha_endereco = " – ".join(filter(None, [
        limpa(linha.get("endereço")),
        limpa(linha.get("bairro")),
        limpa(linha.get("complemento")),
    ]))

    tabela = []
    for _, divida in df_pessoa.iterrows():
        principal = float(divida.get("Principal (Saldo)", 0) or 0)
        total = float(divida.get("Saldo (Atualizado)", 0) or 0)
        encargos = total - principal

        data_vencimento = pd.to_datetime(
            divida.get("Data de Vencimento"), dayfirst=True, errors="coerce"
        )

        tabela.append({
            "competencia": formata_competencia(divida.get("Mês/Ano")),
            "vencimento": data_vencimento.strftime("%d/%m/%Y") if pd.notna(data_vencimento) else "",
            "principal": f"R$ {formatar_valor_br(principal)}",
            "encargos": f"R$ {formatar_valor_br(encargos)}",
            "total": f"R$ {formatar_valor_br(total)}",
        })

    contexto = {
        "nome_cap": nome.title(),
        "nome_upper": nome.upper(),
        "linha_endereco": linha_endereco,
        "CEP": limpa(linha.get("CEP")),
        "cidade": limpa(linha.get("cidade")),
        "uf": limpa(linha.get("uf")),
        "tabela": tabela,
        **contexto_data_atual,
        **contexto_vencimento,
    }

    doc = DocxTemplate(template_path)
    doc.render(contexto)

    sufixo = "aviso" if condicao == "aviso" else "cancelado"
    nome_saida = f"{nome} - {sufixo}.docx"
    doc.save(pasta_saida / nome_saida)
    return True


# =========================================================================
# PROGRAMA PRINCIPAL
# =========================================================================

def main():
    PASTA_SAIDA_CARTAS.mkdir(parents=True, exist_ok=True)
    PASTA_SAIDA_CANCELADOS.mkdir(parents=True, exist_ok=True)

    # --- datas ---
    hoje = datetime.now()
    contexto_data_atual = {
        "dia": hoje.day,
        "mes": MESES_PT[hoje.month],
        "ano": hoje.year,
    }

    dia_venc = ultimo_dia_util_mes(hoje.year, hoje.month)
    contexto_vencimento = {
        "ultimo_dia_util": dia_venc,
        "mes_vencimento": MESES_PT[hoje.month],
        "ano_vencimento": hoje.year,
        # 'ultimo_dia_do_mes' é usado só pelo template de cancelado (data efetiva
        # do cancelamento, sem ajuste de fim de semana)
        "ultimo_dia_do_mes": calendar.monthrange(hoje.year, hoje.month)[1],
    }

    # --- dados ---
    print("Carregando planilhas...")
    df_base = pd.read_excel(ARQUIVO_BASE, engine="odf")
    df_base.columns = df_base.columns.str.strip()

    df_inadimplentes = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Inadimplentes")
    df_cancelados_divida = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Cancelados")
    for df in (df_inadimplentes, df_cancelados_divida):
        df.columns = df.columns.str.strip()
        df["Funcional"] = pd.to_numeric(df["Funcional"], errors="coerce").astype("Int64").astype(str)

    df_base["Nro Funcional"] = (
        pd.to_numeric(df_base["Nro Funcional"], errors="coerce").astype("Int64")
    )

    print("Mapeando PDFs de comprovação de e-mail...")
    pdf_map = mapear_pdfs_por_funcional(PASTA_PDFS, df_base)

    # --- loop principal: uma passada só, ramifica por condição ---
    contadores = {"base": 0, "desligado": 0, "aviso": 0, "cancelado": 0, "ignorado": 0}

    for _, linha in df_base.iterrows():
        condicao = limpa(linha.get("condição")).lower()
        nome = linha.get("Funcionário")

        if pd.isna(linha.get("Nro Funcional")):
            print(f"  ⚠️ Linha sem Nro Funcional, pulando: {nome}")
            contadores["ignorado"] += 1
            continue

        if condicao == "aviso":
            gerado = gerar_carta_aviso_ou_cancelado(
                linha, "aviso", TEMPLATE_AVISO, df_inadimplentes,
                contexto_data_atual, contexto_vencimento, PASTA_SAIDA_CANCELADOS,
            )
            if gerado:
                contadores["aviso"] += 1
            else:
                contadores["ignorado"] += 1

        elif condicao == "cancelado":
            gerado = gerar_carta_aviso_ou_cancelado(
                linha, "cancelado", TEMPLATE_CANCELADO, df_cancelados_divida,
                contexto_data_atual, contexto_vencimento, PASTA_SAIDA_CANCELADOS,
            )
            if gerado:
                contadores["cancelado"] += 1
            else:
                contadores["ignorado"] += 1

        elif condicao == "desligado":
            gerar_carta_base_ou_desligado(
                linha, TEMPLATE_DESLIGADO, contexto_data_atual,
                contexto_vencimento, pdf_map, PASTA_SAIDA_CARTAS,
            )
            contadores["desligado"] += 1

        else:
            gerar_carta_base_ou_desligado(
                linha, TEMPLATE_BASE, contexto_data_atual,
                contexto_vencimento, pdf_map, PASTA_SAIDA_CARTAS,
            )
            contadores["base"] += 1

    print("\n✅ Concluído:")
    for tipo, qtd in contadores.items():
        print(f"   {tipo}: {qtd}")


if __name__ == "__main__":
    main()
