"""
compare_and_generate_list.py
Uso:
    python compare_and_generate_list.py "106 - relatorio.pdf" "106 - frequencia.xlsx"

Descrição:
- Extrai ocorrências do relatorio.pdf
- Lê frequencia.xlsx (pandas)
- Normaliza colunas, calcula qtde dentro do mês (se necessário)
- Compara e produz <numero> - conferencia.pdf com tabela de divergências e texto do memorando
"""

import sys
import re
from pathlib import Path
from datetime import datetime, date, timedelta
import pdfplumber
import pandas as pd
from dateutil.parser import parse as dtparse
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# ---------- util ----------
def texto_para_data(s):
    for fmt in ("%d/%m/%Y","%d-%m-%Y","%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except:
            pass
    # fallback parse
    return dtparse(s, dayfirst=True).date()

def overlap_days(a_start,a_end,b_start,b_end):
    s = max(a_start, b_start)
    e = min(a_end, b_end)
    if e < s:
        return 0
    return (e - s).days + 1

def month_range_from_interval_text(txt):
    # tenta achar "Referente: 01/09/2025 a 30/09/2025"
    m = re.search(r"Referente:\s*([\d/]+)\s*a\s*([\d/]+)", txt, re.I)
    if m:
        d1 = texto_para_data(m.group(1))
        d2 = texto_para_data(m.group(2))
        return d1, d2
    # fallback: return current month
    today = date.today()
    first = date(today.year, today.month, 1)
    if today.month == 12:
        last = date(today.year, 12, 31)
    else:
        nextm = date(today.year, today.month + 1, 1)
        last = nextm - timedelta(days=1)
    return first, last

# ---------- parse relatorio ----------
def parse_relatorio_pdf(path):
    text_all = ""
    rows = []
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            text_all += p.extract_text() + "\n"
    # extrair intervalo de referência
    month_start, month_end = month_range_from_interval_text(text_all)

    # heurística: linhas com padrão "NNN.NNN-N NAME DD/MM/YYYY DD/MM/YYYY Q DESCRIÇÃO"
    lines = [ln.strip() for ln in text_all.splitlines() if ln.strip()]

    # Junta linhas que parecem parte de descrição (heurística)
    merged = []
    buf = ""
    for ln in lines:
        # se linha começa com palavra em maiúsculas + dígito (código), considera início
        if re.match(r"^[A-ZÇÃÉ.\s]+ \d{1,2}\.\d{3}-\d", ln) or re.match(r"^\d{2}\.\d{3}-\d", ln) or re.match(r"^DIVISÃO", ln) or re.match(r"^Referente:", ln):
            if buf:
                merged.append(buf.strip())
            buf = ln
        else:
            # se linha parece continuação:
            if buf:
                buf += " " + ln
            else:
                buf = ln
    if buf:
        merged.append(buf.strip())

    # Outra heurística: procurar sequências que contenham 2 datas e um número (qtde)
    for item in merged:
        # p.ex.: "DIVISÃO DE LANÇAMENTO E FISCALIZAÇÃO 28.049-6 BRENO REI PASSOS LAGOAS 16/09/2025 19/09/2025 4 Tratamento De Saúde"
        m = re.search(r"(\d{1,2}\.\d{3}-\d)\s+([A-ZÀ-ÿ0-9\s\.\-]+?)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})\s+([\d\.,]+)\s+(.+)$", item, re.I)
        if m:
            nro = m.group(1).strip()
            nome = m.group(2).strip()
            d1 = texto_para_data(m.group(3))
            d2 = texto_para_data(m.group(4))
            qtd = float(m.group(5).replace(",",".")) if m.group(5).replace(",","").replace(".","").isdigit() or ',' in m.group(5) else None
            desc = m.group(6).strip()
            rows.append({
                "Funcionário": nro,
                "Pessoa": nome,
                "DataInicial": d1,
                "DataFinal": d2,
                "QtdeDiasRelatorio": qtd,
                "Descrição": desc
            })
        else:
            # tentar outro padrão (sem ponto no código)
            m2 = re.search(r"(\d{2,3}\d*-\d)\s+([A-ZÀ-ÿ0-9\s\.\-]+?)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})\s+([\d\.,]+)\s+(.+)$", item, re.I)
            if m2:
                nro = m2.group(1).strip()
                nome = m2.group(2).strip()
                d1 = texto_para_data(m2.group(3))
                d2 = texto_para_data(m2.group(4))
                qtd = float(m2.group(5).replace(",",".")) if m2.group(5).replace(",","").replace(".","").isdigit() or ',' in m2.group(5) else None
                desc = m2.group(6).strip()
                rows.append({
                    "Funcionário": nro,
                    "Pessoa": nome,
                    "DataInicial": d1,
                    "DataFinal": d2,
                    "QtdeDiasRelatorio": qtd,
                    "Descrição": desc
                })
            else:
                # não conseguiu casar: ignorar ou guardar para inspeção
                pass

    return rows, month_start, month_end, text_all

# ---------- load frequencia ----------
def load_frequencia(path):
    # aceita xlsx ou csv
    if str(path).lower().endswith(".xlsx") or str(path).lower().endswith(".xls"):
        df = pd.read_excel(path, engine="openpyxl")
    else:
        df = pd.read_csv(path)
    # normalizar nomes de colunas para comparação
    cols = {c: c.strip().lower() for c in df.columns}
    df.columns = [c.strip() for c in df.columns]
    # tentar renomear sinônimos
    rename_map = {}
    for c in df.columns:
        lc = c.lower()
        if "funcion" in lc or "nro" in lc or "matrícula" in lc or "matricula" in lc:
            rename_map[c] = "NroFuncional"
        elif "nome" in lc or "pessoa" in lc:
            rename_map[c] = "Nome"
        elif "descri" in lc or "ocorr" in lc or "descrição" in lc:
            rename_map[c] = "Ocorrencia"
        elif re.search(r"data", lc):
            # se coluna for data única
            rename_map[c] = "Data"
        elif "qt" in lc or "quant" in lc or "qtd" in lc:
            rename_map[c] = "Quantidade"
    df = df.rename(columns=rename_map)
    # garantir colunas necessárias
    return df

# ---------- comparar ----------
def gerar_diferencas(rel_rows, freq_df, month_start, month_end):
    diffs = []
    # preparar freq_df agrupado por (NroFuncional, Ocorrencia, Data)
    # se freq tem col Data como intervalo, você pode adaptar
    # Aqui, assumimos uma linha por ocorrência no mês: (NroFuncional, Nome, Ocorrencia, Data, Quantidade)
    for r in rel_rows:
        nro = r["Funcionário"]
        nome = r["Pessoa"]
        d1 = r["DataInicial"]
        d2 = r["DataFinal"]
        desc = r["Descrição"]
        # quantidade efetiva dentro do mês
        qt_rel_no_mes = overlap_days(d1,d2,month_start,month_end)
        # se QtdeDiasRelatorio está presente, preferir QtdeDiasRelatorio ajustado:
        if r.get("QtdeDiasRelatorio"):
            # usar overlap se datas extrapolam o mês
            qt_rel = int(qt_rel_no_mes)
        else:
            qt_rel = int(qt_rel_no_mes)

        # localizar na frequencia: por nro (ou por nome se nro não bater) e por ocorrência similar
        candidates = freq_df.copy()
        if "NroFuncional" in candidates.columns:
            candidates = candidates[candidates["NroFuncional"].astype(str).str.contains(str(nro).split()[0], na=False)]
        else:
            # tentar por Nome
            if "Nome" in candidates.columns:
                candidates = candidates[candidates["Nome"].str.upper().str.contains(nome.split()[0].upper(), na=False)]
        # filtrar por ocorrência (texto)
        if "Ocorrencia" in candidates.columns:
            candidates = candidates[candidates["Ocorrencia"].str.upper().str.contains(desc.split()[0].upper(), na=False)]
        # somar quantidade encontrada na frequência para o mesmo funcionário+ocorrência
        qtd_freq_sum = 0
        if not candidates.empty:
            if "Quantidade" in candidates.columns:
                # converter quantidades
                def to_int_safe(x):
                    try:
                        return int(float(str(x).replace(",",".")))
                    except:
                        return 0
                qtd_freq_sum = candidates["Quantidade"].apply(to_int_safe).sum()
        else:
            qtd_freq_sum = 0

        # regra minutos perdidos -> somente ocorrência (comparar se existe)
        if re.search(r"minutos\s*perd", desc, re.I):
            # se existe ao menos uma linha em candidates, então OK, senão diferença
            existe = (not candidates.empty)
            if not existe:
                diffs.append({
                    "Nro": nro, "Nome": nome, "Descricao": desc,
                    "QtRel": "ocorrência", "QtFreq": 0, "Tipo": "Minutos Perdidos - ausente na frequência"
                })
            # se existe, não reportar diferença de quantidade
        else:
            if qtd_freq_sum != qt_rel:
                diffs.append({
                    "Nro": nro, "Nome": nome, "Descricao": desc,
                    "QtRel": qt_rel, "QtFreq": int(qtd_freq_sum),
                    "Diff": int(qtd_freq_sum) - int(qt_rel)
                })
    return diffs

# ---------- gerar pdf ----------
def gerar_pdf_saida(diffs, memo_text, outpath):
    doc = SimpleDocTemplate(outpath, pagesize=A4)
    styles = getSampleStyleSheet()
    elems = []
    elems.append(Paragraph("Conferência de Frequência - Diferenças Identificadas", styles["Title"]))
    elems.append(Spacer(1,12))

    if not diffs:
        elems.append(Paragraph("Nenhuma divergência encontrada.", styles["Normal"]))
    else:
        data = [["Nro Funcional","Nome","Descrição","Qt Relatório","Qt Frequência","Diferença","Observação"]]
        for d in diffs:
            data.append([
                d.get("Nro",""),
                d.get("Nome",""),
                d.get("Descricao",""),
                str(d.get("QtRel","")),
                str(d.get("QtFreq","")),
                str(d.get("Diff","")) if "Diff" in d else "",
                d.get("Tipo","")
            ])
        table = Table(data, colWidths=[70,120,160,60,60,60,120], repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.grey),
            ('GRID',(0,0),(-1,-1),0.25,colors.black),
            ('FONT', (0,0),(-1,0), 'Helvetica-Bold')
        ]))
        elems.append(table)
        elems.append(Spacer(1,12))

    elems.append(Paragraph("Texto sugerido para memorando de retificação (copiar/colar):", styles["Heading2"]))
    elems.append(Spacer(1,6))
    elems.append(Paragraph(memo_text.replace("\n","<br/>"), styles["Normal"]))

    doc.build(elems)

# ---------- main ----------
def main():
    if len(sys.argv) < 3:
        print("Uso: python compare_and_generate_list.py <relatorio.pdf> <frequencia.xlsx>")
        sys.exit(1)
    rel_path = Path(sys.argv[1])
    freq_path = Path(sys.argv[2])
    # extrair numero do nome do arquivo
    numero = rel_path.stem.split(" - ")[0] if " - " in rel_path.stem else rel_path.stem

    rel_rows, month_start, month_end, all_text = parse_relatorio_pdf(rel_path)
    freq_df = load_frequencia(freq_path)
    diffs = gerar_diferencas(rel_rows, freq_df, month_start, month_end)

    # memo text -- montar texto básico (pode adaptar)
    memo_lines = []
    memo_lines.append(f"Memorando de Solicitação de Retificação de Frequência - Referente: {month_start.strftime('%d/%m/%Y')} a {month_end.strftime('%d/%m/%Y')}")
    memo_lines.append("")
    memo_lines.append("Senhor(a),")
    memo_lines.append("Solicitamos a retificação das ocorrências abaixo, identificadas na conferência entre Relatório de Ocorrência e a Folha de Frequência:")
    memo_lines.append("")
    for d in diffs:
        memo_lines.append(f"- {d.get('Nro','')} | {d.get('Nome','')} | {d.get('Descricao','')} | Qt Rel: {d.get('QtRel','')} | Qt Freq: {d.get('QtFreq','')} | Dif: {d.get('Diff','') if 'Diff' in d else d.get('Tipo','')}")
    memo_lines.append("")
    memo_lines.append("Atenciosamente,")
    memo_lines.append("Setor de Conferência")
    memo_text = "\n".join(memo_lines)

    outpdf = f"{numero} - conferencia.pdf"
    gerar_pdf_saida(diffs, memo_text, outpdf)
    print(f"Arquivo gerado: {outpdf}")

if __name__ == "__main__":
    main()
