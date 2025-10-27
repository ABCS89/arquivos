#!/usr/bin/env python3
# confirmacao_retificacao.py
"""
Uso:
  python confirmacao_retificacao.py "106 - conferencia.pdf" "106 - retificacao_memo.pdf"

Gera: "106 - confirmacao_retificacao.pdf"

Descrição:
 - Este script tenta extrair do memorando (ou do PDF de retificação recebido) os itens retificados
   e compara com a saída de conferência (ou com os dados originais).
 - O PDF de retificação pode variar muito; o script busca por padrões de número funcional (xx.xxx-x)
   e por linhas com "nro | nome | ocorrencia | data | qtde" no texto do memorando.
"""
import sys
from pathlib import Path
import pdfplumber
import re
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

def extract_text(pdf_path):
    txt = ""
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            txt += (p.extract_text() or "") + "\n"
    return txt

def parse_retificacao_from_text(text):
    """
    Busca linhas com padrão de nº funcional e capta os campos.
    Retorna lista de dicts: nro_funcional, nome, ocorrencia, data, quantidade
    """
    results = []
    # padrão comum no memorando que geramos: "- 28.596-0 | MARIA ... | Ocorrência ... | Data ... | Qtde: ..."
    line_pat = re.compile(r"(\d{2,3}\.\d{3}-\d)\s*\|\s*([^|]+)\s*\|\s*([^|]+)\s*(?:\|\s*Data[: ]*([\d/]+))?\s*(?:\|\s*Qtde[: ]*([\d\.,]+))?", flags=re.I)
    for line in text.splitlines():
        m = line_pat.search(line)
        if m:
            nro = m.group(1).strip()
            nome = m.group(2).strip()
            ocorr = m.group(3).strip()
            data = m.group(4) or ""
            qtd = m.group(5) or ""
            try:
                qtd_val = float(qtd.replace(".", "").replace(",", ".")) if qtd else None
            except:
                qtd_val = None
            results.append({
                "nro_funcional": nro,
                "nome": nome,
                "ocorrencia": ocorr,
                "data": data,
                "quantidade": qtd_val
            })
    return results

def generate_confirmation_pdf(prefix, confirmed, remaining_diffs, output_path):
    c = canvas.Canvas(str(output_path), pagesize=A4)
    w, h = A4
    margin = 15*mm
    y = h - margin
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin, y, f"Confirmação de Retificação - {prefix}")
    y -= 18
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin, y, "Itens ajustados conforme memorando recebido:")
    y -= 14
    c.setFont("Helvetica", 9)
    if not confirmed:
        c.drawString(margin, y, "Nenhum item identificado como retificado no memorando.")
        y -= 12
    else:
        for r in confirmed:
            if y < 60:
                c.showPage(); y = h - margin
            c.drawString(margin, y, f"- {r['nro_funcional']} | {r['nome']} | {r['ocorrencia']} | {r.get('data','')} | {r.get('quantidade','')}")
            y -= 12

    if y < 100:
        c.showPage(); y = h - margin
    y -= 8
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin, y, "Diferenças que permaneceram após a retificação:")
    y -= 14
    c.setFont("Helvetica", 9)
    if not remaining_diffs:
        c.drawString(margin, y, "Nenhuma diferença pendente.")
        y -= 12
    else:
        for tag, frec, rel in remaining_diffs:
            if y < 60:
                c.showPage(); y = h - margin
            if tag == "FREQUENCIA_ONLY":
                txt = f"- FREQUÊNCIA_ONLY: {frec['nro_funcional']} | {frec['nome']} | {frec['ocorrencia']} | {frec.get('data','')} | {frec.get('quantidade','')}"
            else:
                txt = f"- RELATORIO_ONLY: {rel['nro_funcional']} | {rel['nome']} | {rel['ocorrencia']} | {rel.get('data','')} | qtde_mes={rel.get('quantidade_mes','')}"
            c.drawString(margin, y, txt)
            y -= 12
    c.showPage()
    c.save()

def main():
    if len(sys.argv) < 3:
        print("Uso: python confirmacao_retificacao.py <arquivo_conferencia.pdf> <arquivo_retificacao_memo.pdf>")
        sys.exit(1)
    conf_path = Path(sys.argv[1])
    memo_path = Path(sys.argv[2])
    prefix = conf_path.name.split(" - ")[0]
    # extrair texto do memorando
    memo_text = extract_text(memo_path)
    parsed = parse_retificacao_from_text(memo_text)
    # load conferência original (precisamos dos diffs que geramos no passo 1)
    # Assumimos que o conferência original foi gerado com o script A e que o usuário
    # também tem os CSVs ou podemos re-run parsing on original PDFs.
    # Aqui vamos: se existir "<prefix> - frequencia.pdf" e "<prefix> - relatório.pdf", reparse e recompute diffs.
    freq_pdf = Path(f"{prefix} - frequencia.pdf")
    rel_pdf = Path(f"{prefix} - relatório.pdf")
    if freq_pdf.exists() and rel_pdf.exists():
        from conferencia_frequencia import parse_frequencia, parse_relatorio, find_referencia_mes_from_frequencia, extract_text_lines, compare_dataframes
        lines = extract_text_lines(freq_pdf)
        ref_year, ref_month = find_referencia_mes_from_frequencia(lines)
        df_freq = parse_frequencia(freq_pdf)
        df_rel = parse_relatorio(rel_pdf)
        matches, diffs = compare_dataframes(df_freq, df_rel, ref_year, ref_month)
    else:
        # Se não existir os PDFs, tentamos ler do arquivo de conferência (apenas texto)
        conf_text = extract_text(conf_path)
        # tentar recuperar as diferenças do PDF de conferência por regex (linhas que começam com '- Registro...')
        diffs = []
        for line in conf_text.splitlines():
            m1 = re.search(r"Registro em FREQUÊNCIA presente e não encontrado no RELATÓRIO: (.+)", line)
            if m1:
                full = m1.group(1)
                parts = [p.strip() for p in full.split("|")]
                if len(parts) >= 4:
                    diffs.append(("FREQUENCIA_ONLY", {
                        "nro_funcional": parts[0],
                        "nome": parts[1],
                        "ocorrencia": parts[2],
                        "data": parts[3] if len(parts) > 3 else ""
                    }, None))
            m2 = re.search(r"Registro em RELATÓRIO presente e não encontrado na FREQUÊNCIA: (.+)", line)
            if m2:
                full = m2.group(1)
                parts = [p.strip() for p in full.split("|")]
                if len(parts) >= 4:
                    diffs.append(("RELATORIO_ONLY", None, {
                        "nro_funcional": parts[0],
                        "nome": parts[1],
                        "ocorrencia": parts[2],
                        "data": parts[3] if len(parts) > 3 else ""
                    }))
    # quais desses diffs foram corrigidos pelo memo? comparando nro_funcional + ocorrencia (substring)
    confirmed = []
    remaining = []
    for item in diffs:
        tag, frec, rel = item
        found = False
        for p in parsed:
            if tag == "FREQUENCIA_ONLY" and frec:
                if p["nro_funcional"] == frec["nro_funcional"] and (p["ocorrencia"].lower() in frec["ocorrencia"].lower() or frec["ocorrencia"].lower() in p["ocorrencia"].lower()):
                    found = True; break
            elif tag == "RELATORIO_ONLY" and rel:
                if p["nro_funcional"] == rel["nro_funcional"] and (p["ocorrencia"].lower() in rel["ocorrencia"].lower() or rel["ocorrencia"].lower() in p["ocorrencia"].lower()):
                    found = True; break
        if found:
            # marcar confirmado
            confirmed.append(p)
        else:
            remaining.append(item)
    out_name = f"{prefix} - confirmacao_retificacao.pdf"
    generate_confirmation_pdf(prefix, confirmed, remaining, out_name)
    print("Gerado:", out_name)

if __name__ == "__main__":
    main()
