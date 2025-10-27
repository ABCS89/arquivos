#!/usr/bin/env python3
# conferencia_frequencia.py
"""
Uso:
  python conferencia_frequencia.py "106 - frequencia.pdf" "106 - relatório.pdf"

Gera: "106 - conferencia.pdf"
"""

import sys
import re
from pathlib import Path
import pdfplumber
import pandas as pd
from dateutil import parser
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# -------------------------
# Helpers de extração
# -------------------------
def extract_text_lines(pdf_path):
    """Extrai todas as linhas de texto do PDF (lista de strings)"""
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            txt = p.extract_text() or ""
            for l in txt.splitlines():
                lines.append(l.strip())
    return lines

def find_referencia_mes_from_frequencia(lines):
    """
    Procura linha tipo: 'Referente: Setembro/2025' ou 'Referente: 01/09/2025 a 30/09/2025' e retorna (ano, mes)
    """
    for l in lines[:30]:
        m = re.search(r"Referente: *([A-Za-zçÇ]+)[ /]*(\d{4})", l)  # exemplo Setembro/2025
        if m:
            # tenta parsear mês por nome; fallback por número
            try:
                mes_str = m.group(1)
                # tenta converter 'Setembro' -> mês número
                dt = parser.parse(f"1 {mes_str} {m.group(2)}", dayfirst=True)
                return dt.year, dt.month
            except Exception:
                pass
        m2 = re.search(r"Referente:\s*(\d{2})/(\d{2})/(\d{4})\s*a\s*(\d{2})/(\d{2})/(\d{4})", l)
        if m2:
            # usa mês da primeira data
            return int(m2.group(3)), int(m2.group(2))
    # fallback: usa mês atual
    now = datetime.now()
    return now.year, now.month

def parse_frequencia(pdf_path):
    """
    Extrai registros do PDF de frequência para um DataFrame com colunas:
    nro_funcional, nome, ocorrencia, data (data inicial), quantidade (float)
    Observações:
     - O PDF tem blocos com cabeçalho 'Nro Funcional Nome Ocorrência Data Quantidade'
     - Algumas linhas repetem ocorrencias (ex: 'Abono' duas linhas com datas)
    """
    lines = extract_text_lines(pdf_path)
    registros = []
    # procura blocos que começam com padrão de número funcional (xx.xxx-x)
    func_pat = re.compile(r"(\d{2,3}\.\d{3}-\d)\s+(.+?)\s+([A-Za-zÇç ].*?)(?:\s+(\d{2}/\d{2}/\d{4}))?\s*([\d\.,]+)?$")
    # também há linhas onde após o nome aparece "Frequência normal" sem data/quant
    for l in lines:
        m = func_pat.search(l)
        if m:
            nro = m.group(1).strip()
            nome = m.group(2).strip()
            ocorr = (m.group(3) or "").strip()
            data = (m.group(4) or "").strip()
            qtd = (m.group(5) or "").strip()
            # normaliza quantidade para float (troca vírgula por ponto)
            qtd_val = None
            if qtd:
                try:
                    qtd_val = float(qtd.replace(".", "").replace(",", "."))
                except:
                    qtd_val = None
            registros.append({
                "nro_funcional": nro,
                "nome": nome,
                "ocorrencia": ocorr,
                "data": data if data else None,
                "quantidade": qtd_val
            })
        else:
            # linhas que começam com 'Falta 17/09/2025 2,0' (ocorrencia para o funcionário anterior)
            m2 = re.match(r"^(Falta|Abono|Férias|Minutos perdidos|Tratamento de saúde|Doação de sangue|Afastamento|Cedido.*|Auxílio.*|Nojo|Aguardando perícia.*)\s*(\d{2}/\d{2}/\d{4})?\s*([\d\.,]+)?", l, flags=re.I)
            if m2 and registros:
                ocorr = m2.group(1).strip()
                data = m2.group(2) or ""
                qtd = m2.group(3) or ""
                qtd_val = None
                if qtd:
                    try:
                        qtd_val = float(qtd.replace(".", "").replace(",", "."))
                    except:
                        qtd_val = None
                # adiciona novo registro para o último funcionário (porque o PDF desmembra)
                registros.append({
                    "nro_funcional": registros[-1]["nro_funcional"],
                    "nome": registros[-1]["nome"],
                    "ocorrencia": ocorr,
                    "data": data if data else None,
                    "quantidade": qtd_val
                })
    df = pd.DataFrame(registros)
    return df

def parse_relatorio(pdf_path):
    """
    Lê o relatório (formato com colunas: Divisão Funcionário Pessoa Data Inicial Data Final Qtde Dias Descrição)
    Normaliza para as mesmas colunas do DataFrame de frequência:
    => nro_funcional, nome, ocorrencia (Descrição), data (Data Inicial), quantidade (Qtde Dias)
    """
    lines = extract_text_lines(pdf_path)
    registros = []
    # padrão número funcional + nome + data inicial + data final + qtde + descricao (pode quebrar linhas)
    pat = re.compile(r"(\d{2,3}\.\d{3}-\d)\s+([A-ZÇÁÉÍÓÚÂÊÔÃÕ0-9 \-\.]+?)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+([\d\.,]+)\s+(.*)$", flags=re.I)
    for l in lines:
        m = pat.search(l)
        if m:
            nro = m.group(1).strip()
            nome = " ".join(m.group(2).split())
            data_ini = m.group(3)
            data_fim = m.group(4)
            qtd = m.group(5)
            desc = m.group(6).strip()
            qtd_val = None
            try:
                qtd_val = float(qtd.replace(".", "").replace(",", "."))
            except:
                qtd_val = None
            registros.append({
                "nro_funcional": nro,
                "nome": nome,
                "ocorrencia": desc,
                "data": data_ini,
                "quantidade": qtd_val,
                "data_final": data_fim
            })
    df = pd.DataFrame(registros)
    return df

# -------------------------
# Normalização e regras
# -------------------------
def normalize_text(s):
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def monthday_overlap_count(start_date_str, end_date_str, ref_year, ref_month):
    """
    Conta quantos dias do período [start_date, end_date] caem dentro do mês (ref_year, ref_month).
    Retorna int (dias).
    """
    if not start_date_str or not end_date_str:
        return 0
    try:
        s = parser.parse(start_date_str, dayfirst=True).date()
        e = parser.parse(end_date_str, dayfirst=True).date()
    except:
        # se parse falhar, tenta só start
        try:
            s = parser.parse(start_date_str, dayfirst=True).date()
            e = s
        except:
            return 0
    # limite ao mês
    from datetime import date
    month_start = datetime(ref_year, ref_month, 1).date()
    # último dia do mês
    if ref_month == 12:
        month_end = datetime(ref_year+1, 1, 1).date() - pd.Timedelta(days=1)
    else:
        month_end = datetime(ref_year, ref_month+1, 1).date() - pd.Timedelta(days=1)
    # overlap
    s2 = max(s, month_start)
    e2 = min(e, month_end)
    if e2 < s2:
        return 0
    return (e2 - s2).days + 1

def compare_dataframes(df_freq, df_rel, ref_year, ref_month):
    """
    Compara os dois DataFrames e retorna:
      - matches (records that align)
      - diffs (records present in one but different in the other)
    Regras especiais:
      - 'Minutos perdidos' : contabiliza somente a ocorrência — não considera divergencia de qtde/data.
      - No relatório, se a data inicial for em mês anterior, ajustamos contagem para apenas os dias do mês de referência (usando monthday_overlap_count).
    """
    # prepara chaves
    dff = df_freq.copy()
    dfr = df_rel.copy()
    # normalize text fields
    for col in ["ocorrencia", "nome", "nro_funcional", "data"]:
        if col in dff.columns:
            dff[col] = dff[col].apply(normalize_text)
    for col in ["ocorrencia", "nome", "nro_funcional", "data"]:
        if col in dfr.columns:
            dfr[col] = dfr[col].apply(normalize_text)
    # expand rel.quantidade to be only days inside month (if data_final present)
    if "data_final" in dfr.columns:
        dfr["quantidade_mes"] = dfr.apply(
            lambda r: monthday_overlap_count(r.get("data", ""), r.get("data_final", ""), ref_year, ref_month), axis=1
        )
    else:
        dfr["quantidade_mes"] = dfr["quantidade"].fillna(0)
    # create simple key for grouping: (nro_funcional, ocorrencia, data) but note ocorrencia text mismatch—so we compare fuzzy-ish:
    # We'll do matching by nro_funcional + nome + ocorrencia normalized (case-insensitive substring)
    matches = []
    diffs = []
    # index rel by nro_funcional
    rel_map = {}
    for _, r in dfr.iterrows():
        key = r["nro_funcional"]
        rel_map.setdefault(key, []).append(r.to_dict())
    # iterate freq
    used_rel = set()
    for _, f in dff.iterrows():
        key = f["nro_funcional"]
        candidates = rel_map.get(key, [])
        matched = False
        for idx, c in enumerate(candidates):
            # regra 'Minutos Perdidos' (texto pode variar, mas vamos normalizar)
            if "minut" in c["ocorrencia"].lower() or "minut" in f["ocorrencia"].lower():
                # match by functional number and ocorrencia containing 'minut' -> always consider matched (count occurrence only)
                matches.append((
                    f.to_dict(),
                    c,
                    "MINUTOS_PERDIDOS_MATCH"
                ))
                matched = True
                used_rel.add((key, idx))
                break
            # otherwise compare: ocorrencia substring or vice-versa OR igualdade de quantidades após ajuste de mês
            occ_f = f["ocorrencia"].lower()
            occ_r = c["ocorrencia"].lower()
            # quantidade freq
            qf = f["quantidade"] if pd.notna(f["quantidade"]) else 0
            qr = c.get("quantidade_mes", 0) if c.get("quantidade_mes") is not None else (c.get("quantidade") or 0)
            # data compare: compare data strings if present
            df_data = f.get("data") or ""
            dr_data = c.get("data") or ""
            same_occ = (occ_f in occ_r) or (occ_r in occ_f) or (occ_f == occ_r)
            same_q = (qf == qr)
            same_data = (df_data == dr_data) or (df_data == "" and dr_data == "")
            if same_occ and (same_q and same_data):
                matches.append((f.to_dict(), c, "FULL_MATCH"))
                matched = True
                used_rel.add((key, idx))
                break
            # if occurrence names differ but numbers differ -> report diff
        if not matched:
            diffs.append( ("FREQUENCIA_ONLY", f.to_dict(), None) )
    # now check rel entries not used -> they exist in rel but not in freq
    for key, lst in rel_map.items():
        for idx, c in enumerate(lst):
            if (key, idx) not in used_rel:
                diffs.append( ("RELATORIO_ONLY", None, c) )
    return matches, diffs

# -------------------------
# Relatório PDF de saída
# -------------------------
def generate_conferencia_pdf(number_prefix, matches, diffs, output_path, ref_year, ref_month):
    c = canvas.Canvas(str(output_path), pagesize=A4)
    w, h = A4
    margin = 15*mm
    y = h - margin
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin, y, f"Conferência de Frequência - {number_prefix}")
    y -= 12
    c.setFont("Helvetica", 10)
    c.drawString(margin, y, f"Referente: {ref_month:02d}/{ref_year}")
    y -= 18

    # Diferenças - resumo
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin, y, "Resumo de diferenças")
    y -= 14
    c.setFont("Helvetica", 9)
    if not diffs:
        c.drawString(margin, y, "Nenhuma diferença encontrada.")
        y -= 12
    else:
        for tag, frec, rel in diffs:
            if y < 60:
                c.showPage(); y = h - margin
            if tag == "FREQUENCIA_ONLY":
                txt = f"- Registro em FREQUÊNCIA presente e não encontrado no RELATÓRIO: {frec['nro_funcional']} | {frec['nome']} | {frec['ocorrencia']} | {frec.get('data','')} | {frec.get('quantidade', '')}"
            else:
                txt = f"- Registro em RELATÓRIO presente e não encontrado na FREQUÊNCIA: {rel['nro_funcional']} | {rel['nome']} | {rel['ocorrencia']} | {rel.get('data','')} | qtde_mes={rel.get('quantidade_mes', rel.get('quantidade',''))}"
            c.drawString(margin, y, txt)
            y -= 12

    # Memorando (texto pronto para solicitar retificação)
    if y < 140:
        c.showPage(); y = h - margin
    y -= 8
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin, y, "Memorando padrão para solicitação de retificação (copiar/colar):")
    y -= 14
    c.setFont("Helvetica", 9)
    memo_lines = []
    memo_lines.append("Ao(À) Senhor(a) Diretor(a) de Recursos Humanos,")
    memo_lines.append("")
    memo_lines.append(f"Solicito a retificação da frequência referente ao mês {ref_month:02d}/{ref_year} para os seguintes servidores listados abaixo, conforme divergências encontradas entre o arquivo de Frequência e o Relatório de Ocorrências.")
    memo_lines.append("")
    # listar itens em formato colado
    for tag, frec, rel in diffs:
        if tag == "FREQUENCIA_ONLY":
            memo_lines.append(f"- {frec['nro_funcional']} | {frec['nome']} | Ocorrência (na frequência): {frec['ocorrencia']} | Data: {frec.get('data','')} | Qtde: {frec.get('quantidade','')}")
        else:
            # rel not in freq
            memo_lines.append(f"- {rel['nro_funcional']} | {rel['nome']} | Ocorrência (no relatório): {rel['ocorrencia']} | Data inicial: {rel.get('data','')} | Qtde no mês: {rel.get('quantidade_mes','')}")
    memo_lines.append("")
    memo_lines.append("Observação: 'Minutos Perdidos' foram apenas contabilizados como ocorrência (não se exige ajuste de quantidade quando houver divergência).")
    memo_lines.append("")
    memo_lines.append("Atenciosamente,")
    memo_lines.append("") 
    memo_lines.append("[Nome do solicitante] - [Cargo] - [Unidade]")

    for ml in memo_lines:
        if y < 50:
            c.showPage(); y = h - margin
        c.drawString(margin, y, ml)
        y -= 12

    c.showPage()
    c.save()


# -------------------------
# Main CLI
# -------------------------
def main():
    if len(sys.argv) < 3:
        print("Uso: python conferencia_frequencia.py <numero - frequencia.pdf> <numero - relatório.pdf>")
        sys.exit(1)
    freq_path = Path(sys.argv[1])
    rel_path = Path(sys.argv[2])
    # extrai numero prefix pelo nome do arquivo (antes do ' - ')
    number_prefix = freq_path.name.split(" - ")[0]
    lines_freq = extract_text_lines(freq_path)
    ref_year, ref_month = find_referencia_mes_from_frequencia(lines_freq)
    print(f"Detectado mês referência: {ref_month:02d}/{ref_year}")
    df_freq = parse_frequencia(freq_path)
    df_rel = parse_relatorio(rel_path)
    matches, diffs = compare_dataframes(df_freq, df_rel, ref_year, ref_month)
    out_name = f"{number_prefix} - conferencia.pdf"
    generate_conferencia_pdf(number_prefix, matches, diffs, out_name, ref_year, ref_month)
    print("Gerado:", out_name)

if __name__ == "__main__":
    main()
