"""
compare_frequencia.py
Entrada:
  - "<numero> - frequencia.pdf"  (geralmente a folha de presença exportada)
  - "<numero> - relatório.pdf"   (outro PDF com informações de ocorrência)
Saída:
  - "<numero> - conferencia.pdf" (PDF de conferência com diferenças + texto de memorando para retificação)
  - "<numero> - conferencia.csv" (tabela com diferenças, útil para inspeção)
Como usar:
  python compare_frequencia.py 123
  (assume que existe "123 - frequencia.pdf" e "123 - relatório.pdf" no mesmo diretório)
"""

import os
import re
import sys
import pdfplumber
import pandas as pd
from datetime import datetime, date, timedelta
from dateutil.parser import parse as dtparse
from dateutil.relativedelta import relativedelta
from fuzzywuzzy import process
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

# ---------- Configurações / mapeamento ----------
# Campos canônicos que vamos usar internamente:
CANON_FIELDS = ["nro_funcional", "nome", "descricao", "data_inicial", "qtde"]

# Possíveis nomes de colunas nos PDFs (português)
POSSIBLE_NAMES = {
    "nro_funcional": ["Nro funcional", "Funcionário", "Matricula", "Nº funcional", "Nº", "Nro"],
    "nome": ["Nome", "Pessoa", "Colaborador", "Funcionario"],
    "descricao": ["Ocorrência", "Descrição", "Motivo", "Tipo"],
    "data_inicial": ["Data", "Data inicial", "Inicio", "Data de início"],
    "qtde": ["Quantidade", "Qtde dias", "Qtde", "Dias", "Qtd"]
}

# palavras-chave que indicam "minutos perdidos" (caso-insens)
MINUTOS_KEYWORDS = ["minut", "minutos", "min"]  # verifica substrings


# ---------- auxiliares de data ----------
def month_bounds_from_text(text):
    """
    Tenta extrair mês/ano de referência dentro do texto do PDF.
    Formatos esperados (exemplos): "Referência: outubro/2025", "Mês de: Maio 2025", "Janeiro/2025"
    Retorna (first_day, last_day) como datetime.date, ou None se não encontrar.
    """
    text = text.lower()
    # lista meses em pt-br
    meses = {
        "janeiro":1,"fevereiro":2,"março":3,"marco":3,"abril":4,"maio":5,"junho":6,
        "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12
    }
    # patterns como "outubro/2025" ou "outubro de 2025" ou "mês de outubro de 2025"
    m = re.search(r'(' + '|'.join(meses.keys()) + r')[\s\/de,]*(\d{4})', text)
    if m:
        mes = meses[m.group(1)]
        ano = int(m.group(2))
        first = date(ano, mes, 1)
        last = (first + relativedelta(months=1) - timedelta(days=1))
        return first, last
    # pattern "mm/yyyy"
    m2 = re.search(r'(\b\d{1,2})[\/\-](\d{4})', text)
    if m2:
        mm = int(m2.group(1))
        yy = int(m2.group(2))
        if 1 <= mm <= 12:
            first = date(yy, mm, 1)
            last = (first + relativedelta(months=1) - timedelta(days=1))
            return first, last
    return None


def overlap_days(start, days_count, month_start, month_end):
    """
    Dado uma data inicial 'start' (datetime.date) e qtde de dias consecutivos,
    calcula quantos desses dias caem dentro do intervalo [month_start, month_end].
    """
    if isinstance(start, datetime):
        start = start.date()
    period_start = start
    period_end = start + timedelta(days=int(days_count)-1)
    a = max(period_start, month_start)
    b = min(period_end, month_end)
    if b < a:
        return 0
    return (b - a).days + 1


# ---------- extração ----------
def extract_tables_from_pdf(path):
    """
    Tenta extrair tabelas usando pdfplumber. Retorna uma lista de dataframes (pandas).
    """
    dfs = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            try:
                tables = page.extract_tables()
            except Exception:
                tables = None
            if tables:
                for table in tables:
                    # transformar em DataFrame com cabeçalho na primeira linha se fizer sentido
                    df = pd.DataFrame(table)
                    # limpezas simples: remover linhas vazias totais
                    df = df.loc[~df.apply(lambda r: r.astype(str).str.strip().replace('nan','').eq('').all(), axis=1)]
                    if df.shape[0] >= 1 and df.shape[1] >= 2:
                        dfs.append(df)
    return dfs


def guess_column_mapping(df):
    """
    Recebe um DataFrame bruto (colunas possivelmente sem nome), tenta detectar cabeçalho
    e mapear colunas para os CANON_FIELDS usando fuzzy matching nas entradas da primeira linha
    ou nomes das colunas.
    Retorna dataframe renomeado (com colunas CANON_FIELDS possivelmente faltando).
    """
    df2 = df.copy()
    # se primeira linha parece ser cabeçalho (conteúdo não numérico em várias colunas) usamos ela
    first_row = df2.iloc[0].astype(str).str.strip().tolist()
    use_first_as_header = sum(1 for v in first_row if re.search(r'[A-Za-zÀ-ú]', v)) >= (len(first_row)/2)
    if use_first_as_header:
        df2.columns = first_row
        df2 = df2.iloc[1:].reset_index(drop=True)
    # se colnames genéricos (0,1,2), converte para str
    cols = [str(c) for c in df2.columns.tolist()]
    mapping = {}
    for canon, candidates in POSSIBLE_NAMES.items():
        # encontrar melhor match entre nomes das colunas e candidatos
        best = None
        best_score = 0
        for col in cols:
            for cand in candidates:
                score = process.extractOne(col, [cand])[1]  # fuzzy score
                if score > best_score:
                    best_score = score
                    best = col
        # se score razoável (>70) mapeia
        if best and best_score >= 70:
            mapping[best] = canon
    # renomear
    df3 = df2.rename(columns=mapping)
    # reduzir só pra colunas canônicas que existam
    keep = [c for c in CANON_FIELDS if c in df3.columns]
    return df3[keep] if keep else df3


def normalize_rows(df):
    """
    Garante que as colunas canônicas existam e conserta tipos (datas, numeros).
    Retorna DataFrame com colunas CANON_FIELDS (preenchendo NaN quando ausente).
    """
    out = pd.DataFrame(columns=CANON_FIELDS)
    for col in CANON_FIELDS:
        if col in df.columns:
            out[col] = df[col].astype(str).str.strip().replace({'nan':''})
        else:
            out[col] = ''
    # tenta parsear data_inicial
    def try_parse_date(s):
        s = str(s).strip()
        if not s:
            return None
        # tenta extrair dd/mm/yyyy ou dd/mm/yy ou yyyy-mm-dd
        s = re.sub(r'[^0-9\/\-\.\s]', ' ', s)
        s = s.strip()
        try:
            return dtparse(s, dayfirst=True).date()
        except Exception:
            return None
    out['data_inicial_parsed'] = out['data_inicial'].apply(try_parse_date)
    # qtde como inteiro quando possível
    def try_int(x):
        try:
            x2 = re.sub(r'[^\d\-]', '', str(x))
            return int(x2) if x2!='' else None
        except:
            return None
    out['qtde_int'] = out['qtde'].apply(try_int)
    return out


# ---------- comparação ----------
def is_minutos(descricao):
    if not descricao: return False
    s = descricao.lower()
    return any(k in s for k in MINUTOS_KEYWORDS)


def build_canonical_table_from_pdf(path):
    """
    Retorna (df_canonical, month_bounds) onde df_canonical tem colunas:
    ['nro_funcional','nome','descricao','data_inicial_parsed','qtde_int','origem_row_text']
    """
    text_all = ""
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            text_all += "\n" + p.extract_text() if p.extract_text() else ""
    # tentar achar mês
    mb = month_bounds_from_text(text_all)
    # tentar extrair tabelas
    dfs = extract_tables_from_pdf(path)
    rows = []
    for df in dfs:
        guessed = guess_column_mapping(df)
        normalized = normalize_rows(guessed)
        for _, r in normalized.iterrows():
            rows.append(r.to_dict())
    # se não extraiu tabelas, tenta extrair linhas por regex do texto
    if not rows:
        # heurística: linhas que tenham um número funcional + nome + data ou qtde
        for line in text_all.splitlines():
            line = line.strip()
            if not line: continue
            # procura algo como "12345 Nome da Pessoa 01/10/2025 3"
            m = re.search(r'(\b\d{3,6}\b)\s+([A-Za-zÀ-ú\-\s]+?)\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\s+(\d{1,3})', line)
            if m:
                rows.append({
                    'nro_funcional': m.group(1),
                    'nome': m.group(2).strip(),
                    'descricao': '',
                    'data_inicial': m.group(3),
                    'data_inicial_parsed': None,
                    'qtde': m.group(4),
                    'qtde_int': int(m.group(4))
                })
    df_all = pd.DataFrame(rows)
    # garantir colunas
    for col in ['nro_funcional','nome','descricao','data_inicial_parsed','qtde_int']:
        if col not in df_all.columns:
            df_all[col] = None
    return df_all, mb, text_all


def compare_tables(df_freq, df_rel, month_start, month_end):
    """
    Algoritmo principal de comparação:
    - Para cada registro (por nro_funcional + descricao) compara qtde de dias (após considerar overlap com mês).
    - Caso a ocorrência seja 'minutos', só checa se existe a ocorrência (presente/ausente).
    Retorna um DataFrame de diferenças.
    """
    # criar chaves simples
    def key_of(r):
        d = (r.get('descricao') or '').strip().lower()
        n = (r.get('nro_funcional') or '').strip()
        return f"{n}||{d}"

    # pre-process: calcular dias no mês para cada registro
    def compute_days_in_month(row):
        start = row.get('data_inicial_parsed')
        qt = row.get('qtde_int')
        # se tiver descricao que contém datas: tenta extrair range "dd/mm a dd/mm" ou "dd/mm/yyyy a dd/mm/yyyy"
        desc = (row.get('descricao') or '')
        # tenta extrair um padrão dd/mm/yyyy ou dd/mm
        if start is None and desc:
            m = re.search(r'(\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)', desc)
            if m:
                try:
                    d = dtparse(m.group(1), dayfirst=True).date()
                    start = d
                except:
                    start = None
        days_in_month = 0
        if start and qt:
            days_in_month = overlap_days(start, qt, month_start, month_end)
        elif qt is not None and not start:
            # sem data, assume todos os dias estão no mês (fallback)
            # mas para segurança, limitar a número de dias do mês
            days_in_month = min(qt, (month_end - month_start).days + 1)
        return int(days_in_month)

    df_freq = df_freq.copy().fillna('')
    df_rel = df_rel.copy().fillna('')

    df_freq['key'] = df_freq.apply(key_of, axis=1)
    df_rel['key'] = df_rel.apply(key_of, axis=1)

    df_freq['dias_no_mes'] = df_freq.apply(compute_days_in_month, axis=1)
    df_rel['dias_no_mes'] = df_rel.apply(compute_days_in_month, axis=1)

    # construir dicionários por chave
    dict_freq = {r['key']: r for _, r in df_freq.iterrows()}
    dict_rel = {r['key']: r for _, r in df_rel.iterrows()}

    keys = set(dict_freq.keys()).union(set(dict_rel.keys()))
    diffs = []
    for k in keys:
        f = dict_freq.get(k)
        r = dict_rel.get(k)
        nro, desc = k.split('||',1)
        desc_display = desc or (r.get('descricao') if r else (f.get('descricao') if f else ''))
        # presença
        presente_freq = f is not None
        presente_rel = r is not None
        dias_freq = f['dias_no_mes'] if f is not None else 0
        dias_rel = r['dias_no_mes'] if r is not None else 0
        minutos = is_minutos(desc_display)
        issue = None
        if minutos:
            # só checamos presença
            if presente_freq and not presente_rel:
                issue = 'Ausente no relatório (minutos)'
            elif not presente_freq and presente_rel:
                issue = 'Ausente na frequência (minutos)'
            else:
                issue = None  # ambos presentes -> OK
        else:
            # checamos quantidade de dias (após overlap)
            if presente_freq and not presente_rel:
                issue = f'Presente na frequência ({dias_freq}d) ausente no relatório'
            elif not presente_freq and presente_rel:
                issue = f'Presente no relatório ({dias_rel}d) ausente na frequência'
            else:
                # ambos presentes -> comparar valores
                if dias_freq != dias_rel:
                    issue = f'Divergência dias: frequência={dias_freq} / relatório={dias_rel}'
                else:
                    issue = None
        diffs.append({
            'nro_funcional': nro,
            'descricao': desc_display,
            'presente_frequencia': presente_freq,
            'presente_relatorio': presente_rel,
            'dias_frequencia': dias_freq,
            'dias_relatorio': dias_rel,
            'minutos': minutos,
            'issue': issue
        })
    df_diffs = pd.DataFrame(diffs)
    # filtrar só divergências
    df_only_issues = df_diffs[df_diffs['issue'].notnull()]
    return df_only_issues.sort_values(['nro_funcional','descricao'])


# ---------- geração de PDF ----------
def generate_conference_pdf(number, diffs_df, month_start, month_end, output_path_pdf, output_csv_path):
    # criar CSV
    diffs_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

    doc = SimpleDocTemplate(output_path_pdf, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    # Título
    story.append(Paragraph(f"Conferência de Frequência - {number}", styles['Title']))
    story.append(Spacer(1,12))
    story.append(Paragraph(f"Referência: {month_start.strftime('%B %Y').capitalize()} ({month_start.isoformat()} a {month_end.isoformat()})", styles['Normal']))
    story.append(Spacer(1,12))

    # Sumário
    tot = len(diffs_df)
    story.append(Paragraph(f"Total de divergências detectadas: {tot}", styles['Normal']))
    story.append(Spacer(1,12))

    # Tabela de divergências
    if tot > 0:
        table_data = [['Nro','Descrição','Minutos?','Dias (Freq)','Dias (Rel)','Observação']]
        for _, r in diffs_df.iterrows():
            table_data.append([
                r['nro_funcional'],
                (r['descricao'] or '')[:60],
                'SIM' if r['minutos'] else 'NÃO',
                r['dias_frequencia'],
                r['dias_relatorio'],
                (r['issue'] or '')
            ])
        t = Table(table_data, colWidths=[50,200,50,60,60,150])
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
            ('GRID',(0,0),(-1,-1),0.25,colors.black),
            ('VALIGN',(0,0),(-1,-1),'TOP'),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold')
        ]))
        story.append(t)
        story.append(Spacer(1,12))
    else:
        story.append(Paragraph("Nenhuma divergência encontrada.", styles['Normal']))
        story.append(Spacer(1,12))

    # Texto do memorando (para copiar/colar no seu sistema)
    story.append(Paragraph("Modelo de Memorando para Solicitação de Retificação de Frequência (copiar/colar):", styles['Heading3']))
    story.append(Spacer(1,6))

    # gerar texto do memorando com as divergências sumarizadas por funcionário
    memo_lines = []
    for nro, group in diffs_df.groupby('nro_funcional'):
        memo_lines.append(f"Funcionário: {nro}")
        for _, r in group.iterrows():
            if r['minutos']:
                memo_lines.append(f" - Ocorrência: {r['descricao']} -> Ajustar presença de minutos (verificar sistema). Observação: {r['issue']}")
            else:
                memo_lines.append(f" - Ocorrência: {r['descricao']} -> Frequência: {r['dias_frequencia']} dias; Relatório: {r['dias_relatorio']}. Observação: {r['issue']}")
        memo_lines.append("")  # linha em branco
    memo_text = "\n".join(memo_lines) if memo_lines else "Não há divergências para retificação."
    # inserir em parágrafos (mantendo linhas)
    for line in memo_text.splitlines():
        story.append(Paragraph(line if line else "&nbsp;", styles['Normal']))

    story.append(Spacer(1,18))
    story.append(Paragraph("Assinatura:", styles['Normal']))
    story.append(Spacer(1,36))
    story.append(Paragraph("__________________________", styles['Normal']))

    doc.build(story)
    print(f"Arquivo de conferência gerado: {output_path_pdf}")
    print(f"CSV auxiliar: {output_csv_path}")


# ---------- MAIN ----------
def main(numero):
    freq_path = f"{numero} - frequencia.pdf"
    rel_path = f"{numero} - relatório.pdf"
    out_pdf = f"{numero} - conferencia.pdf"
    out_csv = f"{numero} - conferencia.csv"

    if not os.path.exists(freq_path) or not os.path.exists(rel_path):
        print("Arquivos de entrada não encontrados. Verifique nomes:", freq_path, rel_path)
        return

    df_freq, mb_freq, text_freq = build_canonical_table_from_pdf(freq_path)
    df_rel, mb_rel, text_rel = build_canonical_table_from_pdf(rel_path)

    # determinar mês de comparação: preferir mês presente em freq, senão relatório, senão heurística
    if mb_freq:
        month_start, month_end = mb_freq
    elif mb_rel:
        month_start, month_end = mb_rel
    else:
        # fallback: use mês atual
        today = date.today()
        month_start = date(today.year, today.month, 1)
        month_end = month_start + relativedelta(months=1) - timedelta(days=1)

    diffs = compare_tables(df_freq, df_rel, month_start, month_end)
    generate_conference_pdf(numero, diffs, month_start, month_end, out_pdf, out_csv)
    print("Pronto.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python compare_frequencia.py <numero>")
    else:
        main(sys.argv[1])
