"""
aplica_retificacao.py
Entrada:
  - "<numero> - conferencia.pdf"  (gerado pelo script anterior)
  - "<numero> - retificacao_memorando.pdf" (memorando enviado pela chefia com retificações)
Saída:
  - "<numero> - confirmacao_retificacao.pdf" (relatório confirmando o que foi retificado)
  - "<numero> - confirmacao_retificacao.csv" (resumo das alterações)
Como usar:
  python aplica_retificacao.py 123
"""

import os
import re
import sys
import pdfplumber
import pandas as pd
from datetime import datetime, date
from dateutil.parser import parse as dtparse
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# heurísticas para extrair dados do memorando:
# vamos procurar linhas com padrão: <nro> <nome> <descricao> <data_inicial> <qtde>
LINE_PATTERN = re.compile(r'(\b\d{3,6}\b).{0,40}?([A-Za-zÀ-ú\-\s]{3,60}).{0,60}?(\d{1,2}[\/\-]\d{1,2}[\/\-]?\d{2,4}).{0,40}?(\d{1,2,3})')

def extract_text_from_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                text += "\n" + t
    return text

def parse_memorando_retirificacao(text):
    """
    Extrai linhas de retificação do memorando.
    Retorna lista de dicts: [{'nro_funcional':..., 'nome':..., 'descricao':..., 'data_inicial':..., 'qtde':...}, ...]
    Observação: esse parser é heurístico — ajuste regex se o memorando tiver formato diferente.
    """
    entries = []
    # primeiro, tentar encontrar blocos "Matricula: 123 Nome: Fulano ..." ou tabulações
    # tentativa 1: linhas com nro + nome + data + qtde
    for line in text.splitlines():
        line = line.strip()
        if not line: continue
        m = LINE_PATTERN.search(line)
        if m:
            nro = m.group(1)
            nome = m.group(2).strip()
            data = m.group(3).strip()
            qtde = m.group(4).strip()
            # descrição: heurística - pegar conteúdo entre nome e data
            try:
                desc = re.search(re.escape(nome) + r'(.+?)' + re.escape(data), line)
                descricao = desc.group(1).strip() if desc else ''
            except:
                descricao = ''
            entries.append({
                'nro_funcional': nro,
                'nome': nome,
                'descricao': descricao,
                'data_inicial': data,
                'qtde': qtde
            })
    # se nada encontrado, tentar extrair tabelas com pdfplumber e transformá-las
    if not entries:
        # tentar extrair qualquer agrupamento de números e texto
        # fallback: procurar "Nro: 123" patterns
        m_all = re.findall(r'Nro[:\s]*?(\d{3,6}).{0,60}?(?:Descricao[:\s]*([A-Za-zÀ-ú\-\s]+))?.{0,60}?(?:Data[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]?\d{2,4}))?.{0,40}?(?:Qtde[:\s]*(\d{1,2,3}))?', text, flags=re.IGNORECASE)
        for t in m_all:
            nro, desc, data = t[0], (t[1] or ''), (t[2] or '')
            qt = t[3] or ''
            entries.append({'nro_funcional': nro, 'nome':'', 'descricao':desc, 'data_inicial':data, 'qtde':qt})
    return entries

def apply_retificacao(conference_csv_path, entries):
    """
    Aplica as retificações sobre o CSV de conferência (que contém divergências).
    Na prática, usamos o arquivo de conferencia (csv) como base para confirmar que as divergências foram
    corrigidas conforme o memorando. Retorna DataFrame de confirmação.
    """
    if not os.path.exists(conference_csv_path):
        raise FileNotFoundError("CSV de conferencia não encontrado: "+conference_csv_path)
    df_conf = pd.read_csv(conference_csv_path, dtype=str).fillna('')
    # para cada entry do memorando, buscamos linhas correspondentes no df_conf e marcamos como "retificado"
    results = []
    for e in entries:
        nro = e.get('nro_funcional','').strip()
        descricao = (e.get('descricao') or '').strip().lower()
        qt = e.get('qtde','').strip()
        # busca por nro e similaridade em descricao
        candidates = df_conf[df_conf['nro_funcional'].astype(str).str.strip() == nro]
        if candidates.empty:
            results.append({
                'nro_funcional': nro,
                'descricao_memorando': descricao,
                'qtde_memorando': qt,
                'status': 'Funcionário não encontrado nas divergências'
            })
            continue
        # procura linha com descricao similar
        matched = None
        for _, row in candidates.iterrows():
            desc_row = (row.get('descricao') or '').strip().lower()
            if descricao in desc_row or desc_row in descricao or descricao=='':
                matched = row
                break
        if matched is None:
            results.append({
                'nro_funcional': nro,
                'descricao_memorando': descricao,
                'qtde_memorando': qt,
                'status': 'Ocorrência não encontrada - verificar texto'
            })
        else:
            results.append({
                'nro_funcional': nro,
                'descricao_memorando': descricao,
                'qtde_memorando': qt,
                'descricao_conferencia': matched.get('descricao',''),
                'dias_frequencia': matched.get('dias_frequencia',''),
                'dias_relatorio': matched.get('dias_relatorio',''),
                'status': 'Retificado (memorando) - verificar no sistema'
            })
    return pd.DataFrame(results)


def generate_confirmation_pdf(number, conf_df, out_pdf, out_csv):
    conf_df.to_csv(out_csv, index=False, encoding='utf-8-sig')
    doc = SimpleDocTemplate(out_pdf, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph(f"Confirmação de Retificação - {number}", styles['Title']))
    story.append(Spacer(1,12))
    story.append(Paragraph(f"Resumo das retificações recebidas no memorando e conferidas em {date.today().isoformat()}", styles['Normal']))
    story.append(Spacer(1,12))

    # tabela
    table_data = [['Nro','Descrição (memorando)','Qtde (memo)','Descrição (conferência)','Dias Freq','Dias Rel','Status']]
    for _, r in conf_df.iterrows():
        table_data.append([
            r.get('nro_funcional',''),
            (r.get('descricao_memorando') or '')[:60],
            r.get('qtde_memorando',''),
            (r.get('descricao_conferencia') or '')[:60],
            r.get('dias_frequencia',''),
            r.get('dias_relatorio',''),
            r.get('status','')
        ])
    t = Table(table_data, colWidths=[50,170,50,170,50,50,100])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
        ('GRID',(0,0),(-1,-1),0.25,colors.black),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold')
    ]))
    story.append(t)
    doc.build(story)
    print("Confirmacao gerada:", out_pdf, "CSV:", out_csv)


def main(numero):
    conf_pdf = f"{numero} - conferencia.pdf"
    conf_csv = f"{numero} - conferencia.csv"
    memo_pdf = f"{numero} - retificacao_memorando.pdf"
    out_pdf = f"{numero} - confirmacao_retificacao.pdf"
    out_csv = f"{numero} - confirmacao_retificacao.csv"

    if not os.path.exists(conf_pdf) or not os.path.exists(memo_pdf):
        print("Arquivos necessários não encontrados. Verifique:", conf_pdf, memo_pdf)
        return

    text_memo = extract_text_from_pdf(memo_pdf)
    entries = parse_memorando_retirificacao(text_memo)
    if not entries:
        print("Nenhuma entrada extraída do memorando. Verifique o formato do memorando.")
    df_confirm = apply_retificacao(conf_csv, entries)
    generate_confirmation_pdf(numero, df_confirm, out_pdf, out_csv)
    print("Pronto.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python aplica_retificacao.py <numero>")
    else:
        main(sys.argv[1])
