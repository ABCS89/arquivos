# compare_fallback.py
# Uso: python compare_fallback.py "106 - relatorio.txt" "106 - frequencia.csv"
import sys, csv, re
from datetime import datetime, date, timedelta

def texto_para_data(s):
    return datetime.strptime(s.strip(), "%d/%m/%Y").date()

def month_range_from_text(txt):
    m = re.search(r"Referente:\s*([\d/]+)\s*a\s*([\d/]+)", txt, re.I)
    if m:
        return texto_para_data(m.group(1)), texto_para_data(m.group(2))
    today = date.today()
    first = date(today.year, today.month, 1)
    if today.month==12:
        last = date(today.year,12,31)
    else:
        last = date(today.year, today.month+1,1) - timedelta(days=1)
    return first,last

def parse_relatorio_txt(path):
    txt = open(path, encoding="utf-8", errors="ignore").read()
    start,end = month_range_from_text(txt)
    lines = [l.strip() for l in txt.splitlines() if l.strip()]
    rows = []
    for ln in lines:
        m = re.search(r"(\d{1,2}\.\d{3}-\d)\s+([A-ZÀ-ÿ0-9\s\.\-]+?)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(.+)", ln, re.I)
        if m:
            d1 = texto_para_data(m.group(3)); d2 = texto_para_data(m.group(4))
            rows.append({"Nro":m.group(1),"Nome":m.group(2),"D1":d1,"D2":d2,"Qt":int(m.group(5)),"Desc":m.group(6)})
    return rows, start, end

def load_freq_csv(path):
    with open(path, newline='', encoding='utf-8', errors='ignore') as f:
        reader = csv.DictReader(f)
        return list(reader)

def overlap_days(a,b,c,d):
    s = max(a,c); e = min(b,d)
    if e < s: return 0
    return (e - s).days + 1

def main():
    if len(sys.argv)<3:
        print("Uso: python compare_fallback.py relatorio.txt frequencia.csv"); return
    rel, freq = sys.argv[1], sys.argv[2]
    rel_rows, ms, me = parse_relatorio_txt(rel)
    freq_rows = load_freq_csv(freq)
    diffs = []
    for r in rel_rows:
        qt_rel = overlap_days(r["D1"], r["D2"], ms, me)
        # buscar no csv
        matches = [fr for fr in freq_rows if str(r["Nro"].split()[0]) in (fr.get("Nro funcional","") or fr.get("NroFuncional","") or fr.get("Funcionário","")) or (r["Nome"].split()[0].upper() in (fr.get("Nome","").upper() if fr.get("Nome") else ""))]
        qtd_freq=0
        for m in matches:
            try:
                qtd_freq += int(float(m.get("Quantidade", m.get("Qtde",0) or 0)))
            except:
                pass
        if "minutos" in r["Desc"].lower():
            if qtd_freq==0:
                diffs.append((r, qt_rel, qtd_freq))
        else:
            if qtd_freq != qt_rel:
                diffs.append((r, qt_rel, qtd_freq))
    # gerar HTML
    numero = Path(rel).stem.split(" - ")[0] if " - " in Path(rel).stem else Path(rel).stem
    html = "<html><head><meta charset='utf-8'><title>Conferência</title></head><body>"
    html += f"<h1>Conferência {numero}</h1>"
    html += f"<p>Referente: {ms.strftime('%d/%m/%Y')} a {me.strftime('%d/%m/%Y')}</p>"
    if not diffs:
        html += "<p>Nenhuma divergência encontrada.</p>"
    else:
        html += "<table border='1' cellpadding='4' cellspacing='0'><tr><th>Nro</th><th>Nome</th><th>Desc</th><th>QtRel</th><th>QtFreq</th></tr>"
        for d,qr,qf in diffs:
            html += f"<tr><td>{d['Nro']}</td><td>{d['Nome']}</td><td>{d['Desc']}</td><td>{qr}</td><td>{qf}</td></tr>"
        html += "</table>"
    html += "</body></html>"
    out = f"{numero} - conferencia.html"
    open(out,"w",encoding="utf-8").write(html)
    print("Gerado:", out)

if __name__ == "__main__":
    from pathlib import Path
    main()
