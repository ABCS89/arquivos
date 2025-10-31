# confirm_fallback.py
import sys, csv
from pathlib import Path
from reportlab.lib.pagesizes import A4

def load_csv(path):
    with open(path, newline='', encoding='utf-8', errors='ignore') as f:
        reader = csv.DictReader(f)
        return list(reader)

def gerar_html_confirm(ret, out):
    html = "<html><head><meta charset='utf-8'></head><body>"
    html += "<h1>Confirmação de Retificações</h1>"
    if not ret:
        html += "<p>Nenhuma retificação encontrada.</p>"
    else:
        html += "<table border='1'><tr><th>Nro</th><th>Nome</th><th>Desc</th><th>DataIni</th><th>DataFim</th><th>Qt</th></tr>"
        for r in ret:
            html += "<tr>"
            html += f"<td>{r.get('Funcionário') or r.get('NroFuncional') or ''}</td>"
            html += f"<td>{r.get('Pessoa') or r.get('Nome') or ''}</td>"
            html += f"<td>{r.get('Descrição') or r.get('Ocorrencia') or ''}</td>"
            html += f"<td>{r.get('DataInicial') or ''}</td>"
            html += f"<td>{r.get('DataFinal') or ''}</td>"
            html += f"<td>{r.get('QtdeDias') or r.get('Quantidade') or ''}</td>"
            html += "</tr>"
        html += "</table>"
    html += "</body></html>"
    open(out,"w",encoding="utf-8").write(html)
    print("Gerado:", out)

def main():
    if len(sys.argv)<3:
        print("Uso: python confirm_fallback.py conferencia.html retificacao.csv"); return
    ret = load_csv(sys.argv[2])
    numero = Path(sys.argv[1]).stem.split(" - ")[0] if " - " in Path(sys.argv[1]).stem else Path(sys.argv[1]).stem
    gerar_html_confirm(ret, f"{numero} - confirmacao_retificacao.html")

if __name__ == "__main__":
    main()
