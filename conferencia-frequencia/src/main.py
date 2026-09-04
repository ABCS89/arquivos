# -*- coding: utf-8 -*-
"""
Uso:
    venv/bin/python src/main.py input/secretaria/2026-04.pdf input/sistema/2026-04.pdf
    venv/bin/python src/main.py input/secretaria/2026-04.pdf input/sistema/2026-04.pdf --mes 2026-04

Gera em output/:
  - conferencia_AAAA-MM.xlsx  (aba "Divergências" + aba "Resumo")
  - divergencias_AAAA-MM.md   (só é gerado SE houver alguma divergência)

A comparação considera somente os dias que caem dentro do mês de
referência (ver src/compare.py).
"""
import sys
import os
from datetime import datetime

from extract_secretaria import extrair as extrair_secretaria
from extract_sistema import extrair as extrair_sistema
from compare import comparar

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

COR_CABECALHO = "1F4E78"
COR_SO_SECRETARIA = "FFF2CC"
COR_SO_SISTEMA = "D9E1F2"
COR_QTDE_DIVERGENTE = "F8CBAD"

MESES_PT = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]


def gerar_excel(divergencias, caminho_saida, total_sec, total_sis, mes_label):
    wb = Workbook()

    # --- aba Divergências ---
    ws = wb.active
    ws.title = "Divergências"
    colunas = ["Matrícula", "Nome", "Tipo de Ocorrência", "Dias (Secretaria)",
               "Dias (Sistema)", "Diferença", "Situação"]
    ws.append(colunas)
    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        cel.alignment = Alignment(horizontal="center")

    cor_por_situacao = {
        "Só na secretaria (sistema não tem)": COR_SO_SECRETARIA,
        "Só no sistema (secretaria não tem)": COR_SO_SISTEMA,
        "Quantidade de dias divergente": COR_QTDE_DIVERGENTE,
    }

    for d in divergencias:
        ws.append([
            d["matricula"], d["nome"], d["tipo_ocorrencia"],
            d["dias_secretaria"], d["dias_sistema"], d["diferenca"], d["situacao"],
        ])
        cor = cor_por_situacao.get(d["situacao"])
        if cor:
            for cel in ws[ws.max_row]:
                cel.fill = PatternFill("solid", fgColor=cor)

    larguras = [12, 34, 26, 16, 14, 12, 32]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # --- aba Resumo ---
    ws2 = wb.create_sheet("Resumo")
    ws2.append([f"Conferência de Frequência — 116 ({mes_label})", ""])
    ws2["A1"].font = Font(bold=True, size=14)
    ws2.append(["Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws2.append(["Registros extraídos (secretaria)", total_sec])
    ws2.append(["Registros extraídos (sistema)", total_sis])
    ws2.append(["Total de divergências", len(divergencias)])
    ws2.append([])
    ws2.append(["Situação", "Quantidade"])
    ws2["A7"].font = Font(bold=True)
    ws2["B7"].font = Font(bold=True)
    contagem = {}
    for d in divergencias:
        contagem[d["situacao"]] = contagem.get(d["situacao"], 0) + 1
    for situacao, qtd in contagem.items():
        ws2.append([situacao, qtd])
    ws2.column_dimensions["A"].width = 40
    ws2.column_dimensions["B"].width = 16

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    wb.save(caminho_saida)


def gerar_markdown(divergencias, caminho_saida, mes_label, total_sec, total_sis):
    """Gera um .md com as divergências, agrupadas por situação. Só deve ser
    chamado quando divergencias não está vazia."""
    linhas = []
    linhas.append(f"# Conferência de Frequência — 116 ({mes_label})")
    linhas.append("")
    linhas.append(f"- Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    linhas.append(f"- Registros extraídos (secretaria): {total_sec}")
    linhas.append(f"- Registros extraídos (sistema): {total_sis}")
    linhas.append(f"- **Total de divergências: {len(divergencias)}**")
    linhas.append("")

    grupos = {}
    for d in divergencias:
        grupos.setdefault(d["situacao"], []).append(d)

    for situacao, itens in grupos.items():
        linhas.append(f"## {situacao} ({len(itens)})")
        linhas.append("")
        linhas.append("| Matrícula | Nome | Tipo de Ocorrência | Dias Secretaria | Dias Sistema | Diferença |")
        linhas.append("|---|---|---|---|---|---|")
        for d in itens:
            ds = "-" if d["dias_secretaria"] is None else int(d["dias_secretaria"])
            di = "-" if d["dias_sistema"] is None else int(d["dias_sistema"])
            linhas.append(
                f"| {d['matricula']} | {d['nome']} | {d['tipo_ocorrencia']} | "
                f"{ds} | {di} | {d['diferenca']:+.0f} |"
            )
        linhas.append("")

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 2:
        print("Uso: python src/main.py <pdf_secretaria> <pdf_sistema> [--mes AAAA-MM]")
        sys.exit(1)

    caminho_secretaria, caminho_sistema = args

    ano_mes = None
    if "--mes" in sys.argv:
        valor = sys.argv[sys.argv.index("--mes") + 1]
        ano_str, mes_str = valor.split("-")
        ano_mes = (int(ano_str), int(mes_str))

    print(f"Lendo secretaria: {caminho_secretaria}")
    regs_sec, nao_reconhecidas = extrair_secretaria(caminho_secretaria)
    print(f"  -> {len(regs_sec)} registros")
    if nao_reconhecidas:
        print(f"  -> {len(nao_reconhecidas)} linha(s) não reconhecida(s) "
              f"(normalmente ruído do PDF; rode extract_secretaria.py direto "
              f"para revisar se quiser conferir)")

    print(f"Lendo sistema: {caminho_sistema}")
    regs_sis = extrair_sistema(caminho_sistema)
    print(f"  -> {len(regs_sis)} registros")

    divergencias, (ano_ref, mes_ref) = comparar(regs_sec, regs_sis, ano_mes=ano_mes)
    mes_label = f"{MESES_PT[mes_ref]}/{ano_ref}" if ano_ref else "mês não identificado"
    print(f"Mês de referência: {mes_label}")
    print(f"Divergências encontradas: {len(divergencias)}")

    nome_saida = os.path.splitext(os.path.basename(caminho_sistema))[0]
    caminho_xlsx = os.path.join("output", f"conferencia_{nome_saida}.xlsx")
    gerar_excel(divergencias, caminho_xlsx, len(regs_sec), len(regs_sis), mes_label)
    print(f"Relatório salvo em: {caminho_xlsx}")

    if divergencias:
        caminho_md = os.path.join("output", f"divergencias_{nome_saida}.md")
        gerar_markdown(divergencias, caminho_md, mes_label, len(regs_sec), len(regs_sis))
        print(f"Resumo em markdown salvo em: {caminho_md}")
    else:
        print("Nenhuma divergência — .md não foi gerado.")


if __name__ == "__main__":
    main()
