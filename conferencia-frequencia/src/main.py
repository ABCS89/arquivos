# -*- coding: utf-8 -*-
"""
Uso:
    venv/bin/python src/main.py input/secretaria/2026-04.pdf input/sistema/2026-04.pdf
    venv/bin/python src/main.py input/secretaria/2026-04.pdf input/sistema/2026-04.pdf --mes 2026-04

Gera em output/:
  - conferencia_AAAA-MM.xlsx    (aba "Retificações" + aba "Resumo")
  - retificacoes_AAAA-MM.md     (só é gerado SE houver retificação — pronto
                                  pra colar num e-mail/chamado, no formato
                                  "matrícula - nome - data - X dia(s) -
                                  tipo atual --> tipo correto")
  - retificacoes_AAAA-MM.txt    (mesma lista, em texto puro linha a linha)

A comparação é feita DIA A DIA dentro do mês de referência: cada
ocorrência (da secretaria e do sistema) é expandida em dias individuais,
recortados às bordas do mês, e só entra no relatório o que efetivamente
diverge dentro desse período. Dias consecutivos com o mesmo "de -> para"
são agrupados num único intervalo.
"""
import sys
import os
from datetime import datetime

from extract_secretaria import extrair as extrair_secretaria
from extract_sistema import extrair as extrair_sistema
from compare import comparar_por_dia

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

COR_CABECALHO = "1F4E78"
COR_SEM_REGISTRO_SECRETARIA = "FFF2CC"   # secretaria não tinha nada -> precisa lançar
COR_SEM_REGISTRO_SISTEMA = "D9E1F2"      # sistema não tinha nada -> conferir/lançar lá
COR_TIPO_DIFERENTE = "F8CBAD"            # os dois têm algo, mas tipos diferentes

MESES_PT = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]


def _cor_da_linha(r):
    if r["tipo_secretaria"] == "SEM REGISTRO":
        return COR_SEM_REGISTRO_SECRETARIA
    if r["tipo_sistema"] == "sem registro em sistema":
        return COR_SEM_REGISTRO_SISTEMA
    return COR_TIPO_DIFERENTE


def _texto_data(r):
    if r["data_inicio"] == r["data_fim"]:
        return r["data_inicio"].strftime("%d/%m/%Y")
    return f'{r["data_inicio"].strftime("%d/%m/%Y")} a {r["data_fim"].strftime("%d/%m/%Y")}'


def gerar_excel(retificacoes, caminho_saida, total_sec, total_sis, mes_label):
    wb = Workbook()

    ws = wb.active
    ws.title = "Retificações"
    colunas = ["Matrícula", "Nome", "Data Início", "Data Fim", "Dias",
               "Tipo Atual (a corrigir)", "Tipo Correto"]
    ws.append(colunas)
    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        cel.alignment = Alignment(horizontal="center")

    for r in retificacoes:
        ws.append([
            r["matricula"], r["nome"],
            r["data_inicio"].strftime("%d/%m/%Y"), r["data_fim"].strftime("%d/%m/%Y"),
            r["dias"], r["tipo_secretaria"], r["tipo_sistema"],
        ])
        cor = _cor_da_linha(r)
        for cel in ws[ws.max_row]:
            cel.fill = PatternFill("solid", fgColor=cor)

    larguras = [12, 34, 14, 14, 8, 30, 30]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("Resumo")
    ws2.append([f"Conferência de Frequência — 116 ({mes_label})", ""])
    ws2["A1"].font = Font(bold=True, size=14)
    ws2.append(["Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws2.append(["Registros extraídos (secretaria)", total_sec])
    ws2.append(["Registros extraídos (sistema)", total_sis])
    ws2.append(["Total de retificações", len(retificacoes)])
    ws2.append([])
    ws2.append(["Legenda de cores", ""])
    ws2["A7"].font = Font(bold=True)
    ws2.append(["Amarelo", "secretaria não tinha nada lançado nesse dia"])
    ws2.append(["Azul", "sistema não tem nada nesse dia (conferir se é pra lançar lá)"])
    ws2.append(["Laranja", "os dois têm algo lançado, mas de tipo diferente"])
    ws2.column_dimensions["A"].width = 40
    ws2.column_dimensions["B"].width = 55

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    wb.save(caminho_saida)


def _linha_texto(r):
    return (f'{r["matricula"]} - {r["nome"]} - {_texto_data(r)} - {r["dias"]} '
            f'dia{"s" if r["dias"] != 1 else ""} - {r["tipo_secretaria"]} --> {r["tipo_sistema"]}')


def gerar_txt(retificacoes, caminho_saida):
    linhas = [_linha_texto(r) for r in retificacoes]
    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")


def gerar_markdown(retificacoes, caminho_saida, mes_label, total_sec, total_sis):
    linhas = []
    linhas.append(f"# Retificações de Frequência — 116 ({mes_label})")
    linhas.append("")
    linhas.append(f"- Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    linhas.append(f"- Registros extraídos (secretaria): {total_sec}")
    linhas.append(f"- Registros extraídos (sistema): {total_sis}")
    linhas.append(f"- **Total de retificações: {len(retificacoes)}**")
    linhas.append("")
    linhas.append("Favor retificar as seguintes frequências:")
    linhas.append("")
    for r in retificacoes:
        linhas.append(f"- {_linha_texto(r)}")
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

    retificacoes, (ano_ref, mes_ref) = comparar_por_dia(regs_sec, regs_sis, ano_mes=ano_mes)
    mes_label = f"{MESES_PT[mes_ref]}/{ano_ref}" if ano_ref else "mês não identificado"
    print(f"Mês de referência: {mes_label}")
    print(f"Retificações encontradas: {len(retificacoes)}")

    nome_saida = os.path.splitext(os.path.basename(caminho_sistema))[0]
    caminho_xlsx = os.path.join("output", f"conferencia_{nome_saida}.xlsx")
    gerar_excel(retificacoes, caminho_xlsx, len(regs_sec), len(regs_sis), mes_label)
    print(f"Relatório salvo em: {caminho_xlsx}")

    if retificacoes:
        caminho_md = os.path.join("output", f"retificacoes_{nome_saida}.md")
        gerar_markdown(retificacoes, caminho_md, mes_label, len(regs_sec), len(regs_sis))
        print(f"Lista em markdown salva em: {caminho_md}")

        caminho_txt = os.path.join("output", f"retificacoes_{nome_saida}.txt")
        gerar_txt(retificacoes, caminho_txt)
        print(f"Lista em texto puro salva em: {caminho_txt}")
    else:
        print("Nenhuma retificação — .md/.txt não foram gerados.")


if __name__ == "__main__":
    main()
