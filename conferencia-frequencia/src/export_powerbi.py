# -*- coding: utf-8 -*-
"""
export_powerbi.py - Consolidador de Dados de Monitoramento para o Power BI

Varre todas as pastas mensais em output/*/monitoramento/*.xlsx e gera
uma base de dados consolidada, padronizada e relacional, pronta para
importação direta no Power BI ou ferramentas de BI.

Saída em:
  - output/powerbi/monitoramento_consolidado.xlsx (com abas: Atrasos, Faltas, Resumo_Secretarias)
  - output/powerbi/atrasos.csv
  - output/powerbi/faltas.csv
  - output/powerbi/resumo_secretarias.csv
"""
import re
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
POWERBI_DIR = OUTPUT_DIR / "powerbi"


def extrair_codigo_e_nome(nome_arquivo):
    """Extrai código e nome da secretaria do nome do arquivo (ex: 'monitoramento_103 - PROCURADORIA GERAL.xlsx')."""
    stem = Path(nome_arquivo).stem
    stem = re.sub(r"^monitoramento_", "", stem, flags=re.IGNORECASE)
    m = re.match(r"^(\d{3})\s*-\s*(.+)$", stem)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m_num = re.match(r"^(\d{3})\b", stem)
    if m_num:
        return m_num.group(1).strip(), stem
    return "", stem


def consolidar_monitoramento():
    """Varre todas as pastas em output/<mes>/monitoramento/*.xlsx e unifica em tabelas."""
    POWERBI_DIR.mkdir(parents=True, exist_ok=True)

    linhas_atrasos = []
    linhas_faltas = []
    linhas_resumo = []

    # Procura pastas mensais em output (ex: output/2026-08/monitoramento)
    # Também suporta se houver a pasta legada output/monitoramento
    pastas_monitoramento = []
    for item in sorted(OUTPUT_DIR.iterdir()):
        if not item.is_dir() or item.name == "powerbi":
            continue
        monit_sub = item / "monitoramento"
        if monit_sub.is_dir():
            pastas_monitoramento.append((item.name, monit_sub))
        elif item.name == "monitoramento":
            pastas_monitoramento.append(("geral", item))

    print(f"[Power BI] Pastas de monitoramento encontradas: {len(pastas_monitoramento)}")

    for mes_id, pasta in pastas_monitoramento:
        arquivos = sorted(pasta.glob("monitoramento_*.xlsx"))
        print(f"  -> Lendo lote '{mes_id}': {len(arquivos)} arquivo(s)")

        for arq in arquivos:
            cod_sec, nome_sec = extrair_codigo_e_nome(arq.name)
            try:
                wb = load_workbook(arq, data_only=True)
            except Exception as e:
                print(f"     [AVISO] Falha ao abrir {arq.name}: {e}")
                continue

            # 1. Ler Aba Resumo
            mes_label = mes_id
            qtd_serv_atraso = 0
            tot_minutos = 0
            qtd_serv_falta = 0
            tot_dias_falta = 0

            if "Resumo" in wb.sheetnames:
                ws_res = wb["Resumo"]
                titulo_res = ws_res.cell(row=1, column=1).value or ""
                m_label = re.search(r"\((.+?)\)", str(titulo_res))
                if m_label:
                    mes_label = m_label.group(1).strip()

                for row in range(3, ws_res.max_row + 1):
                    ind = str(ws_res.cell(row=row, column=1).value or "").strip().lower()
                    val = ws_res.cell(row=row, column=2).value
                    if "servidores com atrasos" in ind:
                        try: qtd_serv_atraso = int(val)
                        except: pass
                    elif "total de minutos perdidos" in ind:
                        m_val = re.search(r"(\d+)\s*min", str(val))
                        if m_val:
                            tot_minutos = int(m_val.group(1))
                        else:
                            try: tot_minutos = int(val)
                            except: pass
                    elif "servidores com faltas" in ind:
                        try: qtd_serv_falta = int(val)
                        except: pass
                    elif "total de faltas acumuladas" in ind:
                        m_dias = re.search(r"(\d+)", str(val))
                        if m_dias:
                            tot_dias_falta = int(m_dias.group(1))

                linhas_resumo.append({
                    "AnoMes": mes_id,
                    "MesReferencia": mes_label,
                    "CodigoSecretaria": cod_sec,
                    "Secretaria": nome_sec,
                    "OrgaoCompleto": f"{cod_sec} - {nome_sec}" if cod_sec else nome_sec,
                    "QtdServidoresAtraso": qtd_serv_atraso,
                    "TotalMinutosPerdidos": tot_minutos,
                    "TotalHorasPerdidas": round(tot_minutos / 60.0, 2),
                    "QtdServidoresFalta": qtd_serv_falta,
                    "TotalDiasFalta": tot_dias_falta,
                })

            # 2. Ler Aba Atrasos (Minutos)
            nome_aba_atrasos = None
            for s in wb.sheetnames:
                if "atraso" in s.lower() or "minuto" in s.lower():
                    nome_aba_atrasos = s
                    break

            if nome_aba_atrasos:
                ws_atr = wb[nome_aba_atrasos]
                for r in range(2, ws_atr.max_row + 1):
                    mat = ws_atr.cell(row=r, column=1).value
                    if not mat:
                        continue
                    nome = ws_atr.cell(row=r, column=2).value or ""
                    min_val = ws_atr.cell(row=r, column=3).value or 0
                    equiv = ws_atr.cell(row=r, column=4).value or ""
                    det = ws_atr.cell(row=r, column=5).value or ""
                    status = ws_atr.cell(row=r, column=6).value or ""

                    try: min_int = int(min_val)
                    except: min_int = 0

                    linhas_atrasos.append({
                        "AnoMes": mes_id,
                        "MesReferencia": mes_label,
                        "CodigoSecretaria": cod_sec,
                        "Secretaria": nome_sec,
                        "OrgaoCompleto": f"{cod_sec} - {nome_sec}" if cod_sec else nome_sec,
                        "Matricula": str(mat).strip(),
                        "NomeServidor": str(nome).strip(),
                        "MinutosAcumulados": min_int,
                        "HorasEquivalentes": round(min_int / 60.0, 2),
                        "EquivalenciaTexto": str(equiv).strip(),
                        "OcorrenciasDatas": str(det).strip(),
                        "SituacaoSistema": str(status).strip(),
                    })

            # 3. Ler Aba Faltas (Dias)
            nome_aba_faltas = None
            for s in wb.sheetnames:
                if "falta" in s.lower():
                    nome_aba_faltas = s
                    break

            if nome_aba_faltas:
                ws_flt = wb[nome_aba_faltas]
                for r in range(2, ws_flt.max_row + 1):
                    mat = ws_flt.cell(row=r, column=1).value
                    if not mat:
                        continue
                    nome = ws_flt.cell(row=r, column=2).value or ""
                    dias_val = ws_flt.cell(row=r, column=3).value or 0
                    tipos = ws_flt.cell(row=r, column=4).value or ""
                    det = ws_flt.cell(row=r, column=5).value or ""
                    status = ws_flt.cell(row=r, column=6).value or ""

                    try: dias_int = int(dias_val)
                    except: dias_int = 0

                    linhas_faltas.append({
                        "AnoMes": mes_id,
                        "MesReferencia": mes_label,
                        "CodigoSecretaria": cod_sec,
                        "Secretaria": nome_sec,
                        "OrgaoCompleto": f"{cod_sec} - {nome_sec}" if cod_sec else nome_sec,
                        "Matricula": str(mat).strip(),
                        "NomeServidor": str(nome).strip(),
                        "DiasFalta": dias_int,
                        "TiposFalta": str(tipos).strip(),
                        "OcorrenciasDatas": str(det).strip(),
                        "SituacaoSistema": str(status).strip(),
                    })

    # Gravar Excel Consolidado
    caminho_xlsx = POWERBI_DIR / "monitoramento_consolidado.xlsx"
    wb_out = Workbook()

    cor_header = "1F4E78"
    font_header = Font(bold=True, color="FFFFFF")
    fill_header = PatternFill("solid", fgColor=cor_header)

    # Aba 1: Resumo_Secretarias
    ws1 = wb_out.active
    ws1.title = "Resumo_Secretarias"
    if linhas_resumo:
        headers1 = list(linhas_resumo[0].keys())
        ws1.append(headers1)
        for cel in ws1[1]:
            cel.font = font_header
            cel.fill = fill_header
            cel.alignment = Alignment(horizontal="center")
        for item in linhas_resumo:
            ws1.append(list(item.values()))
        for col_idx in range(1, len(headers1) + 1):
            ws1.column_dimensions[get_column_letter(col_idx)].width = 22

    # Aba 2: Fato_Atrasos
    ws2 = wb_out.create_sheet("Fato_Atrasos")
    if linhas_atrasos:
        headers2 = list(linhas_atrasos[0].keys())
        ws2.append(headers2)
        for cel in ws2[1]:
            cel.font = font_header
            cel.fill = fill_header
            cel.alignment = Alignment(horizontal="center")
        for item in linhas_atrasos:
            ws2.append(list(item.values()))
        for col_idx in range(1, len(headers2) + 1):
            ws2.column_dimensions[get_column_letter(col_idx)].width = 22

    # Aba 3: Fato_Faltas
    ws3 = wb_out.create_sheet("Fato_Faltas")
    if linhas_faltas:
        headers3 = list(linhas_faltas[0].keys())
        ws3.append(headers3)
        for cel in ws3[1]:
            cel.font = font_header
            cel.fill = fill_header
            cel.alignment = Alignment(horizontal="center")
        for item in linhas_faltas:
            ws3.append(list(item.values()))
        for col_idx in range(1, len(headers3) + 1):
            ws3.column_dimensions[get_column_letter(col_idx)].width = 22

    wb_out.save(caminho_xlsx)
    print(f"\n[OK] Base consolidada salva em: {caminho_xlsx.relative_to(BASE_DIR)}")
    print(f"     -> Resumo Secretarias: {len(linhas_resumo)} linhas")
    print(f"     -> Fato Atrasos:       {len(linhas_atrasos)} registros de servidores")
    print(f"     -> Fato Faltas:        {len(linhas_faltas)} registros de servidores")

    # Gravar CSVs individuais (caso o usuário queira importar CSV)
    import csv
    def salvar_csv(linhas, nome_arquivo):
        if not linhas:
            return
        p_csv = POWERBI_DIR / nome_arquivo
        with open(p_csv, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=linhas[0].keys(), delimiter=";")
            writer.writeheader()
            writer.writerows(linhas)
        print(f"[OK] CSV gerado: {p_csv.relative_to(BASE_DIR)}")

    salvar_csv(linhas_resumo, "resumo_secretarias.csv")
    salvar_csv(linhas_atrasos, "fato_atrasos.csv")
    salvar_csv(linhas_faltas, "fato_faltas.csv")

    return {
        "resumo": len(linhas_resumo),
        "atrasos": len(linhas_atrasos),
        "faltas": len(linhas_faltas),
        "arquivo_excel": caminho_xlsx,
    }


if __name__ == "__main__":
    consolidar_monitoramento()
