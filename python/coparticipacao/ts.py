import os
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.pdfgen import canvas

def list_sheets(file_path):
    xls = pd.ExcelFile(file_path, engine='odf')
    return xls.sheet_names

def read_file(file_path, table_name=None):
    if file_path.endswith('.ods'):
        return pd.read_excel(file_path, sheet_name=table_name, engine='odf')
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path, sep=';')  # Ajustar o separador do CSV
    elif file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)  # Suporte para arquivos .xlsx
    else:
        raise ValueError("Formato de arquivo não suportado")

def get_month_name(month_number):
    month_map = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return month_map.get(month_number, "Mês inválido")

def add_footer(canvas, doc):
    canvas.saveState()
    footer_text = "Criado por André Bueno (DRH)"
    canvas.setFont('Helvetica', 9)
    canvas.drawString(30, 20, footer_text)
    canvas.restoreState()

def generate_pdf(df, nr_funcional, output_pdf):
    nr_funcional = str(int(float(nr_funcional)))
    df['NR_FUNCIONAL'] = df['NR_FUNCIONAL'].fillna('').astype(str).str.strip()
    df['NR_FUNCIONAL'] = df['NR_FUNCIONAL'].apply(lambda x: x.split('.')[0] if '.' in x else x)
    filtered_df = df[df['NR_FUNCIONAL'] == nr_funcional]

    if filtered_df.empty:
        print(f"Nenhum registro encontrado para NR_FUNCIONAL: {nr_funcional}")
        return

    doc = SimpleDocTemplate(output_pdf, pagesize=landscape(A4))
    elements = []

    styles = getSampleStyleSheet()
    title_style = styles['Title']
    title_style.fontSize = 14

    month_reference = get_month_name(int(filtered_df.iloc[0]['MM_REFERENCIA']))
    elements.extend([
        Paragraph(f"Funcional: {nr_funcional}", title_style),
        Paragraph(f"Titular: {filtered_df.iloc[0]['TITULAR']}", title_style),
        Paragraph(f"Mês Referência: {month_reference}", title_style)
    ])

    # Atualize os nomes das colunas da tabela para incluir "QUANTIDADE"
    table_data = [["Realização", "Beneficiário", "Serviço", "Quantidade", "Prestador", "Valor"]]
    total_valor = 0

    content_style = ParagraphStyle(name="Content", fontSize=7, leading=8)

    for index, row in filtered_df.iterrows():
        try:
            valor = float(row['VALOR_COM_TAXA_FM'])
        except ValueError:
            valor = 0.0  # Em caso de erro de conversão, define como 0.0

        table_data.append([
            Paragraph(str(row['DATA_REALIZACAO']), content_style),
            Paragraph(row['NOME'], content_style),
            Paragraph(row['SERVICO'], content_style),
            Paragraph(str(int(row['QUANTIDADE'])), content_style),  # Adicionando a coluna QUANTIDADE
            Paragraph(row['PRESTADOR'], content_style),
            f"R$ {valor:.2f}"
        ])
        total_valor += valor

    table_data.append(["", "", "", "", Paragraph("VALOR TOTAL", content_style), f"R$ {total_valor:.2f}"])

    light_green = colors.Color(red=0.8, green=1.0, blue=0.8)
    table = Table(table_data, colWidths=[80, 100, 200, 60, 150, 60])  # Ajuste as larguras das colunas
    
    # Ajustando o estilo da página do PDF
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),

        ('BACKGROUND', (0, 1), (-1, -2), light_green),
        ('GRID', (0, 1), (-1, -2), 1, colors.black),
        ('ALIGN', (0, 1), (-1, -2), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -2), 9),

        ('BACKGROUND', (0, -1), (-1, -1), light_green),
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, -1), (-1, -1), 10),

        ('BACKGROUND', (0, 'splitlast'), (-1, 'splitlast'), light_green),
        ('TEXTCOLOR', (0, 'splitlast'), (-1, 'splitlast'), colors.black),
        ('ALIGN', (0, 1), (-1, -2), 'CENTER'),
        ('FONTNAME', (0, 'splitlast'), (-1, 'splitlast'), 'Helvetica'),
        ('FONTSIZE', (0, 'splitlast'), (-1, 'splitlast'), 9)
    ]))

    elements.append(table)
    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    print(f"PDF gerado: {output_pdf}")

def generate_pdfs_for_all_functionals(df, output_dir, mes):
    os.makedirs(output_dir, exist_ok=True)
    unique_funcionals = df['NR_FUNCIONAL'].unique()

    for nr_funcional in unique_funcionals:
        if pd.isna(nr_funcional) or str(nr_funcional).strip() == '':
            continue
        output_pdf = os.path.join(output_dir, f"fatura_{str(int(float(nr_funcional)))}_{mes}.pdf")
        generate_pdf(df, nr_funcional, output_pdf)

# Main script
meses = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

# Solicita o mês ao usuário
mes_escolhido = input("Digite o mês desejado ou 'total' para todos os meses: ").strip().lower()

# Verifica se a entrada é válida
if mes_escolhido not in meses and mes_escolhido != 'total':
    print("Mês inválido. Por favor, digite um mês válido ou 'total'.")
    exit()

# Procurando arquivos que contêm o mês no nome ou todos os meses
files = [f for f in os.listdir('.') if os.path.isfile(f) and f.startswith(f"fatura_coparticipacao_{mes_escolhido}")]

if not files:
    print(f"Nenhum arquivo com o padrão 'fatura_coparticipacao_{mes_escolhido}'.")
    exit()

# Identificar o primeiro arquivo válido com extensão suportada
ods_path = None
for f in files:
    if f.endswith(('.ods', '.csv', '.xlsx')):  # Verifica se o arquivo tem formato válido
        ods_path = f
        break

if not ods_path:
    print(f"Nenhum arquivo suportado (.ods, .csv, .xlsx) encontrado com o padrão 'fatura_coparticipacao_{mes_escolhido}'.")
    exit()

print(f"Arquivo encontrado: {ods_path}")

if mes_escolhido == 'total':
    nr_funcional = input("Enter NR_FUNCIONAL ou 'total' para todos os funcionais: ").strip().lower()
    if nr_funcional == 'total':
        output_dir = './faturas_totais'
        for mes in meses:
            for f in files:
                if mes in f.lower() and f.endswith(('.ods', '.csv', '.xlsx')):
                    ods_path = f
                    print(f"Arquivo encontrado: {ods_path}")
                    df = read_file(ods_path)  # Ajuste aqui para carregar qualquer extensão suportada
                    print(f"Primeiras linhas do arquivo:\n{df.head()}")
                    generate_pdfs_for_all_functionals(df, output_dir, mes)
else:
    df = read_file(ods_path)  # Carrega o arquivo selecionado
    print(f"Primeiras linhas do arquivo:\n{df.head()}")
    nr_funcional = input("Enter NR_FUNCIONAL: ").strip()
    output_pdf = f"./faturas_totais/fatura_{nr_funcional}.pdf"
    generate_pdf(df, nr_funcional, output_pdf)
