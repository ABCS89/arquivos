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

def read_file(file_path):
    # Detectar o formato do arquivo com base na extensão
    file_extension = file_path.split('.')[-1].lower()

    if file_extension == 'ods':
        sheet_names = list_sheets(file_path)
        table_name = sheet_names[0]  # Supondo que a tabela desejada é a primeira
        return pd.read_excel(file_path, sheet_name=table_name, engine='odf')
    elif file_extension == 'csv':
        return pd.read_csv(file_path, sep=';', encoding='utf-8')
    elif file_extension in ['xls', 'xlsx']:
        return pd.read_excel(file_path, engine='openpyxl')
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
    canvas.setFont('Helvetica', 7)

    footer_lines = [
        "PREFEITURA DE PIRACICABA | SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO E GOVERNO",
        "Departamento de Recursos Humanos",
        "Rua Antônio Corrêa Barbosa, 2233 – 7º Andar – Centro – Piracicaba/SP",
        "Telefone: (19) 3403-1006"
    ]

    # Altura inicial do rodapé (mais próxima da base do documento)
    y = 40  
    for line in footer_lines:
        canvas.drawCentredString(doc.width / 2.0 + doc.leftMargin, y, line)
        y -= 9  # espaço entre linhas

    canvas.restoreState()

def generate_pdf(df, nr_funcional, output_pdf, mes_escolhido=None):
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
        # Depuração: Exibir o conteúdo original da coluna VALOR_COM_TAXA_FM
        print(f"Original 'VALOR_COM_TAXA_FM': {row['VALOR_COM_TAXA_FM']}")

        # Substituir a vírgula por ponto e tentar converter para número
        valor_com_taxa_fm = str(row['VALOR_COM_TAXA_FM']).replace(',', '.')

        # Garantir que o valor seja numérico após a substituição
        valor_com_taxa_fm = pd.to_numeric(valor_com_taxa_fm, errors='coerce')

        # Depuração: Exibir o valor após a conversão
        print(f"Após conversão 'VALOR_COM_TAXA_FM': {valor_com_taxa_fm}")

        if pd.isna(valor_com_taxa_fm):
            valor_com_taxa_fm = 0.0  # Caso o valor não possa ser convertido

        table_data.append([
            Paragraph(str(row['DATA_REALIZACAO']), content_style),
            Paragraph(row['NOME'], content_style),
            Paragraph(row['SERVICO'], content_style),
            Paragraph(str(int(row['QUANTIDADE'])), content_style),  # Adicione a coluna QUANTIDADE
            Paragraph(row['PRESTADOR'], content_style),
            f"R$ {valor_com_taxa_fm:.2f}"
        ])
        total_valor += valor_com_taxa_fm

    table_data.append(["", "", "", "", Paragraph("VALOR TOTAL", content_style), f"R$ {total_valor:.2f}"])

    light_green = colors.Color(red=0.8, green=1.0, blue=0.8)
    table = Table(table_data, colWidths=[80, 100, 200, 60, 150, 60])  # Ajuste as larguras das colunas
    
    # Estilos de tabela
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
        ('ALIGN', (0, 1), (-1, -2), 'RIGHT'),
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
        generate_pdf(df, nr_funcional, output_pdf, mes)

# Main script
meses = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

# Solicita o mês ao usuário
mes_escolhido = input("Digite o mês desejado ou 'total' para todos os meses: ").strip().lower()

# Verifica se a entrada é válida
if mes_escolhido not in meses and mes_escolhido != 'total':
    print("Mês inválido. Por favor, digite um mês válido ou 'total'.")
    exit()

# Procurando arquivos que contém o mês no nome ou todos os meses
files = [f for f in os.listdir('.') if os.path.isfile(f) and f.startswith('fatura_coparticipacao_')]

if mes_escolhido == 'total':
    nr_funcional = input("Enter NR_FUNCIONAL ou 'total' para todos os funcionais: ").strip().lower()
    if nr_funcional == 'total':
        output_dir = './faturas_totais'
        for mes in meses:
            for f in files:
                if mes in f.lower():
                    file_path = f
                    print(f"Arquivo encontrado: {file_path}")
                    df = read_file(file_path)
                    print(f"Primeiras linhas do arquivo lido:")
                    print(df.head())
                    generate_pdfs_for_all_functionals(df, output_dir, mes)
    else:
        output_dir = './faturas_totais'
        os.makedirs(output_dir, exist_ok=True)
        for mes in meses:
            for f in files:
                if mes in f.lower():
                    file_path = f
                    print(f"Arquivo encontrado: {file_path}")
                    df = read_file(file_path)
                    print(f"Primeiras linhas do arquivo lido:")
                    print(df.head())
                    output_pdf = os.path.join(output_dir, f"fatura_{nr_funcional}_{mes}.pdf")
                    generate_pdf(df, nr_funcional, output_pdf, mes)
else:
    file_path = None
    for f in files:
        if mes_escolhido in f.lower():
            file_path = f
            break

    if file_path:
        print(f"Arquivo encontrado: {file_path}")
        df = read_file(file_path)
        print("Primeiras linhas do arquivo lido:")
        print(df.head())

        nr_funcional = input("Enter NR_FUNCIONAL ou 'total' para todos os funcionais: ").strip().lower()
        if nr_funcional == 'total':
            output_dir = './faturas_totais'
            generate_pdfs_for_all_functionals(df, output_dir, mes_escolhido)
        else:
            output_pdf = f"fatura_{nr_funcional}_{mes_escolhido}.pdf"
            generate_pdf(df, nr_funcional, output_pdf, mes_escolhido)
    else:
        print(f"Nenhum arquivo encontrado com o padrão 'fatura_coparticipacao_{mes_escolhido}'.")
# Fim do script