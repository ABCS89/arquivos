from flask import Flask, render_template, request, send_file
import os
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.pdfgen import canvas
from io import BytesIO

app = Flask(__name__)

# Caminho base para as faturas
FATURAS_DIR = os.path.join('coparticipacao', 'faturas')

# Funções auxiliares (as mesmas do seu script original)
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
    footer_text = "Criado por André Bueno (DRH)"
    canvas.setFont('Helvetica', 9)
    canvas.drawString(30, 20, footer_text)
    canvas.restoreState()

def generate_pdf(df, nr_funcional, mes_escolhido=None):
    nr_funcional = str(int(float(nr_funcional)))
    df['NR_FUNCIONAL'] = df['NR_FUNCIONAL'].fillna('').astype(str).str.strip()
    df['NR_FUNCIONAL'] = df['NR_FUNCIONAL'].apply(lambda x: x.split('.')[0] if '.' in x else x)
    filtered_df = df[df['NR_FUNCIONAL'] == nr_funcional]

    if filtered_df.empty:
        print(f"Nenhum registro encontrado para NR_FUNCIONAL: {nr_funcional}")
        return None  # Retorna None se não houver dados

    buffer = BytesIO()  # Usar BytesIO para gerar o PDF na memória
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
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
    buffer.seek(0)  # Retorna ao início do buffer
    return buffer

# Rota principal (formulário)
@app.route('/', methods=['GET', 'POST'])
def index():
    meses = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    anos = ["2024", "2025"] # Adicione os anos que você tem as faturas

    if request.method == 'POST':
        mes_escolhido = request.form['mes'].strip().lower()
        ano_escolhido = request.form['ano'].strip()
        nr_funcional = request.form['nr_funcional'].strip()

        # Constrói o caminho para o diretório do ano
        ano_dir = os.path.join(FATURAS_DIR, ano_escolhido)

        # Verifica se o diretório do ano existe
        if not os.path.isdir(ano_dir):
            return f"Diretório para o ano {ano_escolhido} não encontrado."

        file_path = None
        # Procura pelo arquivo no diretório do ano
        for filename in os.listdir(ano_dir):
            if mes_escolhido in filename.lower() and filename.startswith('fatura_coparticipacao_'):
                file_path = os.path.join(ano_dir, filename)
                break

        if not file_path:
            return "Arquivo não encontrado para o mês e ano selecionados."

        try:
            df = read_file(file_path)
            pdf_buffer = generate_pdf(df, nr_funcional, mes_escolhido)

            if pdf_buffer:
                # Enviar o PDF como um arquivo para download
                return send_file(
                    pdf_buffer,
                    as_attachment=True,
                    download_name=f"fatura_{nr_funcional}_{mes_escolhido}.pdf",
                    mimetype='application/pdf'
                )
            else:
                return "Nenhum dado encontrado para os parâmetros informados."

        except Exception as e:
            return f"Erro ao processar o arquivo: {str(e)}"

    return render_template('index.html', meses=meses, anos=anos)

if __name__ == '__main__':
    app.run(debug=True)