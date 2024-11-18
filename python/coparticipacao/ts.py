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
        table_data.append([
            Paragraph(str(row['DATA_REALIZACAO']), content_style),
            Paragraph(row['NOME'], content_style),
            Paragraph(row['SERVICO'], content_style),
            Paragraph(str(int(row['QUANTIDADE'])), content_style),  # Adicione a coluna QUANTIDADE
            Paragraph(row['PRESTADOR'], content_style),
            f"R$ {row['VALOR_COM_TAXA_FM']:.2f}"
        ])
        total_valor += row['VALOR_COM_TAXA_FM']

    table_data.append(["", "", "", "", Paragraph("VALOR TOTAL", content_style), f"R$ {total_valor:.2f}"])

    light_green = colors.Color(red=0.8, green=1.0, blue=0.8)
    table = Table(table_data, colWidths=[80, 100, 200, 60, 150, 60])  # Ajuste as larguras das colunas
    
    # Ajustando o estilo da pagina do PDF

    # Cabeçalho
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),

        # Estilo para as linhas do conteudo, exceto a ultima linha
        ('BACKGROUND', (0, 1), (-1, -2), light_green),
        ('GRID', (0, 1), (-1, -2), 1, colors.black),
        ('ALIGN', (0, 1), (-1, -2), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -2), 9),

        # Estilo para a ultima linha (valor total)
        ('BACKGROUND', (0, -1), (-1, -1), light_green),
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (-1, -2), 'RIGHT'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, -1), (-1, -1), 10),

        # caso a ultima linha seja dividida entre paginas, aplicar o estilo tambem
        ('BACKGROUND', (0, 'splitlast'), (-1, 'splitlast'), light_green),
        ('TEXTCOLOR', (0, 'splitlast'), (-1, 'splitlast'), colors.black),
        ('ALIGN', (0, 1), (-1, -2), 'CENTER'),
        ('FONTNAME', (0, 'splitlast'), (-1, 'splitlast'), 'Helvetica'),
        ('FONTSIZE', (0, 'splitlast'), (-1, 'splitlast'), 9)
    ]))

    elements.append(table)

    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    print(f"PDF generated: {output_pdf}")

