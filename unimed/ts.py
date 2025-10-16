import pandas as pd
from docx import Document
import re
import os
import PyPDF2
from num2words import num2words
from datetime import datetime
import calendar


def extract_text_from_pdf(pdf_path):
    text = ''
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                text += reader.pages[page_num].extract_text()
    except Exception as e:
        print(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
    return text

# - Converter numeros para palavras, (escrever por extenso)
def number_to_currency_text_extended(number):
    try:
        # Separa a parte inteira e a parte decimal do número
        inteiro = int(number)
        decimal = int(round((number - inteiro) * 100)) # Pega os dois primeiros dígitos decimais

        # Converte a parte inteira para extenso
        texto_inteiro = num2words(inteiro, lang='pt_BR')

        # Converte a parte decimal para extenso, se houver
        if decimal > 0:
            texto_decimal = num2words(decimal, lang='pt_BR')
            return f'{texto_inteiro} reais e {texto_decimal} centavos'
        else:
            return f'{texto_inteiro} reais'
    except Exception as e:
           print(f"Erro ao converter número para extenso: {number} - {e}")
           return f'VALOR_POR_EXTENSO_ERRO_{number:.2f}'.replace('.', ',')

def extract_info_from_pdf_content(pdf_content):
    email_date_match = re.search(r'Data (\d{4}-\d{2}-\d{2})', pdf_content)
    email_address_match = re.search(r'Para <([^>]+)>', pdf_content)
    
    email_date = email_date_match.group(1) if email_date_match else 'DATA_NAO_ENCONTRADA'
    email_address = email_address_match.group(1) if email_address_match else 'EMAIL_NAO_ENCONTRADO'
    
    email_month = pd.to_datetime(email_date).strftime('%B').capitalize() if email_date != 'DATA_NAO_ENCONTRADA' else 'MES_NAO_ENCONTRADO'
    email_day = pd.to_datetime(email_date).day if email_date != 'DATA_NAO_ENCONTRADA' else 'DIA_NAO_ENCONTRADO'

    return email_date, email_address, email_month, email_day

def number_to_currency_text(number):
    return f'{number:.2f}'.replace('.', ',')

def normalize_name(name):
    # Remove acentos, caracteres especiais, espaços e converte para minúsculas
    name = name.lower()
    name = re.sub(r'[áàãâä]', 'a', name)
    name = re.sub(r'[éèêë]', 'e', name)
    name = re.sub(r'[íìîï]', 'i', name)
    name = re.sub(r'[óòõôö]', 'o', name)
    name = re.sub(r'[úùûü]', 'u', name)
    name = re.sub(r'[ç]', 'c', name)
    name = re.sub(r'[^a-z0-9]', '', name) # Remove qualquer coisa que não seja letra ou número
    return name

def generate_document(data_row, email_info, template_path='template.docx'):
    document = Document(template_path)

    nro_funcional = data_row['Nro Funcional']
    funcionario = data_row['Funcionário']
    total = data_row['Total']

    email_date, email_address, email_month, email_day = email_info

    replacements = {
        '[dia]': email_day,
        '[dia atual]': str(pd.Timestamp.now().day),
        '[ultimo dia do mês]': str(pd.Timestamp.now().days_in_month),
        
        '[mês]': email_month,
        '[mês atual]': str(pd.Timestamp.now().month,),
        
        '[ano atual]': str(pd.Timestamp.now().year),
                
        '[valor numérico]': f'{total:.2f}'.replace('.', ','),
        '[valor por extenso]': number_to_currency_text_extended(total),
        '[endereço de e-mail]': email_address,
        '[nome do servidor]': funcionario,
        '[endereço do servidor]': 'ENDERECO_NAO_ENCONTRADO',
        '[CEP do servidor]': 'CEP_NAO_ENCONTRADO'
    }

    for paragraph in document.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, str(value))

    output_filename = f'documento_{nro_funcional}_{funcionario.replace(" ", "_")}.docx'
    document.save(output_filename)
    print(f'Documento gerado: {output_filename}')

if __name__ == '__main__':
    ods_path = 'teste.ods'
    pdf_directory = '.' # Onde os PDFs de e-mail estão localizados
    df = pd.read_excel(ods_path, engine='odf')

    # Mapear PDFs para funcionários
    pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith('email.pdf')]
    
    pdf_map = {}
    for index, row in df.iterrows():
        funcionario_nome_planilha = row['Funcionário']
        normalized_funcionario_name = normalize_name(funcionario_nome_planilha)
        
        found_pdf = None
        for pdf_file in pdf_files:
            normalized_pdf_filename = normalize_name(pdf_file.replace('_email.pdf', ''))
            
            # Tenta encontrar o nome normalizado da planilha no nome normalizado do PDF
            if normalized_funcionario_name in normalized_pdf_filename:
                found_pdf = pdf_file
                break
        
        if found_pdf:
            pdf_map[row['Nro Funcional']] = found_pdf
        else:
            print(f"Aviso: Nenhum PDF de e-mail encontrado para o funcionário: {funcionario_nome_planilha} (Nro Funcional: {row['Nro Funcional']})")

    # Gerar um documento para cada linha da planilha
    for index, row in df.iterrows():
        nro_funcional = row['Nro Funcional']
        if nro_funcional in pdf_map:
            current_pdf_path = os.path.join(pdf_directory, pdf_map[nro_funcional])
            pdf_content = extract_text_from_pdf(current_pdf_path)
            email_date, email_address, email_month, email_day = extract_info_from_pdf_content(pdf_content)
            email_info = (email_date, email_address, email_month, email_day)
            generate_document(row, email_info, template_path='template.docx')
        else:
            # Este aviso já foi dado no loop anterior, mas é bom ter certeza
            pass