from PyPDF2 import PdfReader
import re
import pandas as pd

def extract_data_from_pdf(pdf_path):
    data = []
    current_employee = {}
    
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text = page.extract_text()
            
            lines = text.split('\n')
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Ignorar linhas que parecem ser cabeçalhos ou rodapés
                if re.search(r'Nro Funcional|Nome|Ocorrência|Data|Quantidade|PREFEITURA|Página:|sistemas.pmp.sp.gov.br', line, re.IGNORECASE):
                    continue

                # Regex para identificar uma nova entrada de funcionário
                # Nro Funcional (XX.XXX-X ou XXX.XXX-X), Nome (qualquer coisa até a ocorrência)
                # O nome deve ser o mais específico possível para não invadir a ocorrência
                # A ocorrência geralmente começa com uma palavra capitalizada ou uma frase específica
                # Vamos tentar capturar o nome até o início da ocorrência ou até o final da linha se não houver ocorrência na mesma linha
                
                # Novo regex para capturar Nro Funcional e Nome de forma mais precisa
                # O nome pode conter letras maiúsculas, minúsculas, acentos e espaços.
                # A ocorrência geralmente começa com uma palavra como 'Frequência', 'Férias', 'Abono', 'Tratamento', 'Aguardando'
                employee_and_occurrence_match = re.match(r'(\d{2,3}\.\d{3}-\d)\s+([A-ZÀ-Ú][A-ZÀ-Úa-zà-ú\s\-]+?)\s+(Frequência normal|Férias regulamentares|Abono|Tratamento de saúde|Auxílio doença|Aguardando perícia sempem)(?:\s+(\d{2}/\d{2}/\d{4}))?(?:\s+([\d,]+))?$', line)
                
                if employee_and_occurrence_match:
                    if current_employee:
                        data.append(current_employee)
                    
                    nro_funcional = employee_and_occurrence_match.group(1)
                    nome = employee_and_occurrence_match.group(2).strip()
                    ocorrencia = employee_and_occurrence_match.group(3).strip()
                    data_ocorrencia = employee_and_occurrence_match.group(4) if employee_and_occurrence_match.group(4) else ''
                    quantidade = employee_and_occurrence_match.group(5).replace(',', '.') if employee_and_occurrence_match.group(5) else ''
                    
                    current_employee = {
                        'Nro Funcional': nro_funcional,
                        'Nome': nome,
                        'Ocorrência': [ocorrencia],
                        'Data': [data_ocorrencia],
                        'Quantidade': [quantidade]
                    }

                elif current_employee: # Linhas subsequentes para o mesmo funcionário (ocorrências adicionais)
                    # Regex para capturar ocorrências, data e quantidade em linhas subsequentes
                    occurrence_only_match = re.match(r'\s*(Frequência normal|Férias regulamentares|Abono|Tratamento de saúde|Auxílio doença|Aguardando perícia sempem)(?:\s+(\d{2}/\d{2}/\d{4}))?(?:\s+([\d,]+))?$', line)
                    if occurrence_only_match:
                        current_employee['Ocorrência'].append(occurrence_only_match.group(1).strip())
                        if occurrence_only_match.group(2):
                            current_employee['Data'].append(occurrence_only_match.group(2))
                        else:
                            current_employee['Data'].append('')
                        if occurrence_only_match.group(3):
                            current_employee['Quantidade'].append(occurrence_only_match.group(3).replace(',', '.'))
                        else:
                            current_employee['Quantidade'].append('')

        if current_employee:
            data.append(current_employee)

    # Normalizar os dados para criar um DataFrame
    final_data = []
    for entry in data:
        nro_funcional = entry['Nro Funcional']
        nome = entry['Nome']
        
        # Garantir que todas as listas tenham o mesmo comprimento para zip
        max_len = max(len(entry['Ocorrência']), len(entry['Data']), len(entry['Quantidade']))
        
        ocorrencias = entry['Ocorrência'] + [''] * (max_len - len(entry['Ocorrência']))
        datas = entry['Data'] + [''] * (max_len - len(entry['Data']))
        quantidades = entry['Quantidade'] + [''] * (max_len - len(entry['Quantidade']))

        for i in range(max_len):
            final_data.append({
                'Nro Funcional': nro_funcional,
                'Nome': nome,
                'Ocorrência': ocorrencias[i],
                'Data': datas[i],
                'Quantidade': quantidades[i]
            })

    df = pd.DataFrame(final_data)
    return df

if __name__ == '__main__':
    pdf_file = 'frequencia.pdf'  # Nome do seu arquivo PDF
    excel_file = 'frequencia.xlsx'
    
    df = extract_data_from_pdf(pdf_file)
    df.to_excel(excel_file, index=False)
    print(f'Dados extraídos e salvos em {excel_file}')
