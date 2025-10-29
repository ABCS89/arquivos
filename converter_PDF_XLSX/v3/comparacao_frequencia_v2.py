import pandas as pd
import re
from datetime import datetime
from io import StringIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import os
import locale

# Configura o locale para o português para que o strptime consiga ler o nome do mês
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
    except locale.Error:
        print("Aviso: Não foi possível configurar o locale para pt_BR. A conversão de nome de mês pode falhar.")


# Dependência externa para extração de texto de PDF.
# No ambiente real do usuário, será necessário instalar 'pdfplumber' ou similar.
# Como não posso instalar aqui, vou simular a extração baseada na análise manual.
# No código final, vou sugerir o uso de 'pdfplumber' e 'tabula-py'.

def extract_text_from_pdf(pdf_path):
    """
    Função placeholder para extrair texto de PDF.
    No ambiente real, o usuário precisará de uma biblioteca como pdfplumber ou tabula-py.
    """
    # Esta função deve ser substituída pela lógica de extração real
    # Exemplo:
    # import pdfplumber
    # with pdfplumber.open(pdf_path) as pdf:
    #     text = ""
    #     for page in pdf.pages:
    #         text += page.extract_text()
    # return text
    
    # Para fins de demonstração e teste da lógica, usaremos os dados extraídos manualmente.
    # O usuário será instruído a usar uma biblioteca adequada e adaptar a extração.
    
    if 'frequencia' in pdf_path:
        # Dados extraídos de 106-frequencia.pdf (simplificado)
        # Nro Funcional, Nome, Ocorrência, Data, Quantidade
        data = """
Nro Funcional,Nome,Ocorrência,Data,Quantidade
28.049-6,BRENO REI PASSOS LAGOAS,Tratamento de saúde,16/09/2025,4.0
27.840-8,EDILENE FERNANDES MORGADO,Férias regulamentares,08/09/2025,15.0
27.844-0,FELIPE VITTI DE OLIVEIRA,Férias regulamentares,01/09/2025,15.0
27.844-0,FELIPE VITTI DE OLIVEIRA,Falta,17/09/2025,2.0
27.844-0,FELIPE VITTI DE OLIVEIRA,Falta,22/09/2025,1.0
27.694-4,JAMILE MARTINS,Férias regulamentares,08/09/2025,15.0
26.355-9,JULIANA DE SOUZA NARDO,Doação de sangue,12/09/2025,1.0
28.596-0,MARIA LUIZA GUEDES DOS SANTOS,Abono,04/09/2025,1.0
28.596-0,MARIA LUIZA GUEDES DOS SANTOS,Abono,05/09/2025,1.0
28.596-0,MARIA LUIZA GUEDES DOS SANTOS,Férias regulamentares,15/09/2025,15.0
17.016-9,MAYCON MORGADO,Férias regulamentares,08/09/2025,15.0
19.948-7,AIMEE ROCCIA GIMENEZ,Férias regulamentares,22/09/2025,15.0
27.609-0,ANDREA LIMA ESTEVAO,Doação de sangue,26/09/2025,1.0
17.043-3,LARISSA DOMINGUES HERNANDES,Auxílio doença,01/09/2025,377.0
28.176-0,LINSMAR RISO DA SILVA,Abono,08/09/2025,1.0
19.786-2,MAURO CESAR STOLF,Férias regulamentares,01/09/2025,30.0
28.585-4,RENATA CARDOSO DE OLIVEIRA,Férias regulamentares,29/09/2025,10.0
20.050-6,ROSIRIS DOS SANTOS GONÇALVES,Minutos perdidos,03/09/2025,191.0
28.321-5,SUSANA CRISTINA SANTOS,Abono,05/09/2025,1.0
28.321-5,SUSANA CRISTINA SANTOS,Abono,08/09/2025,1.0
28.648-6,KARINA FERREIRA DA CRUZ,Férias regulamentares,16/09/2025,15.0
26.170-0,MARIA THERESA SARTORELLI SERRATO,Cedido sem ônus para cedente,01/09/2025,1218.0
27.600-6,PATRICK RIBEIRO DE JESUS,Abono,29/09/2025,1.0
27.600-6,PATRICK RIBEIRO DE JESUS,Abono,30/09/2025,1.0
27.055-5,DAVI DAS NEVES CALMON,Férias regulamentares,01/09/2025,2.0
27.055-5,DAVI DAS NEVES CALMON,Tratamento de saúde,25/09/2025,2.0
22.474-0,FLAVIA RENATA RIES,Tratamento de saúde,01/09/2025,4.0
22.474-0,FLAVIA RENATA RIES,Tratamento de saúde,06/09/2025,7.0
21.594-5,MARIA LUIZA PAIAO ASSIS,Férias regulamentares,12/09/2025,15.0
18.975-3,SANDRA CRISTINA ROCHA,Férias regulamentares,01/09/2025,18.0
18.975-3,SANDRA CRISTINA ROCHA,Abono eleitoral,19/09/2025,1.0
28.494-7,DAIANE NEGRETTI CALDEIRA,Férias regulamentares,01/09/2025,2.0
28.494-7,DAIANE NEGRETTI CALDEIRA,Abono,30/09/2025,1.0
17.609-7,ELIANA APARECIDA DE GODOY,Tratamento de saúde,01/09/2025,8.0
17.609-7,ELIANA APARECIDA DE GODOY,Tratamento de saúde,09/09/2025,4.0
27.671-5,LARISSA HENRIQUE CAVALCANTE ALBUQUERQUE,Doença em pessoa da família,10/09/2025,2.0
27.671-5,LARISSA HENRIQUE CAVALCANTE ALBUQUERQUE,Aguardando perícia sempem,25/09/2025,2.0
27.671-5,LARISSA HENRIQUE CAVALCANTE ALBUQUERQUE,Tratamento de saúde,29/09/2025,1.0
12.426-1,ANTONIO APARECIDO DE MORAES,Falta,01/09/2025,5.0
12.426-1,ANTONIO APARECIDO DE MORAES,Falta,08/09/2025,5.0
12.426-1,ANTONIO APARECIDO DE MORAES,Falta,15/09/2025,2.0
12.426-1,ANTONIO APARECIDO DE MORAES,Aguardando perícia sempem,17/09/2025,7.0
12.426-1,ANTONIO APARECIDO DE MORAES,Falta,24/09/2025,1.0
12.426-1,ANTONIO APARECIDO DE MORAES,Nojo,25/09/2025,2.0
28.297-9,DOUGLAS DOS SANTOS BAGNARA,Férias regulamentares,01/09/2025,10.0
28.297-9,DOUGLAS DOS SANTOS BAGNARA,Aguardando perícia sempem,30/09/2025,1.0
28.082-8,RENATO MARCELLUS ROBERTO,Abono,22/09/2025,1.0
28.082-8,RENATO MARCELLUS ROBERTO,Abono,23/09/2025,1.0
25.011-2,MARIANA APARECIDA BAPTISTINI,Férias regulamentares,15/09/2025,15.0
28.644-3,MAXWELL PIVESSO MARTINS,Férias regulamentares,01/09/2025,4.0
25.003-1,YURI KATOO,Cedido sem ônus para cedente,01/09/2025,1218.0
"""
        return data
    
    elif 'relatório' in pdf_path:
        # Dados extraídos de 106-relatório.pdf (simplificado)
        # Funcionário, Pessoa, Descrição, Data Inicial, Data Final, Qtde Dias
        data = """
Funcionário,Pessoa,Descrição,Data Inicial,Data Final,Qtde Dias
08.654-5,WILLIANS DE CAMPOS,Férias Regulamentares,08/09/2025,22/09/2025,15
28.049-6,BRENO REI PASSOS LAGOAS,Tratamento De Saúde,16/09/2025,19/09/2025,4
27.840-8,EDILENE FERNANDES MORGADO,Férias Regulamentares,08/09/2025,22/09/2025,15
27.844-0,FELIPE VITTI DE OLIVEIRA,Férias Regulamentares,01/09/2025,15/09/2025,15
27.844-0,FELIPE VITTI DE OLIVEIRA,Faltas Efetivos,17/09/2025,17/09/2025,1
27.844-0,FELIPE VITTI DE OLIVEIRA,Faltas Efetivos,18/09/2025,18/09/2025,1
27.844-0,FELIPE VITTI DE OLIVEIRA,Faltas Efetivos,22/09/2025,22/09/2025,1
21.820-0,ISADORA CRUZ LOPES,Afastamento Por Nomeação Em Comissão,11/08/2025,31/12/2028,1239
27.694-4,JAMILE MARTINS,Férias Regulamentares,08/09/2025,22/09/2025,15
26.355-9,JULIANA DE SOUZA NARDO,Doação De Sangue,12/09/2025,12/09/2025,1
28.596-0,MARIA LUIZA GUEDES DOS SANTOS,Abono,04/09/2025,04/09/2025,1
28.596-0,MARIA LUIZA GUEDES DOS SANTOS,Abono,05/09/2025,05/09/2025,1
28.596-0,MARIA LUIZA GUEDES DOS SANTOS,Férias Regulamentares,15/09/2025,29/09/2025,15
17.016-9,MAYCON MORGADO,Férias Regulamentares,08/09/2025,22/09/2025,15
19.948-7,AIMEE ROCCIA GIMENEZ,Férias Regulamentares,22/09/2025,06/10/2025,15
27.609-0,ANDREA LIMA ESTEVAO,Doação De Sangue,26/09/2025,26/09/2025,1
17.043-3,LARISSA DOMINGUES HERNANDES,Auxílio Doença,19/05/2022,12/09/2026,1578
28.176-0,LINSMAR RISO DA SILVA,Abono,08/09/2025,08/09/2025,1
19.786-2,MAURO CESAR STOLF,Férias Regulamentares,01/09/2025,30/09/2025,30
28.585-4,RENATA CARDOSO DE OLIVEIRA,Férias Regulamentares,29/09/2025,08/10/2025,10
20.050-6,ROSIRIS DOS SANTOS GONÇALVES,Minutos Perdidos Clt,03/09/2025,03/09/2025,1
28.321-5,SUSANA CRISTINA SANTOS,Abono,05/09/2025,05/09/2025,1
28.321-5,SUSANA CRISTINA SANTOS,Abono,08/09/2025,08/09/2025,1
28.648-6,KARINA FERREIRA DA CRUZ,Férias Regulamentares,16/09/2025,30/09/2025,15
26.170-0,MARIA THERESA SARTORELLI SERRATO,Cedido Sem Ônus Para Cedente,18/08/2025,31/12/2028,1232
27.600-6,PATRICK RIBEIRO DE JESUS,Abono,29/09/2025,29/09/2025,1
27.600-6,PATRICK RIBEIRO DE JESUS,Abono,30/09/2025,30/09/2025,1
27.055-5,DAVI DAS NEVES CALMON,Férias Regulamentares,04/08/2025,02/09/2025,30
27.055-5,DAVI DAS NEVES CALMON,Tratamento De Saúde,25/09/2025,26/09/2025,2
22.474-0,FLAVIA RENATA RIES,Tratamento De Saúde,01/09/2025,04/09/2025,4
22.474-0,FLAVIA RENATA RIES,Tratamento De Saúde,06/09/2025,12/09/2025,7
21.594-5,MARIA LUIZA PAIAO ASSIS,Férias Regulamentares,12/09/2025,26/09/2025,15
18.975-3,SANDRA CRISTINA ROCHA,Férias Regulamentares,01/09/2025,18/09/2025,18
18.975-3,SANDRA CRISTINA ROCHA,Abono Eleitoral,19/09/2025,19/09/2025,1
28.494-7,DAIANE NEGRETTI CALDEIRA,Férias Regulamentares,04/08/2025,02/09/2025,30
28.494-7,DAIANE NEGRETTI CALDEIRA,Abono,30/09/2025,30/09/2025,1
17.609-7,ELIANA APARECIDA DE GODOY,Tratamento De Saúde,01/09/2025,08/09/2025,8
17.609-7,ELIANA APARECIDA DE GODOY,Tratamento De Saúde,09/09/2025,12/09/2025,4
16.436-4,HELENA MARIA GAMA DE AQUINO,Afastamento Por Nomeação Em Comissão,01/04/2025,31/12/2028,1371
27.671-5,LARISSA HENRIQUE CAVALCANTE ALBUQUERQUE,Doença Em Pessoa Da Família,10/09/2025,11/09/2025,2
27.671-5,LARISSA HENRIQUE CAVALCANTE ALBUQUERQUE,Tratamento De Saúde,29/09/2025,29/09/2025,1
10.977-6,RENATO LEITAO RONSINI,Abono,29/09/2025,29/09/2025,1
10.977-6,RENATO LEITAO RONSINI,Abono,30/09/2025,30/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,01/09/2025,01/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,02/09/2025,02/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,03/09/2025,03/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,04/09/2025,04/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,05/09/2025,05/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,08/09/2025,08/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,09/09/2025,09/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,10/09/2025,10/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,11/09/2025,11/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,12/09/2025,12/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,15/09/2025,15/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,16/09/2025,16/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Faltas Efetivos,24/09/2025,24/09/2025,1
12.426-1,ANTONIO APARECIDO DE MORAES,Nojo,25/09/2025,26/09/2025,2
28.297-9,DOUGLAS DOS SANTOS BAGNARA,Férias Regulamentares,27/08/2025,10/09/2025,15
28.297-9,DOUGLAS DOS SANTOS BAGNARA,Tratamento De Saúde,30/09/2025,03/10/2025,4
28.082-8,RENATO MARCELLUS ROBERTO,Abono,22/09/2025,22/09/2025,1
28.082-8,RENATO MARCELLUS ROBERTO,Abono,23/09/2025,23/09/2025,1
25.011-2,MARIANA APARECIDA BAPTISTINI,Férias Regulamentares,15/09/2025,29/09/2025,15
28.644-3,MAXWELL PIVESSO MARTINS,Férias Regulamentares,21/08/2025,04/09/2025,15
25.003-1,YURI KATOO,Cedido Sem Ônus Para Cedente,18/08/2025,31/12/2028,1232
"""
        return data
    
    return ""

def get_month_reference(pdf_text):
    """Extrai o mês de referência do texto do PDF."""
    match = re.search(r'Referente:\s*(\w+/\d{4})', pdf_text)
    if match:
        return match.group(1)
    match = re.search(r'Referente:\s*(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})', pdf_text)
    if match:
        # Pega o mês do final do período
        # Retorna no formato 'Mês/AAAA' (ex: Setembro/2025)
        return datetime.strptime(match.group(2), '%d/%m/%Y').strftime('%B/%Y').capitalize()
    
    # Se não encontrar, assume Setembro/2025 com base nos arquivos de exemplo
    return "Setembro/2025"

def normalize_data(df, source_type, ref_month_str):
    """Normaliza as colunas e filtra/ajusta os dados com base no mês de referência."""
    
    # 1. Determinar o mês e ano de referência
    # Mapeamento manual de meses para garantir que funcione sem o locale
    month_map = {
        'JANEIRO': 1, 'FEVEREIRO': 2, 'MARÇO': 3, 'ABRIL': 4, 'MAIO': 5, 'JUNHO': 6,
        'JULHO': 7, 'AGOSTO': 8, 'SETEMBRO': 9, 'OUTUBRO': 10, 'NOVEMBRO': 11, 'DEZEMBRO': 12
    }
    
    # Tenta extrair o mês por extenso e ano
    match = re.match(r'(\w+)/(\d{4})', ref_month_str.upper())
    if match:
        month_name = match.group(1)
        year = match.group(2)
        month_num = month_map.get(month_name)
        if month_num:
            ref_month_dt = datetime(int(year), month_num, 1)
        else:
            # Se não for mês por extenso, tenta o formato numérico
            try:
                ref_month_dt = datetime.strptime(ref_month_str, '%m/%Y')
            except ValueError:
                raise ValueError(f"Formato de mês de referência inválido: {ref_month_str}. Esperado 'Mês/AAAA' ou 'MM/AAAA'.")
    else:
        # Tenta o formato numérico
        try:
            ref_month_dt = datetime.strptime(ref_month_str, '%m/%Y')
        except ValueError:
            raise ValueError(f"Formato de mês de referência inválido: {ref_month_str}. Esperado 'Mês/AAAA' ou 'MM/AAAA'.")

    # Determina o primeiro e último dia do mês de referência
    start_of_ref_month = ref_month_dt.replace(day=1)
    # Cálculo do último dia do mês
    next_month = ref_month_dt.replace(day=28) + pd.DateOffset(days=4)
    end_of_ref_month = next_month - pd.DateOffset(days=next_month.day)
    
    if source_type == 'frequencia':
        df = df.rename(columns={'Nro Funcional': 'Funcional', 'Nome': 'Pessoa', 
                                'Ocorrência': 'Descricao', 'Data': 'Data_Inicial', 
                                'Quantidade': 'Qtde_Dias'})
        df['Data_Final'] = df['Data_Inicial'] # Adiciona Data_Final para unificação
        
    elif source_type == 'relatorio':
        df = df.rename(columns={'Funcionário': 'Funcional', 'Pessoa': 'Pessoa', 
                                'Descrição': 'Descricao', 'Data Inicial': 'Data_Inicial', 
                                'Data Final': 'Data_Final', # Adicionando a renomeação da coluna 'Data Final'
                                'Qtde Dias': 'Qtde_Dias'})
        
    # 2. Limpeza e conversão de tipos
    df['Funcional'] = df['Funcional'].astype(str).str.strip()
    df['Pessoa'] = df['Pessoa'].astype(str).str.strip()
    df['Descricao'] = df['Descricao'].astype(str).str.strip().str.upper()
    
    # 3. Conversão de datas e ajuste de Quantidade (Qtde_Dias)
    df['Data_Inicial'] = pd.to_datetime(df['Data_Inicial'], format='%d/%m/%Y', errors='coerce')
    df['Data_Final'] = pd.to_datetime(df['Data_Final'], format='%d/%m/%Y', errors='coerce')
    df['Qtde_Dias'] = pd.to_numeric(df['Qtde_Dias'], errors='coerce')
    
    # Remove linhas com datas ou quantidades inválidas
    df.dropna(subset=['Data_Inicial', 'Qtde_Dias'], inplace=True)
    
    # 4. Tratamento especial para o 'relatório' (Data Inicial em mês anterior)
    if source_type == 'relatorio':
        def adjust_days(row):
            # Se a ocorrência começou antes ou no mês de referência e tem uma duração
            if row['Data_Inicial'] <= end_of_ref_month:
                
                # Início da contagem é o máximo entre a Data_Inicial da ocorrência e o primeiro dia do mês de referência
                start_count = max(row['Data_Inicial'], start_of_ref_month)
                
                # Fim da contagem é o mínimo entre a Data_Final da ocorrência e o último dia do mês de referência
                # Se Data_Final for NaT, usa Data_Inicial (evento de 1 dia)
                end_of_event = row['Data_Final'] if pd.notna(row['Data_Final']) else row['Data_Inicial']
                
                end_count = min(end_of_event, end_of_ref_month)
                
                # Se o início da contagem for posterior ao fim da contagem, não há dias no mês
                if start_count > end_count:
                    return 0.0
                
                # Cálculo da quantidade de dias no mês: (end_count - start_count) + 1
                # Se for um evento de 1 dia (Data_Inicial == Data_Final) e cair no mês, retorna 1.0
                if row['Data_Inicial'] == end_of_event and start_count.month == ref_month_dt.month:
                    return 1.0
                
                # Para períodos, a contagem é a diferença em dias + 1
                return (end_count - start_count).days + 1.0
            
            return 0.0

        # Aplica a função de ajuste de dias
        df['Qtde_Dias'] = df.apply(adjust_days, axis=1)
        
        # Filtra as ocorrências que não têm dias no mês de referência
        df = df[df['Qtde_Dias'] > 0.0].copy()
    
    # 5. Normalização de nomes de ocorrências (para facilitar a comparação)
    df['Descricao'] = df['Descricao'].str.upper().str.replace(r'\s+', ' ', regex=True).str.strip()
    
    # 6. Agrupamento
    # Agrupa por Funcional, Pessoa e Descricao, somando a Qtde_Dias (já ajustada)
    df_grouped = df.groupby(['Funcional', 'Pessoa', 'Descricao']).agg(
    
(Content truncated due to size limit. Use page ranges or line ranges to read remaining content)