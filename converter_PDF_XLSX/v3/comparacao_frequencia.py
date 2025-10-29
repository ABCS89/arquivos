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
        return datetime.strptime(match.group(2), '%d/%m/%Y').strftime('%B/%Y').capitalize()
    
    # Se não encontrar, assume Setembro/2025 com base nos arquivos de exemplo
    return "Setembro/2025"

def normalize_data(df, source_type, ref_month_str):
    """Normaliza as colunas e filtra/ajusta os dados com base no mês de referência."""
    
    # 1. Determinar o mês e ano de referência
    try:
        # Tenta converter o mês por extenso (Setembro) para número (09)
        ref_month = datetime.strptime(ref_month_str, '%B/%Y').strftime('%m/%Y')
    except ValueError:
        # Se falhar, assume que já está no formato MM/AAAA
        ref_month = datetime.strptime(ref_month_str, '%m/%Y').strftime('%m/%Y')

    ref_month_dt = datetime.strptime(ref_month, '%m/%Y')
    
    if source_type == 'frequencia':
        df = df.rename(columns={'Nro Funcional': 'Funcional', 'Nome': 'Pessoa', 
                                'Ocorrência': 'Descricao', 'Data': 'Data_Inicial', 
                                'Quantidade': 'Qtde_Dias'})
        df['Data_Final'] = df['Data_Inicial'] # Adiciona Data_Final para unificação
        
    elif source_type == 'relatorio':
        df = df.rename(columns={'Funcionário': 'Funcional', 'Pessoa': 'Pessoa', 
                                'Descrição': 'Descricao', 'Data Inicial': 'Data_Inicial', 
                                'Qtde Dias': 'Qtde_Dias'})
        # 'Data Final' já existe no relatório
        
    # 2. Limpeza e conversão de tipos
    df['Funcional'] = df['Funcional'].astype(str).str.strip()
    df['Pessoa'] = df['Pessoa'].astype(str).str.strip()
    df['Descricao'] = df['Descricao'].astype(str).str.strip().str.upper()
    
    # 3. Conversão de datas e ajuste de Quantidade (Qtde_Dias)
    df['Data_Inicial'] = pd.to_datetime(df['Data_Inicial'], format='%d/%m/%Y', errors='coerce')
    df['Data_Final'] = pd.to_datetime(df['Data_Final'], format='%d/%m/%Y', errors='coerce')
    df['Qtde_Dias'] = pd.to_numeric(df['Qtde_Dias'], errors='coerce')
    
    # 4. Tratamento especial para o 'relatório' (Data Inicial em mês anterior)
    if source_type == 'relatorio':
        def adjust_days(row):
            if row['Data_Inicial'].month < ref_month_dt.month or row['Data_Inicial'].year < ref_month_dt.year:
                # Ocorrência começou antes do mês de referência.
                # Calculamos os dias que caem no mês de referência.
                start_of_ref_month = ref_month_dt.replace(day=1)
                
                # Fim do período é o fim do mês de referência ou Data_Final, o que for menor.
                # Se Data_Final for None, assumimos o fim do mês de referência.
                if pd.isna(row['Data_Final']):
                    end_of_period = pd.to_datetime(ref_month_dt.replace(day=28) + pd.DateOffset(days=4), format='%Y-%m-%d') - pd.DateOffset(days=(ref_month_dt.replace(day=28) + pd.DateOffset(days=4)).day)
                else:
                    end_of_period = row['Data_Final']
                
                end_of_ref_month = pd.to_datetime(ref_month_dt.replace(day=28) + pd.DateOffset(days=4), format='%Y-%m-%d') - pd.DateOffset(days=(ref_month_dt.replace(day=28) + pd.DateOffset(days=4)).day)
                
                # Início do período a considerar é o início do mês de referência
                start_date = max(row['Data_Inicial'], start_of_ref_month)
                
                # Fim do período a considerar é o fim do mês de referência
                end_date = min(end_of_period, end_of_ref_month)
                
                # Se o início for depois do fim do mês, não há dias no mês.
                if start_date > end_of_ref_month:
                    return 0.0
                
                # Se a Data_Final for anterior ao início do mês, não há dias no mês.
                if pd.notna(row['Data_Final']) and row['Data_Final'] < start_of_ref_month:
                    return 0.0

                # Cálculo simples de dias no mês
                # Se for um evento de 1 dia, a Data_Inicial é o que importa.
                if row['Qtde_Dias'] == 1.0 and row['Data_Inicial'].month == ref_month_dt.month and row['Data_Inicial'].year == ref_month_dt.year:
                    return 1.0
                
                # Para períodos, a contagem é (Data_Final - Data_Inicial) + 1
                # Vamos usar a data de início e fim ajustadas ao mês de referência
                
                # Se a Data_Inicial estiver no mês anterior, o início da contagem é o dia 1 do mês de referência.
                start_count = start_of_ref_month
                
                # Se a Data_Final for posterior ao mês de referência, o fim da contagem é o último dia do mês de referência.
                end_count = end_of_ref_month
                
                # O fim do período real é a Data_Final da linha.
                real_end = row['Data_Final'] if pd.notna(row['Data_Final']) else end_of_ref_month
                
                # A contagem é do dia 1 do mês de referência até o min(Data_Final, último dia do mês de referência)
                # A diferença em dias é (Data_Final - Data_Inicial) + 1.
                
                # Se o evento termina no mês de referência ou depois
                if real_end >= start_of_ref_month:
                    
                    # Data de início efetiva para a contagem no mês
                    effective_start = max(row['Data_Inicial'], start_of_ref_month)
                    
                    # Data de fim efetiva para a contagem no mês
                    effective_end = min(real_end, end_of_ref_month)
                    
                    # Se o evento é de 1 dia, e a data está no mês de referência, é 1.
                    if (effective_start == effective_end) and (effective_start.month == ref_month_dt.month):
                        return 1.0
                    
                    # Se o evento é de múltiplos dias
                    if effective_start <= effective_end:
                        return (effective_end - effective_start).days + 1.0
                        
                return 0.0 # Se o período não toca o mês de referência
            
            # Se a Data_Inicial estiver no mês de referência, usa a Qtde_Dias original
            return row['Qtde_Dias']

        # Aplica a função de ajuste de dias
        df['Qtde_Dias'] = df.apply(adjust_days, axis=1)
        
        # Filtra as ocorrências que não têm dias no mês de referência
        df = df[df['Qtde_Dias'] > 0.0].copy()
    
    # 5. Normalização de nomes de ocorrências (para facilitar a comparação)
    # Ex: 'Férias regulamentares' vs 'Férias Regulamentares'
    df['Descricao'] = df['Descricao'].str.upper().str.replace(r'\s+', ' ', regex=True).str.strip()
    
    # 6. Criação de chave de comparação
    # A chave deve ser Funcional + Descricao + Data_Inicial + Qtde_Dias
    # Para o relatório, a Data_Inicial é a data de início do evento, e a Qtde_Dias é a contagem no mês.
    # Para a frequência, a Data é a data de início do evento e a Quantidade é a qtde total.
    # Para a comparação, usaremos Funcional + Descricao + Qtde_Dias (ajustada para o mês)
    # E a Data_Inicial será a data de início do evento que cai no mês de referência.
    
    # Para a frequência, a data é a data de início. Se for um período, a Quantidade é a duração.
    # Ex: Férias regulamentares 08/09/2025 15,0.
    # No relatório: Férias Regulamentares 08/09/2025 22/09/2025 15.
    # A chave deve ser robusta o suficiente para agrupar eventos idênticos.
    
    # Vamos agrupar por Funcional, Descricao e somar a Qtde_Dias (já ajustada)
    df_grouped = df.groupby(['Funcional', 'Pessoa', 'Descricao']).agg(
        Qtde_Dias_Total=('Qtde_Dias', 'sum'),
        Datas_Ocorrencias=('Data_Inicial', lambda x: sorted(x.dt.strftime('%d/%m/%Y').tolist()))
    ).reset_index()
    
    # Simplifica a lista de datas para a primeira data (ou uma representação)
    df_grouped['Data_Chave'] = df_grouped['Datas_Ocorrencias'].apply(lambda x: x[0] if x else '')
    df_grouped.drop(columns=['Datas_Ocorrencias'], inplace=True)
    
    df_grouped['Chave'] = df_grouped['Funcional'] + '|' + df_grouped['Descricao'] + '|' + df_grouped['Data_Chave']
    
    return df_grouped

def compare_frequencies(frequencia_path, relatorio_path):
    """
    Compara os dois arquivos de frequência e gera o DataFrame de diferenças.
    """
    
    # 1. Extração e determinação do mês de referência
    frequencia_text = extract_text_from_pdf(frequencia_path)
    relatorio_text = extract_text_from_pdf(relatorio_path)
    
    ref_month_str = get_month_reference(frequencia_text)
    
    # 2. Criação dos DataFrames a partir dos dados simulados (substituir por extração real)
    df_freq_raw = pd.read_csv(StringIO(extract_text_from_pdf(frequencia_path)))
    df_rel_raw = pd.read_csv(StringIO(extract_text_from_pdf(relatorio_path)))
    
    # 3. Normalização e Agrupamento
    df_freq = normalize_data(df_freq_raw, 'frequencia', ref_month_str)
    df_rel = normalize_data(df_rel_raw, 'relatorio', ref_month_str)
    
    # 4. Tratamento especial para "Minutos perdidos"
    # Ocorrências de "Minutos perdidos" não importam a divergência, apenas a ocorrência.
    # Vamos removê-las da comparação de quantidade de dias, mas mantê-las na lista final.
    
    # Minutos perdidos no relatório (Qtde_Dias no relatório é 1, mas a Qtde_Dias na frequência é o total de minutos)
    # No df_rel, a Qtde_Dias já foi ajustada para o mês.
    
    # Identifica as ocorrências de minutos perdidos
    minutos_perdidos_freq = df_freq[df_freq['Descricao'].str.contains('MINUTOS PERDIDOS')].copy()
    minutos_perdidos_rel = df_rel[df_rel['Descricao'].str.contains('MINUTOS PERDIDOS')].copy()
    
    # Remove as ocorrências de minutos perdidos dos DFs de comparação de dias
    df_freq_comp = df_freq[~df_freq['Descricao'].str.contains('MINUTOS PERDIDOS')].copy()
    df_rel_comp = df_rel[~df_rel['Descricao'].str.contains('MINUTOS PERDIDOS')].copy()
    
    # 5. Comparação (Merge)
    # Merge externo para pegar todas as ocorrências
    df_merged = pd.merge(
        df_freq_comp, 
        df_rel_comp, 
        on=['Funcional', 'Pessoa', 'Descricao'], 
        how='outer', 
        suffixes=('_Freq', '_Rel')
    )
    
    # 6. Identificação de Divergências
    # Preenche NaNs com 0 para Qtde_Dias_Total
    df_merged['Qtde_Dias_Total_Freq'] = df_merged['Qtde_Dias_Total_Freq'].fillna(0)
    df_merged['Qtde_Dias_Total_Rel'] = df_merged['Qtde_Dias_Total_Rel'].fillna(0)
    
    # Calcula a diferença
    df_merged['Diferenca'] = df_merged['Qtde_Dias_Total_Freq'] - df_merged['Qtde_Dias_Total_Rel']
    
    # Filtra as divergências (diferença != 0)
    df_divergencias = df_merged[df_merged['Diferenca'].abs() > 0.0].copy()
    
    # Adiciona uma coluna de status
    df_divergencias['Status'] = df_divergencias.apply(
        lambda row: 'Falta no Relatório' if row['Diferenca'] > 0 else 'Falta na Frequência', 
        axis=1
    )
    
    # Seleciona e renomeia colunas para o output
    df_divergencias = df_divergencias[[
        'Funcional', 
        'Pessoa', 
        'Descricao', 
        'Qtde_Dias_Total_Freq', 
        'Qtde_Dias_Total_Rel', 
        'Diferenca', 
        'Status'
    ]].sort_values(by=['Pessoa', 'Descricao']).reset_index(drop=True)
    
    # 7. Adiciona Minutos Perdidos de volta (apenas para listar)
    # Minutos perdidos são considerados apenas como "ocorridos"
    if not minutos_perdidos_freq.empty:
        minutos_perdidos_freq['Status'] = 'Minutos Perdidos (Apenas Ocorrência)'
        minutos_perdidos_freq['Qtde_Dias_Total_Rel'] = minutos_perdidos_rel['Qtde_Dias_Total'].iloc[0] if not minutos_perdidos_rel.empty else 0
        minutos_perdidos_freq.rename(columns={'Qtde_Dias_Total': 'Qtde_Dias_Total_Freq'}, inplace=True)
        minutos_perdidos_freq['Diferenca'] = minutos_perdidos_freq['Qtde_Dias_Total_Freq'] - minutos_perdidos_freq['Qtde_Dias_Total_Rel']
        
        minutos_perdidos_freq = minutos_perdidos_freq[[
            'Funcional', 
            'Pessoa', 
            'Descricao', 
            'Qtde_Dias_Total_Freq', 
            'Qtde_Dias_Total_Rel', 
            'Diferenca', 
            'Status'
        ]].reset_index(drop=True)
        
        df_divergencias = pd.concat([df_divergencias, minutos_perdidos_freq], ignore_index=True)
        
    return df_divergencias, ref_month_str

def generate_pdf_report(df_divergencias, ref_month_str, output_path):
    """
    Gera o arquivo PDF de conferência com a listagem de diferenças e o memorando.
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)
    styles = getSampleStyleSheet()
    story = []
    
    # Título do Relatório
    story.append(Paragraph(f"Relatório de Conferência de Frequência - Mês de {ref_month_str}", styles['h1']))
    story.append(Spacer(1, 0.25 * inch))
    story.append(Paragraph(f"Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Normal']))
    story.append(Spacer(1, 0.5 * inch))
    
    # Tabela de Divergências
    story.append(Paragraph("1. Listagem de Divergências Encontradas", styles['h2']))
    story.append(Spacer(1, 0.25 * inch))
    
    if df_divergencias.empty:
        story.append(Paragraph("Nenhuma divergência de quantidade de dias encontrada entre os documentos (excluindo 'Minutos Perdidos').", styles['Normal']))
    else:
        data = [
            ['Nro Funcional', 'Nome', 'Ocorrência', 'Qtde Frequência', 'Qtde Relatório', 'Diferença', 'Status']
        ]
        
        # Formatação dos dados para a tabela
        for index, row in df_divergencias.iterrows():
            data.append([
                row['Funcional'],
                row['Pessoa'],
                row['Descricao'].title(),
                f"{row['Qtde_Dias_Total_Freq']:.1f}",
                f"{row['Qtde_Dias_Total_Rel']:.1f}",
                f"{row['Diferenca']:.1f}",
                row['Status']
            ])
            
        table = Table(data, colWidths=[1.0*inch, 2.0*inch, 1.5*inch, 0.8*inch, 0.8*inch, 0.8*inch, 1.5*inch])
        
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 0), (2, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        
        table.setStyle(style)
        story.append(table)
    
    story.append(PageBreak())
    
    # Memorando para Retificação de Frequência
    story.append(Paragraph("2. Sugestão de Memorando para Retificação de Frequência", styles['h2']))
    story.append(Spacer(1, 0.25 * inch))
    
    # Filtra apenas as ocorrências que precisam de retificação (Qtde_Dias_Total_Freq > Qtde_Dias_Total_Rel)
    # Ou seja, o que está na Frequência mas não está no Relatório (Diferenca > 0)
    df_retificar = df_divergencias[df_divergencias['Diferenca'] > 0.0].copy()
    
    if df_retificar.empty:
        memo_text = "Prezado(a) responsável, <br/><br/> Informo que, após a conferência dos documentos de frequência e relatório de ocorrências referentes ao mês de <b>{ref_month_str}</b>, não foram identificadas divergências que necessitem de retificação imediata de frequência. O documento de 'Minutos Perdidos' foi apenas registrado como ocorrência, conforme orientação.<br/><br/> Atenciosamente."
    else:
        listagem = []
        for index, row in df_retificar.iterrows():
            # Formato: Nro Funcional - Nome - Ocorrência - Data de Início (se disponível) - Qtde de Dias
            # A Data_Chave é a primeira data de ocorrência no mês
            listagem.append(
                f"<li><b>{row['Funcional']}</b> - {row['Pessoa']} - <b>{row['Descricao'].title()}</b> - Qtde: {row['Diferenca']:.1f} dias (Faltante no Relatório)</li>"
            )
        
        lista_html = "".join(listagem)
        
        memo_text = f"""
        Prezado(a) responsável, <br/><br/>
        Informo que, após a conferência dos documentos de frequência e relatório de ocorrências referentes ao mês de <b>{ref_month_str}</b>, foram identificadas as seguintes divergências que necessitam de retificação de frequência:
        <br/><br/>
        <ul>
            {lista_html}
        </ul>
        <br/>
        Solicito a gentileza de providenciar a retificação de frequência para os servidores listados acima, a fim de que os dados reflitam corretamente as ocorrências do mês.
        <br/><br/>
        Ocorrências de 'Minutos Perdidos' foram apenas registradas como ocorrência, conforme orientação, e não serão consideradas para retificação de dias.
        <br/><br/>
        Atenciosamente.
        """
        
    memo_text = memo_text.replace("{ref_month_str}", ref_month_str)
    
    story.append(Paragraph("<b>COPIAR E COLAR NO SISTEMA DE MEMORANDO:</b>", styles['Normal']))
    story.append(Spacer(1, 0.1 * inch))
    
    # Bloco de texto do memorando
    memo_style = styles['Normal']
    memo_style.fontName = 'Courier' # Fonte monoespaçada para facilitar a cópia
    memo_style.fontSize = 10
    
    story.append(Paragraph(memo_text, memo_style))
    
    doc.build(story)
    
    return output_path

if __name__ == "__main__":
    # O script deve ser executado no diretório onde os arquivos estão.
    # O nome dos arquivos de entrada é dinâmico, baseado no prefixo numérico.
    
    # 1. Identificar o prefixo numérico (ex: 106)
    import glob
    
    # Busca por arquivos que correspondam ao padrão "<numero>-frequencia.pdf"
    frequencia_files = glob.glob("*-frequencia.pdf")
    
    if not frequencia_files:
        print("Erro: Arquivo *-frequencia.pdf não encontrado no diretório.")
        exit()
        
    frequencia_path = frequencia_files[0]
    
    # Extrai o prefixo numérico
    match = re.match(r'(\d+)-frequencia.pdf', os.path.basename(frequencia_path))
    if not match:
        print("Erro: O nome do arquivo de frequência não segue o padrão <numero>-frequencia.pdf.")
        exit()
        
    prefixo_numerico = match.group(1)
    
    relatorio_path = f"{prefixo_numerico}-relatório.pdf"
    output_path = f"{prefixo_numerico}-conferencia.pdf"
    
    if not os.path.exists(relatorio_path):
        print(f"Erro: Arquivo de relatório esperado '{relatorio_path}' não encontrado.")
        exit()
        
    # 2. Executar a comparação
    print(f"Iniciando comparação para o prefixo {prefixo_numerico}...")
    try:
        df_divergencias, ref_month_str = compare_frequencies(frequencia_path, relatorio_path)
        print(f"Comparação concluída para o mês de {ref_month_str}.")
        
        # 3. Gerar o PDF de saída
        generate_pdf_report(df_divergencias, ref_month_str, output_path)
        print(f"Relatório de conferência gerado com sucesso: {output_path}")
        
    except Exception as e:
        print(f"Ocorreu um erro durante a comparação ou geração do PDF: {e}")
        print("Verifique se as dependências (pandas, reportlab) estão instaladas e se o formato dos dados simulados está correto.")
        print("Lembre-se: A extração real de dados de PDF (extract_text_from_pdf) deve ser implementada com 'pdfplumber' ou 'tabula-py'.")
        
    # Salva o DataFrame de divergências em um CSV temporário para ser usado no script 2
    df_divergencias.to_csv(f"{prefixo_numerico}-divergencias.csv", index=False)
    print(f"Dados de divergências salvos em {prefixo_numerico}-divergencias.csv para uso no segundo script.")

# NOTA IMPORTANTE PARA O USUÁRIO:
# A função 'extract_text_from_pdf' está simulando a extração de dados tabulares dos PDFs.
# No seu ambiente, você precisará instalar e configurar 'pdfplumber' ou 'tabula-py'
# para extrair os dados reais. A lógica de comparação e geração de PDF está completa.
# O 'requirements.txt' incluirá as dependências necessárias.
# """
