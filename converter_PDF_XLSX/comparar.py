import pandas as pd

def compare_frequencia_data(frequencia_path, relatorio_path):
    df_frequencia = pd.read_excel(frequencia_path)
    df_relatorio = pd.read_excel(relatorio_path)

    # --- Limpeza e Padronização df_frequencia ---
    df_frequencia_clean = df_frequencia.copy()
    df_frequencia_clean = df_frequencia_clean.rename(columns={
        'Nro Funcional': 'Funcionario',
        # 'Nome': 'Nome',
        'Ocorrência': 'Ocorrencia',
        # 'Data': 'Data_Frequencia',
        'Quantidade': 'Quantidade_Frequencia'
    })
    df_frequencia_clean['ID'] = df_frequencia_clean['ID'].astype(str).str.strip()
    df_frequencia_clean['Nome'] = df_frequencia_clean['Nome'].str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()
    df_frequencia_clean['Ocorrencia'] = df_frequencia_clean['Ocorrencia'].str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()
    df_frequencia_clean['Data_Frequencia'] = pd.to_datetime(df_frequencia_clean['Data_Frequencia'], errors='coerce', format='%d/%m/%Y')
    df_frequencia_clean['Quantidade_Frequencia'] = pd.to_numeric(df_frequencia_clean['Quantidade_Frequencia'].astype(str).str.replace(',', '.'), errors='coerce')

    # Filtrar apenas ocorrências que não são 'FREQUÊNCIA NORMAL' para comparação mais detalhada
    df_frequencia_events = df_frequencia_clean[df_frequencia_clean['Ocorrencia'] != 'FREQUÊNCIA NORMAL'].copy()

    # --- Limpeza e Padronização df_relatorio ---
    df_relatorio_clean = df_relatorio.copy()
    # Remover linhas de cabeçalho/rodapé que não contêm dados de funcionário
    df_relatorio_clean = df_relatorio_clean[df_relatorio_clean['Funcionário'].notna()]
    df_relatorio_clean = df_relatorio_clean.rename(columns={
        'Funcionário': 'ID',
        # 'Pessoa': 'Nome',
        'Descrição': 'Ocorrencia',
        # 'Data Inicial': 'Data_Inicio_Relatorio',
        # 'Data Final': 'Data_Fim_Relatorio',
        'Qtde Dias': 'Quantidade_Dias_Relatorio'
    })
    df_relatorio_clean['ID'] = df_relatorio_clean['ID'].astype(str).str.strip()
    # df_relatorio_clean['Nome'] = df_relatorio_clean['Nome'].str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()
    df_relatorio_clean['Ocorrencia'] = df_relatorio_clean['Ocorrencia'].str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()
    df_relatorio_clean['Data_Inicio_Relatorio'] = pd.to_datetime(df_relatorio_clean['Data_Inicio_Relatorio'], errors='coerce', format='%d/%m/%Y')
    df_relatorio_clean['Data_Fim_Relatorio'] = pd.to_datetime(df_relatorio_clean['Data_Fim_Relatorio'], errors='coerce', format='%d/%m/%Y')
    df_relatorio_clean['Quantidade_Dias_Relatorio'] = pd.to_numeric(df_relatorio_clean['Quantidade_Dias_Relatorio'], errors='coerce')

    # --- Lógica de Comparação ---
    # O objetivo é encontrar eventos que estão em um, mas não no outro, ou que têm dados conflitantes.
    # Vamos criar uma chave de comparação para cada evento, focando em ID, Nome e Ocorrência.

    # Criar uma chave de identificação para o evento principal (ID, Nome, Ocorrencia)
    df_frequencia_events['Chave_Principal'] = df_frequencia_events['ID'] + '_' + df_frequencia_events['Nome'] + '_' + df_frequencia_events['Ocorrencia']
    df_relatorio_clean['Chave_Principal'] = df_relatorio_clean['ID'] + '_' + df_relatorio_clean['Nome'] + '_' + df_relatorio_clean['Ocorrencia']

    # Merge para encontrar correspondências e diferenças
    # Usamos um outer merge para pegar tudo que está em um ou em outro
    merged_df = pd.merge(df_frequencia_events, df_relatorio_clean,
                           on=['ID', 'Nome', 'Ocorrencia'],
                           how='outer',
                           suffixes=('_Frequencia', '_Relatorio'))

    # Identificar eventos que estão apenas na frequência
    diff_frequencia_only = merged_df[merged_df['Chave_Principal_Relatorio'].isna()].copy()
    diff_frequencia_only['Origem'] = 'Frequência (não no Relatório)'
    diff_frequencia_only = diff_frequencia_only[[
        'ID', 'Nome', 'Ocorrencia', 'Data_Frequencia', 'Quantidade_Frequencia', 'Origem'
    ]]
    diff_frequencia_only = diff_frequencia_only.rename(columns={
        'Data_Frequencia': 'Data_Inicio_Referencia',
        'Quantidade_Frequencia': 'Quantidade_Referencia'
    })
    diff_frequencia_only['Data_Fim_Referencia'] = pd.NaT # Adicionar para consistência

    # Identificar eventos que estão apenas no relatório
    diff_relatorio_only = merged_df[merged_df['Chave_Principal_Frequencia'].isna()].copy()
    diff_relatorio_only['Origem'] = 'Relatório (não na Frequência)'
    diff_relatorio_only = diff_relatorio_only[[
        'ID', 'Nome', 'Ocorrencia', 'Data_Inicio_Relatorio', 'Data_Fim_Relatorio', 'Quantidade_Dias_Relatorio', 'Origem'
    ]]
    diff_relatorio_only = diff_relatorio_only.rename(columns={
        'Data_Inicio_Relatorio': 'Data_Inicio_Referencia',
        'Data_Fim_Relatorio': 'Data_Fim_Referencia',
        'Quantidade_Dias_Relatorio': 'Quantidade_Referencia'
    })

    # Identificar eventos que estão em ambos, mas com datas ou quantidades diferentes
    # Para isso, vamos filtrar as linhas onde ambos existem e comparar os detalhes
    common_events = merged_df[merged_df['Chave_Principal_Frequencia'].notna() & merged_df['Chave_Principal_Relatorio'].notna()].copy()

    # Uma forma de comparar é verificar se a data da frequência está dentro do período do relatório
    # e se a quantidade da frequência corresponde à quantidade do relatório para aquela data.
    # Isso exigiria uma lógica mais complexa de iteração ou expansão de datas.

    # Por simplicidade inicial, vamos considerar uma diferença se as datas de início ou quantidades não correspondem
    # (assumindo que 'Data_Frequencia' deveria ser 'Data_Inicio_Relatorio' para eventos de um dia, ou dentro do período)
    
    # Vamos focar em comparar a 'Data_Frequencia' com 'Data_Inicio_Relatorio' e 'Quantidade_Frequencia' com 'Quantidade_Dias_Relatorio'
    # Para eventos que *deveriam* ser de um dia ou ter uma correspondência direta.
    
    # Para uma comparação mais robusta, seria necessário definir regras claras sobre como as datas e quantidades se relacionam.
    # Por exemplo, se uma 'Ocorrencia' na frequência é 'Abono' em '04/09/2025' com '1.0' quantidade,
    # e no relatório é 'Abono' de '04/09/2025' a '04/09/2025' com '1.0' dias.

    # Vamos criar uma chave de comparação de detalhes para eventos comuns
    common_events['Chave_Detalhes_Frequencia'] = common_events['Data_Frequencia'].dt.strftime('%Y-%m-%d').fillna('NaT') + '_' + common_events['Quantidade_Frequencia'].astype(str).fillna('NaN')
    common_events['Chave_Detalhes_Relatorio'] = common_events['Data_Inicio_Relatorio'].dt.strftime('%Y-%m-%d').fillna('NaT') + '_' + common_events['Quantidade_Dias_Relatorio'].astype(str).fillna('NaN')

    # Eventos com a mesma chave principal, mas detalhes diferentes
    diff_details = common_events[common_events['Chave_Detalhes_Frequencia'] != common_events['Chave_Detalhes_Relatorio']].copy()
    if not diff_details.empty:
        # Para cada linha de diferença, vamos criar duas entradas: uma para a frequência e outra para o relatório
        # para mostrar o que foi encontrado em cada um.
        diff_freq_part = diff_details[[
            'ID', 'Nome', 'Ocorrencia', 'Data_Frequencia', 'Quantidade_Frequencia'
        ]].copy()
        diff_freq_part['Origem'] = 'Frequência (detalhes diferentes)'
        diff_freq_part = diff_freq_part.rename(columns={
            'Data_Frequencia': 'Data_Inicio_Referencia',
            'Quantidade_Frequencia': 'Quantidade_Referencia'
        })
        diff_freq_part['Data_Fim_Referencia'] = pd.NaT

        diff_rel_part = diff_details[[
            'ID', 'Nome', 'Ocorrencia', 'Data_Inicio_Relatorio', 'Data_Fim_Relatorio', 'Quantidade_Dias_Relatorio'
        ]].copy()
        diff_rel_part['Origem'] = 'Relatório (detalhes diferentes)'
        diff_rel_part = diff_rel_part.rename(columns={
            'Data_Inicio_Relatorio': 'Data_Inicio_Referencia',
            'Data_Fim_Relatorio': 'Data_Fim_Referencia',
            'Quantidade_Dias_Relatorio': 'Quantidade_Referencia'
        })
        all_differences = pd.concat([diff_frequencia_only, diff_relatorio_only, diff_freq_part, diff_rel_part], ignore_index=True)
    else:
        all_differences = pd.concat([diff_frequencia_only, diff_relatorio_only], ignore_index=True)

    # Reordenar e selecionar colunas finais
    final_columns = [
        'ID', 'Nome', 'Ocorrencia', 'Data_Inicio_Referencia', 'Data_Fim_Referencia', 'Quantidade_Referencia', 'Origem'
    ]
    all_differences = all_differences[final_columns]

    return all_differences

if __name__ == '__main__':
    frequencia_file = 'frequencia.xlsx'  # Nome do seu arquivo de frequência
    relatorio_file = 'relatorio_ocorrencia_geral.xlsx'  # Nome do seu arquivo de relatório
    output_diff_file = 'diferencas_frequencia.xlsx'

    differences_df = compare_frequencia_data(frequencia_file, relatorio_file)
    differences_df.to_excel(output_diff_file, index=False)
    print(f'Diferenças salvas em {output_diff_file}')

