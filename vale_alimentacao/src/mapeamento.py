import os
import unicodedata
import re
import pandas as pd

def normalizar_texto(texto):
    if pd.isna(texto) or texto is None:
        return ""
    t = str(texto).strip().upper()
    t = unicodedata.normalize('NFKD', t).encode('ASCII', 'ignore').decode('ASCII')
    return " ".join(t.split())

def padronizar_empresa_va(empresa_str):
    if pd.isna(empresa_str) or not str(empresa_str).strip():
        return "NÃO INFORMADA"
    
    e_norm = normalizar_texto(empresa_str)
    
    if "PLUXEE" in e_norm or "SODEXO" in e_norm:
        return "Pluxee – Sodexo"
    elif "VR" in e_norm:
        return "VR Benefícios e Serv. de processamento S/A"
    elif "USECRED" in e_norm:
        return "USECRED"
    elif "TICKET" in e_norm:
        return "Ticket Serviços S/A"
    elif "VERO" in e_norm:
        return "VEROCARD"
    elif "MEGAVALE" in e_norm:
        return "Megavale Card"
    elif "LECARD" in e_norm or "LE CARD" in e_norm:
        return "Le Card Adm.de Cartões Ltda"
    elif "EBA" in e_norm:
        return "EBA Empresa de Benefícios Amigáveis S/A"
    elif "VOLUS" in e_norm:
        return "VÓLUS"
    elif "UP" in e_norm:
        return "UP"
    elif "NAO OPTOU" in e_norm:
        return "Não optou por nenhuma"
    else:
        return str(empresa_str).strip()

def carregar_mapa_secretarias(caminho_ods="data/siglas secretarias.ods"):
    """
    Retorna dois dicionarios:
    - por_codigo: { 107: '107 - SECRETARIA MUNICIPAL DE EDUCAÇÃO', ... }
    - por_termo_norm: { 'EDUCACAO': '107 - SECRETARIA MUNICIPAL DE EDUCAÇÃO', '107': ..., 'SEMTRE': ... }
    """
    mapa_codigo = {}
    mapa_termos = {}

    if not os.path.exists(caminho_ods):
        return mapa_codigo, mapa_termos

    df = pd.read_excel(caminho_ods, engine="odf")
    for _, row in df.iterrows():
        txt = str(row['SECRETARIA']).strip()
        # Ex: "107 – Educação", "109 – Desenvolvimento Social – ASOCIAL", "116 – Guarda Civil"
        m = re.match(r'^(\d{3})\s*[-–—]\s*(.+)$', txt)
        if m:
            cod = int(m.group(1))
            desc = m.group(2).strip()
            nome_oficial = f"{cod} - SECRETARIA MUNICIPAL DE {desc.upper()}"
            if "GUARDA CIVIL" in desc.upper():
                nome_oficial = f"{cod} - GUARDA CIVIL DO MUNICÍPIO DE PIRACICABA"
            elif "PROCURADORIA" in desc.upper():
                nome_oficial = f"{cod} - PROCURADORIA GERAL"
            elif "CORREGEDORIA" in desc.upper():
                nome_oficial = f"{cod} - CORREGEDORIA GERAL DO MUNICÍPIO"
            elif "GABINETE" in desc.upper():
                nome_oficial = f"{cod} - GABINETE DO PREFEITO"
            elif not desc.upper().startswith("SECRETARIA"):
                nome_oficial = f"{cod} - SECRETARIA MUNICIPAL DE {desc.upper()}"

            mapa_codigo[cod] = nome_oficial
            mapa_termos[str(cod)] = nome_oficial
            
            # Adiciona siglas e termos
            partes = re.split(r'[-–—]', txt)
            for p in partes:
                p_norm = normalizar_texto(p)
                if p_norm:
                    mapa_termos[p_norm] = nome_oficial

    # Mapeamentos diretos extras comuns
    extras = {
        "EDUCACAO": mapa_codigo.get(107, "107 - SECRETARIA MUNICIPAL DE EDUCAÇÃO"),
        "SAUDE": mapa_codigo.get(114, "114 - SECRETARIA MUNICIPAL DE SAÚDE"),
        "ADMGOV": mapa_codigo.get(102, "102 - SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO E GOVERNO"),
        "ADMINISTRACAO": mapa_codigo.get(102, "102 - SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO E GOVERNO"),
        "SEMAD": mapa_codigo.get(102, "102 - SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO E GOVERNO"),
        "PGM": mapa_codigo.get(103, "103 - PROCURADORIA GERAL"),
        "PROCURADORIA": mapa_codigo.get(103, "103 - PROCURADORIA GERAL"),
        "HABITACAO": mapa_codigo.get(104, "104 - SECRETARIA MUNICIPAL DE HABITAÇÃO E REGULARIZAÇÃO FUNDIÁRIA"),
        "FINANCAS": mapa_codigo.get(106, "106 - SECRETARIA MUNICIPAL DE FINANÇAS"),
        "OBRAS": mapa_codigo.get(108, "108 - SECRETARIA MUNICIPAL DE OBRAS, INFRAESTRUTURA E SERVIÇOS PÚBLICOS"),
        "SMADS": mapa_codigo.get(109, "109 - SECRETARIA MUNICIPAL DE ASSISTÊNCIA, DESENVOLVIMENTO SOCIAL E FAMÍLIA"),
        "ASOCIAL": mapa_codigo.get(109, "109 - SECRETARIA MUNICIPAL DE ASSISTÊNCIA, DESENVOLVIMENTO SOCIAL E FAMÍLIA"),
        "AGRIMA": mapa_codigo.get(110, "110 - SECRETARIA MUNICIPAL DE AGRICULTURA, ABASTECIMENTO E MEIO AMBIENTE"),
        "CULTURA": mapa_codigo.get(112, "112 - SECRETARIA MUNICIPAL DE CULTURA"),
        "TURISMO": mapa_codigo.get(113, "113 - SECRETARIA MUNICIPAL DE TURISMO"),
        "GUARDA": mapa_codigo.get(116, "116 - GUARDA CIVIL DO MUNICÍPIO DE PIRACICABA"),
        "SELAM": mapa_codigo.get(119, "119 - SECRETARIA MUNICIPAL DE ESPORTES, LAZER E ATIVIDADES MOTORAS"),
        "SEMDEC": mapa_codigo.get(120, "120 - SECRETARIA MUNICIPAL DE DESENVOLVIMENTO ECONÔMICO, INDÚSTRIA E COMÉRCIO"),
        "CORREGEDORIA": mapa_codigo.get(121, "121 - CORREGEDORIA GERAL DO MUNICÍPIO"),
        "GABINETE": mapa_codigo.get(122, "122 - GABINETE DO PREFEITO"),
        "PARCERIAS": mapa_codigo.get(123, "123 - SECRETARIA MUNICIPAL DE CIDADANIA E PARCERIAS"),
        "CIDADANIA": mapa_codigo.get(123, "123 - SECRETARIA MUNICIPAL DE CIDADANIA E PARCERIAS"),
        "SEGTRANS": mapa_codigo.get(124, "124 - SECRETARIA MUNICIPAL DE SEGURANÇA PÚBLICA, TRÂNSITO E TRANSPORTES"),
        "SEMTRE": mapa_codigo.get(125, "125 - SECRETARIA MUNICIPAL DE TRABALHO, EMPREGO E RENDA")
    }
    for k, v in extras.items():
        if k not in mapa_termos:
            mapa_termos[k] = v

    return mapa_codigo, mapa_termos

def identificar_secretaria(valor_sec, mapa_codigo, mapa_termos):
    if pd.isna(valor_sec) or valor_sec is None:
        return "NÃO IDENTIFICADA"
    
    # Se for int ou float
    try:
        cod_int = int(float(valor_sec))
        if cod_int in mapa_codigo:
            return mapa_codigo[cod_int]
    except (ValueError, TypeError):
        pass

    # Limpa prefixos de operadoras comumente encontrados em nomes de abas
    v_limpo = str(valor_sec)
    for prefixo in ["USECRED", "PLUXEE", "VEROCHEQUE", "VEROCARD", "VR", "LECARD", "TICKET", "MEGAVALE", "VOLUS", "EBA", "UP"]:
        v_limpo = re.sub(rf'\b{prefixo}\b', '', v_limpo, flags=re.IGNORECASE)
        
    v_norm = normalizar_texto(v_limpo)
    if v_norm in mapa_termos:
        return mapa_termos[v_norm]
    
    for k, v in mapa_termos.items():
        if k in v_norm or v_norm in k:
            return v
            
    return str(valor_sec).strip()

def obter_centro_custo_usecred(secretaria_str):
    """
    Regra USECRED:
    - 107 / Educação -> Centro de Custo: Educação
    - 114 / Saúde -> Centro de Custo: Saúde
    - Demais secretarias -> Centro de Custo: ADMGOV
    """
    sec_norm = normalizar_texto(secretaria_str)
    if "107" in sec_norm or "EDUCACAO" in sec_norm:
        return "Educação"
    elif "114" in sec_norm or "SAUDE" in sec_norm:
        return "Saúde"
    else:
        return "ADMGOV"

def verificar_troca_centro_custo_usecred(sec_anterior, sec_nova):
    """
    Verifica se houve alteração de Centro de Custo entre secretarias para a USECRED.
    Retorna: (houve_troca: bool, cc_antigo: str, cc_novo: str)
    """
    cc_ant = obter_centro_custo_usecred(sec_anterior)
    cc_novo = obter_centro_custo_usecred(sec_nova)
    return (cc_ant != cc_novo), cc_ant, cc_novo

