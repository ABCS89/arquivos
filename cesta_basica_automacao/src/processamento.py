"""
Etapa 2: agrupar os funcionários lidos da Relação Mensal por código de secretaria.
"""

from collections import defaultdict


def agrupar_por_secretaria(registros):
    """
    Recebe a lista de registros (funcional, nome, codigo_secretaria) e devolve
    um dicionário {codigo_secretaria: [registros...]}.
    """
    grupos = defaultdict(list)
    for registro in registros:
        grupos[registro["codigo_secretaria"]].append(registro)
    return grupos
