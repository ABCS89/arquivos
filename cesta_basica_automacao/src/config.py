"""
Configurações centralizadas da automação de Cesta Básica.

Se algum nome de coluna, aba ou linha mudar no arquivo original,
ajuste apenas aqui — o resto do código não precisa ser tocado.
"""

from pathlib import Path

# --------------------------------------------------------------------
# Pastas do projeto
# --------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent

ENTRADA_DIR = BASE_DIR / "entrada"
SAIDA_DIR = BASE_DIR / "saida"
SAIDA_SECRETARIAS_DIR = SAIDA_DIR / "secretarias"

# --------------------------------------------------------------------
# Nomes das colunas usadas na Relação Mensal
# --------------------------------------------------------------------
COLUNA_FUNCIONAL = "Nro Funcional"
COLUNA_NOME = "Nome"
COLUNA_SECRETARIA = "Cód. Secretaria"

# --------------------------------------------------------------------
# Layout do cesta_basica.xlsx
# --------------------------------------------------------------------
# Aba que recebe TODOS os funcionários (cópia integral da Relação Mensal)
ABA_CONSOLIDADA = "Planilha1"

# Em cada aba de secretaria:
#   linha 1 -> e-mail de contato (célula mesclada A1:C1)
#   linha 2 -> cabeçalho (Nro Funcional / Nome / Cód. Secretaria)
#   linha 3 em diante -> dados dos funcionários daquela secretaria
LINHA_EMAIL = 1
LINHA_CABECALHO = 2
LINHA_INICIO_DADOS = 3
