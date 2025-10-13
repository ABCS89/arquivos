# instalar biblioteca - tabula-py
# isntalar o java (jdk)
# importar a biblioteca tabula
# extrair a tabela do PDF no excel

import converter_PDF_XLSX.convert_tabula as convert_tabula

# extrair as tabelas do pdf usandoo tabula-py

tabela = convert_tabula.read_pdf("cns_relatorio_ocorrencia_geral.pdf", pages="all", multiple_tables=True)
print(tabela)
