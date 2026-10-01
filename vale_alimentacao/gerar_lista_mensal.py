import os
import glob
import re
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
from src.gerador_lista_va import processar_mes

def main():
    print("=" * 60)
    print(" AUTOMATIZADOR DE LISTA MENSAL - VALE ALIMENTAÇÃO ")
    print(" Prefeitura do Município de Piracicaba ")
    print("=" * 60)

    # 1. Localiza os PDFs disponíveis em data/
    pdfs = glob.glob(os.path.join("data", "*.pdf"))
    if not pdfs:
        print("❌ ERRO: Nenhum arquivo PDF encontrado na pasta 'data/'.")
        print("Por favor, coloque o relatório 'Funcionário por Cargo - MM-AAAA.pdf' na pasta 'data/'.")
        sys.exit(1)

    # Ordena para pegar o mais recente ou o primeiro
    pdf_selecionado = sorted(pdfs)[-1]
    nome_pdf = os.path.basename(pdf_selecionado)
    print(f"📄 Arquivo PDF detectado: {nome_pdf}")

    # Tenta extrair mês e ano do nome do arquivo (ex: "09-2026")
    m = re.search(r'(\d{2})[-_](\d{4})', nome_pdf)
    if m:
        mes_num, ano_num = m.group(1), m.group(2)
        meses_extenso = {
            "01": "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril",
            "05": "Maio", "06": "Junho", "07": "Julho", "08": "Agosto",
            "09": "Setembro", "10": "Outubro", "11": "Novembro", "12": "Dezembro"
        }
        mes_nome = meses_extenso.get(mes_num, mes_num)
        mes_ref = f"{mes_nome}/{ano_num}"
    else:
        mes_ref = "Setembro/2026"

    print(f"📅 Mês de Referência: {mes_ref}")
    print("\nProcessando cruzamento e gerando folhas de entrega...")

    try:
        arquivo_gerado = processar_mes(pdf_selecionado, mes_ano_referencia=mes_ref)
        print("\n" + "=" * 60)
        print("🎉 PROCESSO CONCLUÍDO COM SUCESSO!")
        print(f"📁 Planilha salva em: {arquivo_gerado}")
        print("=" * 60)
    except Exception as e:
        print(f"\n❌ ERRO durante o processamento: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
