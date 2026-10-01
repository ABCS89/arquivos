# Projeto: Automação do Vale Alimentação

Automação para geração e conferência mensal da lista de novos cartões de Vale Alimentação da Prefeitura do Município de Piracicaba.

---

## 🚀 Como Executar Todo Mês

1. **Coloque os arquivos do mês na pasta `data/`:**
   - O relatório de admitidos em PDF: `Funcionário por Cargo - MM-AAAA.pdf` (ex: `Funcionário por Cargo - 10-2026.pdf`).
   - A planilha atualizada de opções: `data/novos_e_alteracoes/Escolhas das empresas de Vale Alimentação.ods`.

2. **Execute o comando no terminal:**
   ```powershell
   .\.venv\Scripts\python gerar_lista_mensal.py
   ```

3. **Confira o resultado na pasta `output/`:**
   - O arquivo gerado terá o nome `output/lista_VA_[mes-ano].xlsx` com três abas:
     - **`Consulta`**: Folha oficial por Secretaria, pronta para impressão e colheita de assinatura na entrega dos cartões.
     - **`Planilha1`**: Base cadastral consolidada com a coluna `recebido` para controle interno.
     - **`Pendências e Alertas`**: Lista de servidores admitidos que ainda não escolheram operadora ou dispensaram o benefício.

