# Regras de Negócio - Vale Alimentação

Documento para mapeamento das regras de concessão, descontos, afastamentos e particularidades da lista mensal.

---

## 🏛️ De-Para Oficial: Código, Sigla e Secretaria
| Código | Sigla / Nome Aba | Nome Oficial |
| :---: | :--- | :--- |
| **102** | ADMGOV | Secretaria Municipal de Administração e Governo |
| **103** | PGM | Procuradoria Geral |
| **104** | HABITAÇÃO | Secretaria Municipal de Habitação e Regularização Fundiária |
| **106** | FINANÇAS | Secretaria Municipal de Finanças |
| **107** | EDUCAÇÃO | Secretaria Municipal de Educação |
| **107** | EDUCAÇÃO (Temporário) | Secretaria Municipal de Educação (Temporário) |
| **108** | OBRAS | Secretaria Municipal de Obras, Infraestrutura e Serviços Públicos |
| **109** | ASOCIAL | Secretaria Municipal de Assistência, Desenvolvimento Social e Família |
| **110** | AGRIMA | Secretaria Municipal de Agricultura, Abastecimento e Meio Ambiente |
| **112** | CULTURA | Secretaria Municipal de Cultura |
| **113** | TURISMO | Secretaria Municipal de Turismo |
| **114** | SAÚDE | Secretaria Municipal de Saúde |
| **116** | GUARDA | Guarda Civil do Município de Piracicaba |
| **119** | SELAM | Secretaria Municipal de Esportes, Lazer e Atividades Motoras |
| **120** | SEMDEC | Secretaria Municipal de Desenvolvimento Econômico, Indústria e Comércio |
| **121** | CORREGEDORIA | Corregedoria Geral do Município |
| **122** | GABINETE | Gabinete do Prefeito |
| **123** | CIDADANIA | Secretaria Municipal de Cidadania e Parcerias |
| **124** | SEGTRANS | Secretaria Municipal de Segurança Pública, Trânsito e Transportes |
| **125** | SEMTRE | Secretaria Municipal de Trabalho, Emprego e Renda |


---

## 1. Critérios de Elegibilidade e Emissão de Cartão
- Implementação inicial em Abril.
- Servidores que já receberam na entrega inicial.
- Novos servidores admitidos (precisam de cartão).
- Servidores com alteração de cargo (precisam de novo cartão).

---

## 💳 Regra Especial: Operadora USECRED (Centros de Custo)
A operadora **USECRED** possui uma regra particular de emissão vinculada exclusivamente aos seus **Centros de Custo contratuais**:
1. **Centros de Custo USECRED:**
   - **Educação**: Secretaria 107 (Educação / Temporários)
   - **Saúde**: Secretaria 114 (Saúde)
   - **ADMGOV**: Abrange **todas as demais secretarias** municipais (102, 103, 104, 106, 108, 109, 110, 112, 113, 116, 119, 120, 121, 122, 123, 124, 125).

2. **Gatilhos para Emissão de Novo Cartão USECRED:**
   - **Admissão:** Novo servidor admitido com opção USECRED.
   - **Troca de Operadora:** Servidor vindo de outra operadora (ex: Pluxee, VR) para a USECRED.
   - **Mudança de Centro de Custo:** Servidor que já possui USECRED e mudou de secretaria/local acarretando **troca de Centro de Custo** (ex: *Educação ➔ ADMGOV*, *Saúde ➔ Educação*, *ADMGOV ➔ Saúde*).

3. **⚠️ Exceção Importante (Quando NÃO gera novo cartão USECRED):**
   - **Troca de Cargo no mesmo Centro de Custo:** Se o servidor já possui cartão USECRED e continua dentro do **mesmo centro de custo** (mesmo que tenha mudado de cargo ou mudado de secretaria dentro de ADMGOV), **NÃO** deve ser gerado novo cartão. O saldo continua caindo no cartão USECRED atual.



## 2. Afastamentos e Licenças (Impacto no benefício)
- *(a definir)*

## 3. Faltas e Descontos Proporcionais
- *(a definir)*

## 4. Fontes de Dados / Arquivos de Entrada
- `data/entrega_inicial/`: Planilhas `.ods` separadas por empresa com servidores divididos por secretaria.
- `data/novos_e_alteracoes/`: Planilha/arquivo com novos servidores e trocas de cargo.

## 5. Formato de Saída Esperado
- *(a definir)*

