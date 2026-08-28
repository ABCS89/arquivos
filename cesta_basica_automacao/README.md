# Automação Cesta Básica

Automatiza o processo de:
1. Pegar `Nro Funcional`, `Nome` e `Cód. Secretaria` da Relação Mensal.
2. Colar esses dados no `cesta_basica.xlsx`, tanto na aba consolidada (`Planilha1`)
   quanto na aba de cada secretaria.
3. Separar cada aba de secretaria em um arquivo `.xlsx` próprio, pronto para
   anexar no e-mail daquela secretaria.
4. Gerar uma lista (`lista_envio.md`) com o e-mail e a quantidade de
   funcionários de cada secretaria.

## Estrutura de pastas

```
cesta_basica_automacao/
├── main.py                 # roda tudo (ponto de entrada)
├── requirements.txt
├── README.md
├── entrada/                 # <- todo mês, coloque aqui os 2 arquivos de origem
│   ├── Relação Mensal - MMAAAA.xlsx
│   └── cesta_basica.xlsx
├── saida/                   # gerado automaticamente a cada execução
│   ├── cesta_basica - MM-AAAA.xlsx      (arquivo consolidado, todas as abas)
│   ├── lista_envio.md                   (resumo para te ajudar a enviar os e-mails)
│   └── secretarias/
│       ├── Cesta Basica - 102 - MM-AAAA.xlsx
│       ├── Cesta Basica - 106 - MM-AAAA.xlsx
│       └── ... (um arquivo por secretaria)
└── src/
    ├── config.py            # nomes de colunas, abas e pastas (ajuste tudo aqui)
    ├── leitura.py            # etapa 1: acha os arquivos e lê a Relação Mensal
    ├── processamento.py      # etapa 2: agrupa os funcionários por secretaria
    └── exportacao.py         # etapa 3: preenche o cesta_basica.xlsx e separa os arquivos
```

## Como configurar (só na primeira vez)

Dentro da pasta `cesta_basica_automacao/`:

```bash
# 1. Criar o ambiente virtual
python3 -m venv venv

# 2. Ativar o ambiente virtual
# Linux/Mac:
source venv/bin/activate
# Windows:
venv\Scripts\activate

# 3. Instalar as dependências
pip install -r requirements.txt
```

## Como usar todo mês

1. Ative o ambiente virtual (se não estiver ativo): `source venv/bin/activate`
2. Copie os 2 arquivos do mês para dentro de `entrada/`:
   - a Relação Mensal (o nome do arquivo precisa conter a palavra "Relação")
   - o `cesta_basica.xlsx` (o nome precisa conter a palavra "cesta")
3. Rode:
   ```bash
   python main.py
   ```
4. Confira a pasta `saida/`:
   - `secretarias/` tem um `.xlsx` por secretaria — é o que você anexa em cada e-mail.
   - `lista_envio.md` tem o e-mail e a quantidade de funcionários de cada uma,
     para te ajudar a montar o envio manual.

## Observações importantes

- O e-mail (linha 1) e o cabeçalho (linha 2) de cada aba do `cesta_basica.xlsx`
  **não são apagados nem alterados** — o script só escreve a partir da linha 3.
- Se a Relação Mensal trouxer um código de secretaria que **não existe** como
  aba no `cesta_basica.xlsx`, o script avisa no terminal (não trava, mas
  aqueles funcionários não são distribuídos até você criar a aba).
- Se uma secretaria **não tiver nenhum funcionário** no mês, o script também
  avisa — o arquivo dela ainda é gerado, só que vazio.
- O nome dos arquivos de saída usa a competência (mês/ano) detectada a partir
  do nome do arquivo da Relação Mensal (ex.: `082026` → `08-2026`). Se o
  script não conseguir identificar, usa `sem-data` — nesse caso vale renomear
  o arquivo de entrada para incluir o padrão `MMAAAA` no nome.
- Toda vez que você rodar `python main.py`, os arquivos em `saida/` são
  sobrescritos — não precisa apagar nada manualmente entre execuções.
