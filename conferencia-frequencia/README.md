# Conferência de Frequência — GCM 116

Confere, mês a mês, o relatório de frequência gerado pela **secretaria**
("Frequência dos Funcionários") contra o relatório gerado pelo **sistema**
("Relatório Ocorrência Geral"), e aponta as diferenças em um Excel.

## Estrutura

```
conferencia-frequencia/
  venv/                     # ambiente virtual (não versionar)
  input/
    secretaria/AAAA-MM.pdf  # PDF que a secretaria manda
    sistema/AAAA-MM.pdf     # PDF que o sistema gera
  output/
    conferencia_AAAA-MM.xlsx
  src/
    extract_secretaria.py   # lê o PDF de texto corrido da secretaria
    extract_sistema.py      # lê o PDF em tabela do sistema
    compare.py               # agrega por (matrícula, tipo de ocorrência) e compara
    main.py                  # roda tudo e gera o Excel
  requirements.txt
```

## Uso mensal

1. Salve os dois PDFs do mês em:
   - `input/secretaria/AAAA-MM.pdf`
   - `input/sistema/AAAA-MM.pdf`

2. Rode:
   ```bash
   ./venv/bin/python src/main.py input/secretaria/AAAA-MM.pdf input/sistema/AAAA-MM.pdf
   ```
   O script descobre o mês de referência sozinho (pela data mais comum no
   relatório da secretaria). Se quiser forçar, use `--mes AAAA-MM`:
   ```bash
   ./venv/bin/python src/main.py input/secretaria/2026-04.pdf input/sistema/2026-04.pdf --mes 2026-04
   ```

3. Confira a saída em `output/`:
   - `conferencia_AAAA-MM.xlsx` — sempre gerado, aba **Divergências** (colorida)
     + aba **Resumo**.
   - `divergencias_AAAA-MM.md` — **só é gerado se houver divergência**. Serve
     pra colar num e-mail, chamado, ou revisão rápida sem precisar abrir Excel.

(Se preferir, dá pra criar um `.bat`/alias/atalho chamando esses dois
comandos, ou eu adapto para vasculhar `input/` inteiro e gerar todos os
meses pendentes de uma vez — é só pedir.)

## O que o script faz

- **Extrai** os registros dos dois PDFs (matrícula, nome, tipo de
  ocorrência, datas, quantidade de dias).
- **Ignora** "Frequência normal" (não é uma ocorrência, é a ausência de
  qualquer evento no mês — só aparece no relatório da secretaria).
- **Recorta ao mês de referência**: para os registros do sistema, que trazem
  data inicial e final, só entra na soma a quantidade de dias que cai
  *dentro* do mês (ex.: um período de 16/03 a 14/04 conta só os 14 dias de
  abril). Isso evita apontar divergência só porque um evento também ocupa
  dias de outro mês.
- **Agrupa** por matrícula + tipo de ocorrência (tratando "Abono" =
  "abono", "Tratamento de saúde" = "Tratamento De Saúde" etc. — a
  comparação ignora maiúscula/minúscula e acentuação) e soma os dias.
- **Aponta 3 tipos de divergência**:
  - *Só na secretaria*: ocorrência que a secretaria lançou e o sistema não tem.
  - *Só no sistema*: ocorrência que o sistema tem e a secretaria não lançou.
  - *Quantidade de dias divergente*: ambos têm o mesmo tipo de ocorrência
    para a mesma pessoa, mas a soma de dias (já recortada ao mês) não bate.

## Um padrão que ainda vai aparecer bastante

Mesmo recortando ao mês, sobra uma categoria de divergência sistemática:
quando o evento **começa dentro do mês mas continua depois** (ex.: férias
que começam em 15/04 e só terminam em 14/05), a secretaria parece registrar
o **total do evento inteiro** (30 dias), não só a parte de abril — enquanto
o sistema, já recortado por este script, mostra só os dias dentro do mês
(16). Isso é diferente do caso "começou antes do mês", onde a secretaria já
parece mostrar só a parte de abril. Ou seja: a secretaria parece contar o
evento inteiro no mês em que ele *começa*, e a parte restante nos meses
seguintes só quando o evento também começou antes deles. Vale confirmar essa
regra com quem gera o relatório da secretaria — mas dá pra reconhecer esse
padrão pelo nome da ocorrência (Férias regulamentares, Férias prêmio,
Afastamento sem vencimentos, Auxílio doença, Cedido sem ônus para cedente
são os tipos que mais aparecem aqui) e pela quantidade da secretaria ser bem
maior que a do sistema recortado.

## Ajustando o "de-para" de tipos de ocorrência

Se aparecer um tipo de ocorrência escrito de formas diferentes nos dois
relatórios (tipo "Falta" na secretaria virando "Faltas Efetivos" no
sistema, que eu já mapeei), adicione o par em `ALIASES`, no topo de
`src/compare.py`:

```python
ALIASES = {
    "falta": "faltas efetivos",
    # "outro tipo escrito diferente": "tipo equivalente no sistema",
}
```
(As chaves/valores devem estar em minúsculas e sem acento — é assim que
o código normaliza antes de comparar.)

## Limitações conhecidas

- O PDF da secretaria vem em texto corrido (não é uma tabela de verdade),
  e o rodapé/marca d'água de algumas páginas vaza como "linhas não
  reconhecidas" nos logs — isso é normal e não afeta os dados extraídos
  (já são descartadas). Se quiser conferir, rode
  `./venv/bin/python src/extract_secretaria.py input/secretaria/AAAA-MM.pdf`
  direto e olhe a lista impressa.
- A comparação depende de a matrícula ser idêntica nos dois relatórios
  (ela é, nos dois modelos vistos até agora — `NN.NNN-N`).
- O recorte ao mês só é aplicado ao lado do **sistema** (que tem data
  inicial e final). O lado da **secretaria** só tem uma data e uma
  quantidade — é usado como está, sem recorte (ver seção acima sobre o
  padrão de eventos que começam dentro do mês).
