# -*- coding: utf-8 -*-
"""
Compara os registros extraídos da SECRETARIA com os do SISTEMA,
agregando por (matrícula, tipo de ocorrência) e apontando divergências:
  - ocorrência só na secretaria
  - ocorrência só no sistema
  - mesma ocorrência, mas quantidade de dias diferente

A comparação considera SOMENTE os dias que caem dentro do mês de
referência: se um período do sistema começa antes ou termina depois do
mês (ex.: férias que começaram em março), só a fatia de dias dentro do
mês entra na soma. Isso evita apontar como "divergência" algo que na
verdade é só o evento continuar fora do mês.
"""
import unicodedata
from datetime import date
from collections import Counter

# Tipos que representam "nada a registrar" e não entram na comparação
TIPOS_IGNORADOS = {"frequencia normal"}

# Aliases: grafias diferentes para o mesmo tipo de ocorrência
# (chave e valor já devem estar normalizados por _normalizar)
ALIASES = {
    "falta": "faltas efetivos",
}


def _normalizar(texto):
    """minúsculas, sem acento, espaços colapsados — para comparar tipos
    escritos com capitalização/formatação diferente nos dois relatórios."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(sem_acento.lower().split())


def _tipo_canonico(texto):
    norm = _normalizar(texto)
    return ALIASES.get(norm, norm)


def _parse_data(txt):
    if not txt:
        return None
    try:
        d, m, a = txt.split("/")
        return date(int(a), int(m), int(d))
    except (ValueError, AttributeError):
        return None


def mes_referencia(regs_secretaria):
    """Descobre o mês/ano de referência do relatório a partir da moda das
    datas dos registros da secretaria. Retorna (ano, mes) ou None."""
    contagem = Counter()
    for r in regs_secretaria:
        d = _parse_data(r.get("data"))
        if d:
            contagem[(d.year, d.month)] += 1
    if not contagem:
        return None
    return contagem.most_common(1)[0][0]


def _limites_mes(ano, mes):
    primeiro = date(ano, mes, 1)
    ultimo = (date(ano, mes + 1, 1) if mes < 12 else date(ano + 1, 1, 1))
    return primeiro, ultimo  # ultimo é exclusivo (1º dia do mês seguinte)


def _dias_dentro_do_mes(d_ini, d_fim, primeiro_dia, ultimo_dia_excl):
    """Quantos dias do intervalo [d_ini, d_fim] (inclusive) caem dentro do
    mês de referência. Se não houver sobreposição, retorna 0."""
    if d_ini is None:
        return None  # sem data para recortar, mantém o valor original
    if d_fim is None:
        d_fim = d_ini
    inicio = max(d_ini, primeiro_dia)
    fim = min(d_fim, date.fromordinal(ultimo_dia_excl.toordinal() - 1))
    delta = (fim - inicio).days + 1
    return max(delta, 0)


def agregar(registros, campo_tipo, campo_qtde, campo_data_ini=None,
            campo_data_fim=None, limites_mes=None):
    """Soma dias por (matricula, tipo_canonico). Retorna dict:
    {(matricula, tipo_canonico): {"dias", "nome", "rotulo"}}

    Se `limites_mes` (primeiro_dia, ultimo_dia_excl) for informado e o
    registro tiver datas, a quantidade somada é recortada para os dias
    que caem dentro do mês — não o total do evento inteiro.
    """
    agregados = {}
    for r in registros:
        tipo_canon = _tipo_canonico(r[campo_tipo])
        if tipo_canon in TIPOS_IGNORADOS:
            continue

        qtd = r[campo_qtde] or 0.0
        if limites_mes and campo_data_ini:
            d_ini = _parse_data(r.get(campo_data_ini))
            d_fim = _parse_data(r.get(campo_data_fim)) if campo_data_fim else d_ini
            dias_no_mes = _dias_dentro_do_mes(d_ini, d_fim, *limites_mes)
            if dias_no_mes is not None:
                qtd = dias_no_mes

        chave = (r["matricula"], tipo_canon)
        if chave not in agregados:
            agregados[chave] = {"dias": 0.0, "nome": r["nome"], "rotulo": r[campo_tipo]}
        item = agregados[chave]
        item["dias"] += qtd
        if len(r[campo_tipo] or "") > len(item["rotulo"] or ""):
            item["rotulo"] = r[campo_tipo]
    return agregados


def comparar(regs_secretaria, regs_sistema, tolerancia_dias=0.0, ano_mes=None):
    """Retorna (divergencias, (ano_ref, mes_ref)).

    ano_mes: opcional, tupla (ano, mes) para forçar o mês de referência
    em vez de descobrir pela moda das datas da secretaria.
    """
    ano_ref, mes_ref = ano_mes or mes_referencia(regs_secretaria) or (None, None)
    limites = _limites_mes(ano_ref, mes_ref) if ano_ref else None

    agr_sec = agregar(regs_secretaria, "ocorrencia", "qtde_dias", "data")
    agr_sis = agregar(regs_sistema, "descricao", "qtde_dias", "data_inicial",
                       "data_final", limites_mes=limites)

    todas_chaves = set(agr_sec) | set(agr_sis)
    divergencias = []

    for chave in sorted(todas_chaves, key=lambda c: (c[0], c[1])):
        matricula, tipo = chave
        sec = agr_sec.get(chave)
        sis = agr_sis.get(chave)

        dias_sec = sec["dias"] if sec else None
        dias_sis = sis["dias"] if sis else None
        nome = (sec or sis)["nome"]
        rotulo = (sec or sis)["rotulo"]

        if sec and not sis:
            situacao = "Só na secretaria (sistema não tem)"
        elif sis and not sec:
            situacao = "Só no sistema (secretaria não tem)"
        elif abs((dias_sec or 0) - (dias_sis or 0)) > tolerancia_dias:
            situacao = "Quantidade de dias divergente"
        else:
            continue  # bate certinho, não é divergência

        divergencias.append({
            "matricula": matricula,
            "nome": nome,
            "tipo_ocorrencia": rotulo,
            "dias_secretaria": dias_sec,
            "dias_sistema": dias_sis,
            "diferenca": (dias_sec or 0) - (dias_sis or 0),
            "situacao": situacao,
        })

    return divergencias, (ano_ref, mes_ref)
