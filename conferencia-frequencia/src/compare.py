# -*- coding: utf-8 -*-
"""
Compara os registros extraídos da SECRETARIA com os do SISTEMA dia a dia,
dentro do mês de referência, e monta a lista de "retificações": para
cada matrícula e cada dia em que o tipo registrado pela secretaria não
bate com o tipo registrado pelo sistema, gera uma linha no formato

    matricula - nome - data(s) - X dia(s) - tipo_secretaria --> tipo_sistema

(dias consecutivos com o mesmo par "de/para" são agrupados num intervalo,
igual ao que se faz na conferência manual).

Comparar dia a dia — em vez de somar quantidades totais por tipo — já
resolve sozinho o problema de eventos que também ocupam dias de outro
mês: nos dias que efetivamente caem dentro do mês de referência, o tipo
bate normalmente, mesmo que a duração total do evento (fora do mês) não
bata com o que está na secretaria.
"""
import unicodedata
from datetime import date, timedelta
from collections import Counter

# Tipos que representam "nada a registrar" e não entram na comparação
TIPOS_IGNORADOS = {"frequencia normal"}

# Aliases: grafias diferentes para o mesmo tipo de ocorrência
# (chave e valor já devem estar normalizados por _normalizar)
ALIASES = {
    "falta": "faltas efetivos",
}

ROTULO_SEM_REGISTRO_SECRETARIA = "SEM REGISTRO"
ROTULO_SEM_REGISTRO_SISTEMA = "sem registro em sistema"


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
    ultimo_excl = (date(ano, mes + 1, 1) if mes < 12 else date(ano + 1, 1, 1))
    ultimo_incl = ultimo_excl - timedelta(days=1)
    return primeiro, ultimo_incl


def _expandir_dias(registros, campo_tipo, campo_data_ini, campo_data_fim,
                    campo_qtde, primeiro_dia, ultimo_dia_incl):
    """Expande cada ocorrência em entradas por dia, recortadas ao mês.
    Retorna (mapa, nomes):
      mapa: {(matricula, date): rotulo_original_da_ocorrencia}
      nomes: {matricula: nome}
    """
    mapa = {}
    nomes = {}
    for r in registros:
        matricula = r["matricula"]
        nomes.setdefault(matricula, r["nome"])

        tipo_raw = r[campo_tipo]
        tipo_canon = _tipo_canonico(tipo_raw)
        if tipo_canon in TIPOS_IGNORADOS:
            continue

        d_ini = _parse_data(r.get(campo_data_ini))
        if d_ini is None:
            continue

        if campo_data_fim:
            d_fim = _parse_data(r.get(campo_data_fim)) or d_ini
        else:
            # a secretaria só informa a data de início + quantidade de
            # dias corridos a partir dela
            qtd = r.get(campo_qtde) or 1
            try:
                qtd_i = max(int(round(qtd)), 1)
            except (TypeError, ValueError):
                qtd_i = 1
            d_fim = d_ini + timedelta(days=qtd_i - 1)

        inicio = max(d_ini, primeiro_dia)
        fim = min(d_fim, ultimo_dia_incl)
        d = inicio
        while d <= fim:
            mapa[(matricula, d)] = tipo_raw
            d += timedelta(days=1)

    return mapa, nomes


def comparar_por_dia(regs_secretaria, regs_sistema, ano_mes=None):
    """Retorna (retificacoes, (ano_ref, mes_ref)).

    Cada item de `retificacoes`:
      matricula, nome, data_inicio, data_fim, dias,
      tipo_secretaria (rótulo ou "SEM REGISTRO"),
      tipo_sistema (rótulo ou "sem registro em sistema")
    """
    ano_ref, mes_ref = ano_mes or mes_referencia(regs_secretaria) or (None, None)
    if not ano_ref:
        return [], (None, None)

    primeiro_dia, ultimo_dia_incl = _limites_mes(ano_ref, mes_ref)

    mapa_sec, nomes_sec = _expandir_dias(
        regs_secretaria, "ocorrencia", "data", None, "qtde_dias",
        primeiro_dia, ultimo_dia_incl)
    mapa_sis, nomes_sis = _expandir_dias(
        regs_sistema, "descricao", "data_inicial", "data_final", "qtde_dias",
        primeiro_dia, ultimo_dia_incl)

    nomes = {**nomes_sis, **nomes_sec}  # nomes_sec prevalece (fonte com todo mundo)

    retificacoes = []
    for matricula in sorted(nomes):
        nome = nomes[matricula]
        pendente = None  # dict: par, inicio, fim

        def fechar():
            if pendente is None:
                return
            val_sec, val_sis = pendente["par"]
            dias = (pendente["fim"] - pendente["inicio"]).days + 1
            retificacoes.append({
                "matricula": matricula,
                "nome": nome,
                "data_inicio": pendente["inicio"],
                "data_fim": pendente["fim"],
                "dias": dias,
                "tipo_secretaria": val_sec or ROTULO_SEM_REGISTRO_SECRETARIA,
                "tipo_sistema": val_sis or ROTULO_SEM_REGISTRO_SISTEMA,
            })

        d = primeiro_dia
        while d <= ultimo_dia_incl:
            val_sec = mapa_sec.get((matricula, d))
            val_sis = mapa_sis.get((matricula, d))
            tipo_sec_canon = _tipo_canonico(val_sec) if val_sec else None
            tipo_sis_canon = _tipo_canonico(val_sis) if val_sis else None

            if tipo_sec_canon != tipo_sis_canon:
                par = (val_sec, val_sis)
                if pendente and pendente["par"] == par:
                    pendente["fim"] = d
                else:
                    fechar()
                    pendente = {"par": par, "inicio": d, "fim": d}
            else:
                fechar()
                pendente = None
            d += timedelta(days=1)
        fechar()

    retificacoes.sort(key=lambda x: (x["matricula"], x["data_inicio"]))
    return retificacoes, (ano_ref, mes_ref)
