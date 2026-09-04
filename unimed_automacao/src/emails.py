"""
emails.py - Geração dos modelos de e-mails em Markdown com tabelas de débitos
"""
from datetime import datetime
from pathlib import Path
import pandas as pd

from .config import (
    ARQUIVO_BASE,
    ARQUIVO_DEVEDORES,
    EMAILS_TEMPLATES_DIR,
    EMAILS_DIR,
)
from .utils import (
    limpa,
    formatar_valor_br,
    formatar_moeda,
    ultimo_dia_util_do_mes,
    mes_referencia_texto,
    formata_competencia,
)


def _gerar_tabela_markdown(tabela_itens):
    """Gera tabela visual em Markdown para exibição dos débitos."""
    if not tabela_itens:
        return "Sem débitos."

    linhas = [
        "| Competência | Vencimento | Principal | Encargos | Total |",
        "|-------------|------------|-----------|----------|-------|",
    ]
    for item in tabela_itens:
        linhas.append(
            f"| {item['competencia']} | {item['vencimento']} | {item['principal']} | {item['encargos']} | {item['total']} |"
        )
    return "\n".join(linhas)


def _salvar_arquivo_md(caminho_arquivo, titulo, lista_emails):
    """Salva a lista de e-mails formatados em um arquivo Markdown estruturado."""
    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        f.write(f"# {titulo}\n\n")
        f.write(f"> Total de destinatários: **{len(lista_emails)}**\n\n")

        for email in lista_emails:
            f.write("---\n\n")
            f.write(f"## 👤 {email['nome']}\n\n")
            f.write(f"**Para:** `{email['email']}`  \n")
            f.write(f"**Assunto:** `{email['assunto']}`  \n\n")
            f.write("### ✉️ Mensagem:\n\n")
            f.write(f"{email['mensagem']}\n\n")


def gerar_todos_emails():
    """Processa a base e os inadimplentes gerando os 4 arquivos de e-mail em saida/emails/."""
    print("\n" + "=" * 60)
    print(">>> INICIANDO GERACAO DE MODELOS DE E-MAILS")
    print("=" * 60)

    EMAILS_DIR.mkdir(parents=True, exist_ok=True)

    if not ARQUIVO_BASE.exists() or not ARQUIVO_DEVEDORES.exists():
        print("[ERRO] Arquivos de entrada ausentes (teste.ods ou devedores.xlsx).")
        return

    # Carregar templates de texto
    with open(EMAILS_TEMPLATES_DIR / "email_desligados.txt", "r", encoding="utf-8") as f:
        template_desligado = f.read()
    with open(EMAILS_TEMPLATES_DIR / "email_normal.txt", "r", encoding="utf-8") as f:
        template_normal = f.read()
    with open(EMAILS_TEMPLATES_DIR / "email_aviso.txt", "r", encoding="utf-8") as f:
        template_aviso = f.read()
    with open(EMAILS_TEMPLATES_DIR / "email_cancelado.txt", "r", encoding="utf-8") as f:
        template_cancelado = f.read()

    # Ler dados
    df_base = pd.read_excel(ARQUIVO_BASE, engine="odf")
    df_dividas = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Inadimplentes")
    df_cancelados = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Cancelados")

    df_base["Nro Funcional"] = pd.to_numeric(df_base["Nro Funcional"], errors="coerce").astype("Int64").astype(str)
    df_dividas["Funcional"] = pd.to_numeric(df_dividas["Funcional"], errors="coerce").astype("Int64").astype(str)
    df_cancelados["Funcional"] = pd.to_numeric(df_cancelados["Funcional"], errors="coerce").astype("Int64").astype(str)

    hoje = datetime.today()
    data_envio = hoje.strftime("%d/%m/%Y")
    data_vencimento = ultimo_dia_util_do_mes(hoje.year, hoje.month).strftime("%d/%m/%Y")
    referencia = mes_referencia_texto(hoje)

    emails_desligados = []
    emails_normais = []
    emails_aviso = []
    emails_cancelados = []

    for _, row in df_base.iterrows():
        condicao = limpa(row.get("condição")).lower()

        if "não enviar" in condicao or "nao enviar" in condicao:
            continue
        elif condicao == "desligado":
            tipo = "desligado"
        elif condicao == "aviso":
            tipo = "aviso"
        elif condicao == "cancelado":
            tipo = "cancelado"
        elif condicao in ("", "nan"):
            tipo = "normal"
        else:
            continue

        nome = limpa(row.get("Funcionário"))
        mail = limpa(row.get("mail"))
        matricula = limpa(row.get("Nro Funcional"))
        total_base = row.get("Total", "")

        tabela_md = ""
        valor_total_final = formatar_moeda(total_base) if isinstance(total_base, (int, float)) else str(total_base)

        if tipo in ["aviso", "cancelado"]:
            df_func = df_dividas[df_dividas["Funcional"] == matricula] if tipo == "aviso" else df_cancelados[df_cancelados["Funcional"] == matricula]

            total_dividas = pd.to_numeric(df_func["Saldo (Atualizado)"], errors="coerce").fillna(0).sum()
            valor_total_final = formatar_moeda(total_dividas)

            tabela_itens = []
            for _, d in df_func.iterrows():
                principal = float(d.get("Principal (Saldo)", 0) or 0)
                saldo_div = float(d.get("Saldo (Atualizado)", 0) or 0)
                encargos = saldo_div - principal

                data_venc = pd.to_datetime(d.get("Data de Vencimento"), dayfirst=True, errors="coerce")
                venc_str = data_venc.strftime("%d/%m/%Y") if pd.notna(data_venc) else ""

                tabela_itens.append({
                    "competencia": formata_competencia(d.get("Mês/Ano")),
                    "vencimento": venc_str,
                    "principal": formatar_moeda(principal),
                    "encargos": formatar_moeda(encargos),
                    "total": formatar_moeda(saldo_div),
                })
            tabela_md = _gerar_tabela_markdown(tabela_itens)

        if tipo == "desligado":
            template = template_desligado
        elif tipo == "aviso":
            template = template_aviso
        elif tipo == "cancelado":
            template = template_cancelado
        else:
            template = template_normal

        corpo = template.format(
            nome=nome,
            valor_total=valor_total_final,
            valores=valor_total_final,
            data_final_mes=data_vencimento,
            referencia=referencia,
            data=data_envio,
            tabela=tabela_md,
        )
        assunto = f"Boleto do Plano de Saúde Unimed – Referente a {referencia}"

        bloco = {
            "nome": nome,
            "email": mail,
            "assunto": assunto,
            "mensagem": corpo,
        }

        if tipo == "desligado":
            emails_desligados.append(bloco)
        elif tipo == "aviso":
            emails_aviso.append(bloco)
        elif tipo == "cancelado":
            emails_cancelados.append(bloco)
        else:
            emails_normais.append(bloco)

    # Salvar saídas
    _salvar_arquivo_md(EMAILS_DIR / "emails_normais.md", "📧 Emails Normais", emails_normais)
    _salvar_arquivo_md(EMAILS_DIR / "emails_desligados.md", "📧 Emails de Desligados", emails_desligados)
    _salvar_arquivo_md(EMAILS_DIR / "emails_aviso.md", "📧 Emails de Aviso de Cancelamento", emails_aviso)
    _salvar_arquivo_md(EMAILS_DIR / "emails_cancelados.md", "📧 Emails de Cancelados", emails_cancelados)

    print("\n[OK] Modelos de e-mail gerados em saida/emails/:")
    print(f"   * Normais: {len(emails_normais)}")
    print(f"   * Desligados: {len(emails_desligados)}")
    print(f"   * Avisos: {len(emails_aviso)}")
    print(f"   * Cancelados: {len(emails_cancelados)}")
