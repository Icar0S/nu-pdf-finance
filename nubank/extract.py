"""Estagio 1: PDF -> RawFatura. Sem interpretacao, sem categoria.

O PDF do Nubank tem camada de texto real, entao nao ha OCR envolvido. O
extract_text() do pdfplumber ja devolve as linhas na ordem certa e com as
colunas colapsadas de forma estavel: 'DD MMM <descricao> R$ <valor>'.

A unica armadilha real e o rotulo 'Total a pagar', que aparece tres vezes no
documento: duas na tabela de simulacao de parcelamento (com dois valores na
mesma linha) e uma no resumo. Por isso o resumo e lido por secao ancorada, e
nao por varredura do documento inteiro.
"""

from __future__ import annotations

import hashlib
import re
from decimal import Decimal
from pathlib import Path

import pdfplumber

from .errors import ErroExtracao
from .models import LinhaRaw, RawFatura
from .money import parse_brl

MESES = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12, "JUL": 7, "AGO": 8,
}
_MES_RE = "|".join(MESES)

_RE_TRANSACAO = re.compile(
    r"^(?P<dia>\d{2}) (?P<mes>" + _MES_RE + r") "
    r"(?P<desc>.+?) "
    r"(?P<sinal>[-−–—])?R\$\s?(?P<valor>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})$"
)

# Blocos de cambio que seguem uma transacao internacional. Duas formas:
#   USD 7.00 / Conversao: USD 1 = R$ 5,37
#   BRL 110.00 = USD 21.57 / Conversao: BRL 5.29 = USD 1 = R$ 5,29
_RE_CONVERSAO = re.compile(r"^(?:[A-Z]{3} [\d.,]+|Convers[ãa]o:.*)$")

_RE_VENCIMENTO = re.compile(
    r"Data de vencimento:\s*(\d{2}\s+(?:" + _MES_RE + r")\s+\d{4})"
)
_RE_PERIODO = re.compile(
    r"Per[íi]odo vigente:\s*(\d{2}\s+(?:" + _MES_RE + r"))\s+a\s+(\d{2}\s+(?:" + _MES_RE + r"))"
)
_RE_FECHAMENTO = re.compile(
    r"Fechamento da pr[óo]xima fatura\s+(\d{2}\s+(?:" + _MES_RE + r")\s+\d{4})"
)

MARCADOR_TRANSACOES = "TRANSAÇÕES DE"
MARCADOR_RESUMO = "RESUMO DA FATURA ATUAL"
MARCADOR_PROXIMAS = "PRÓXIMAS FATURAS"

# Rotulo no PDF -> chave no dict de resumo. A ordem importa: 'Total de compras'
# precisa ser testado antes de qualquer prefixo mais curto.
CAMPOS_RESUMO = (
    ("fatura_anterior", r"Fatura anterior"),
    ("pagamentos", r"Pagamento recebido"),
    ("compras", r"Total de compras de todos os cart[õo]es"),
    ("iof", r"IOF de compras internacionais"),
    ("outros_lancamentos", r"Outros lan[çc]amentos"),
    ("total_a_pagar", r"Total a pagar"),
)


def sha256_arquivo(caminho: Path) -> str:
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def _paginas_texto(caminho: Path) -> list[list[str]]:
    with pdfplumber.open(caminho) as pdf:
        return [
            [linha.strip() for linha in (p.extract_text() or "").split("\n")]
            for p in pdf.pages
        ]


def _le_resumo(linhas: list[str], arquivo: str) -> dict[str, Decimal]:
    """Le o bloco RESUMO DA FATURA ATUAL, ancorado pelo cabecalho da secao."""
    try:
        inicio = linhas.index(MARCADOR_RESUMO)
    except ValueError:
        raise ErroExtracao(
            f"{arquivo}: secao '{MARCADOR_RESUMO}' nao encontrada. "
            "Isso nao parece uma fatura de cartao do Nubank."
        ) from None

    fim = len(linhas)
    for i in range(inicio, len(linhas)):
        if linhas[i] == MARCADOR_PROXIMAS:
            fim = i
            break

    resumo: dict[str, Decimal] = {}
    for chave, rotulo in CAMPOS_RESUMO:
        padrao = re.compile(
            r"^" + rotulo + r".*?\s([-−–—]?)R\$\s?(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})$"
        )
        for linha in linhas[inicio:fim]:
            m = padrao.match(linha)
            if m:
                valor = parse_brl(m.group(2))
                resumo[chave] = -valor if m.group(1) else valor
                break

    faltando = [c for c, _ in CAMPOS_RESUMO if c not in resumo]
    if faltando:
        raise ErroExtracao(
            f"{arquivo}: campos do resumo nao encontrados: {', '.join(faltando)}"
        )

    # 'Pagamento recebido' vem negativo no PDF. Guardamos positivo: e o quanto
    # foi pago, e o sinal fica explicito na formula de reconciliacao.
    resumo["pagamentos"] = abs(resumo["pagamentos"])
    return resumo


def _le_transacoes(paginas: list[list[str]]) -> tuple[LinhaRaw, ...]:
    """Coleta as linhas datadas depois do primeiro marcador TRANSACOES DE.

    Tudo que nao e uma linha datada e ignorado: cabecalho de pagina, totais de
    secao ('Icaro S Oliveira R$ x', 'Pagamentos -R$ x'), rodape legal. As
    linhas de cambio sao anexadas a transacao imediatamente anterior.
    """
    linhas_raw: list[LinhaRaw] = []
    em_transacoes = False
    pendente_conversao: list[str] = []

    for n_pagina, linhas in enumerate(paginas):
        for linha in linhas:
            if linha.startswith(MARCADOR_TRANSACOES):
                em_transacoes = True
                continue
            if not em_transacoes or not linha:
                continue

            m = _RE_TRANSACAO.match(linha)
            if m:
                if linhas_raw and pendente_conversao:
                    ultimo = linhas_raw[-1]
                    linhas_raw[-1] = LinhaRaw(
                        dia=ultimo.dia,
                        mes_abrev=ultimo.mes_abrev,
                        descricao=ultimo.descricao,
                        valor=ultimo.valor,
                        pagina=ultimo.pagina,
                        conversao=tuple(pendente_conversao),
                    )
                pendente_conversao = []

                valor = parse_brl(m.group("valor"))
                if m.group("sinal"):
                    valor = -valor
                linhas_raw.append(
                    LinhaRaw(
                        dia=int(m.group("dia")),
                        mes_abrev=m.group("mes"),
                        descricao=m.group("desc").strip(),
                        valor=valor,
                        pagina=n_pagina,
                    )
                )
            elif _RE_CONVERSAO.match(linha):
                pendente_conversao.append(linha)

    if linhas_raw and pendente_conversao:
        ultimo = linhas_raw[-1]
        linhas_raw[-1] = LinhaRaw(
            dia=ultimo.dia,
            mes_abrev=ultimo.mes_abrev,
            descricao=ultimo.descricao,
            valor=ultimo.valor,
            pagina=ultimo.pagina,
            conversao=tuple(pendente_conversao),
        )

    return tuple(linhas_raw)


def extrai(caminho: str | Path) -> RawFatura:
    """PDF -> RawFatura. Levanta ErroExtracao se o layout nao for reconhecido."""
    caminho = Path(caminho)
    arquivo = caminho.name
    paginas = _paginas_texto(caminho)
    linhas = [linha for pagina in paginas for linha in pagina]

    m_venc = next(filter(None, (_RE_VENCIMENTO.search(l) for l in linhas)), None)
    if not m_venc:
        raise ErroExtracao(f"{arquivo}: 'Data de vencimento' nao encontrada.")

    m_per = next(filter(None, (_RE_PERIODO.search(l) for l in linhas)), None)
    if not m_per:
        raise ErroExtracao(f"{arquivo}: 'Período vigente' nao encontrado.")

    resumo = _le_resumo(linhas, arquivo)
    linhas_raw = _le_transacoes(paginas)
    if not linhas_raw:
        raise ErroExtracao(f"{arquivo}: nenhuma transacao encontrada.")

    m_fech = next(filter(None, (_RE_FECHAMENTO.search(l) for l in linhas)), None)

    def _saldo(rotulo: str) -> Decimal | None:
        padrao = re.compile(
            r"^" + rotulo + r"\s+([-−–—]?)R\$\s?(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})$"
        )
        for linha in linhas:
            m = padrao.match(linha)
            if m:
                valor = parse_brl(m.group(2))
                return -valor if m.group(1) else valor
        return None

    return RawFatura(
        caminho=str(caminho),
        pdf_sha256=sha256_arquivo(caminho),
        vencimento_txt=m_venc.group(1),
        periodo_txt=f"{m_per.group(1)} a {m_per.group(2)}",
        resumo=resumo,
        fechamento_proximo_txt=m_fech.group(1) if m_fech else None,
        saldo_aberto_proximo=_saldo(r"Saldo em aberto da pr[óo]xima fatura"),
        saldo_aberto_total=_saldo(r"Saldo em aberto total"),
        linhas=linhas_raw,
    )
