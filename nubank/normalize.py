"""Estagio 2: RawFatura -> Fatura. Ainda sem categoria.

Tres extracoes estruturais acontecem aqui, e nenhuma delas e cosmetica:

  'Parcela 2/6'          -> parcela_atual/parcela_total + merchant limpo
  'IOF de "Render.Com"'  -> linha filha, ligada ao merchant pai
  bloco de cambio        -> moeda de origem, valor de origem, taxa

Sem a primeira, todo parcelamento vira um merchant diferente por mes. Sem a
segunda, o IOF fica boiando sem dono e o custo real do merchant fica errado
para baixo.

A inferencia de ano tambem mora aqui: o PDF escreve '03 JUL' sem ano, e na
fatura de janeiro isso significa dezembro do ano anterior.
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date
from decimal import Decimal

from .errors import ErroExtracao
from .extract import MESES
from .models import Fatura, LinhaRaw, RawFatura, TipoLancamento, Transacao
from .money import parse_brl

_RE_CARTAO = re.compile(r"^•+\s*(\d{4})\s*")
_RE_PARCELA = re.compile(r"\s*-\s*Parcela\s+(\d+)/(\d+)\s*$")
_RE_IOF = re.compile(r'^IOF de\s+"?(.+?)"?\s*$')
_RE_ESTORNO = re.compile(r'^Estorno de\s+"?(.+?)"?\s*$')
_RE_PAGAMENTO = re.compile(r"^Pagamento em\s+\d{2}\s+\w{3}\s*$")
_RE_AJUSTE = re.compile(r"^Saldo restante da fatura anterior\s*$", re.IGNORECASE)

_RE_MOEDA = re.compile(r"^([A-Z]{3})\s+([\d.]+)")
_RE_TAXA = re.compile(r"=\s*R\$\s*([\d.]+,\d{2})\s*$")


def _data_com_mes_abrev(texto: str, ano: int) -> date:
    """'03 JUL' ou '03 JUL 2026' -> date, usando `ano` quando nao houver ano."""
    partes = texto.split()
    dia = int(partes[0])
    mes = MESES[partes[1]]
    if len(partes) >= 3:
        ano = int(partes[2])
    return date(ano, mes, dia)


def _resolve_periodo(periodo_txt: str, vencimento: date) -> tuple[date, date]:
    """'03 JUL a 03 AGO' + vencimento -> (inicio, fim) com anos corretos.

    O periodo sempre termina antes do vencimento e comeca antes do fim. Cada
    ponta recua um ano quando a data candidata cai depois da sua ancora, que e
    o que resolve a virada dez/jan sem nenhum caso especial.
    """
    ini_txt, fim_txt = [p.strip() for p in periodo_txt.split(" a ")]

    fim = _data_com_mes_abrev(fim_txt, vencimento.year)
    if fim > vencimento:
        fim = fim.replace(year=fim.year - 1)

    ini = _data_com_mes_abrev(ini_txt, fim.year)
    if ini > fim:
        ini = ini.replace(year=ini.year - 1)

    return ini, fim


def _resolve_ano_transacao(dia: int, mes: int, ini: date, fim: date) -> date:
    """Escolhe o ano que faz a transacao cair dentro do periodo vigente."""
    for ano in dict.fromkeys((ini.year, fim.year)):
        try:
            candidata = date(ano, mes, dia)
        except ValueError:  # 29 FEV em ano nao bissexto
            continue
        if ini <= candidata <= fim:
            return candidata

    # Fora do periodo (o Nubank ocasionalmente lista algo na borda). Fica com a
    # candidata mais proxima do periodo em vez de inventar um ano.
    candidatas = []
    for ano in dict.fromkeys((ini.year, fim.year)):
        try:
            candidatas.append(date(ano, mes, dia))
        except ValueError:
            continue
    if not candidatas:
        raise ErroExtracao(f"data invalida: dia {dia}, mes {mes}")
    return min(candidatas, key=lambda d: min(abs((d - ini).days), abs((d - fim).days)))


def normaliza_merchant(descricao: str) -> str:
    """Descricao crua -> chave estavel de merchant.

    Tira o sufixo de cartao, o sufixo de parcela, acentos e caixa, e colapsa o
    '*' dos agregadores ('Shopee *Utilshome' e 'shopee*utilshome' sao o mesmo
    padrao de merchant e precisam bater na mesma regra).
    """
    s = _RE_CARTAO.sub("", descricao)
    s = _RE_PARCELA.sub("", s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.lower().strip()
    s = re.sub(r"\s*\*\s*", "*", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip(" .-")


def _le_conversao(
    linhas: tuple[str, ...],
) -> tuple[str | None, Decimal | None, Decimal | None]:
    """Bloco de cambio -> (moeda, valor na origem, taxa em R$)."""
    moeda = valor = taxa = None
    for linha in linhas:
        if linha.startswith("Conversão") or linha.startswith("Conversao"):
            m = _RE_TAXA.search(linha)
            if m:
                taxa = parse_brl(m.group(1))
        elif moeda is None:
            m = _RE_MOEDA.match(linha)
            if m:
                moeda = m.group(1)
                valor = Decimal(m.group(2))
    return moeda, valor, taxa


def _transacao(linha: LinhaRaw, ini: date, fim: date) -> Transacao:
    descricao = linha.descricao
    cartao = None

    m_cartao = _RE_CARTAO.match(descricao)
    if m_cartao:
        cartao = m_cartao.group(1)
        descricao_sem_cartao = _RE_CARTAO.sub("", descricao)
    else:
        descricao_sem_cartao = descricao

    parcela_atual = parcela_total = None
    m_parcela = _RE_PARCELA.search(descricao_sem_cartao)
    if m_parcela:
        parcela_atual = int(m_parcela.group(1))
        parcela_total = int(m_parcela.group(2))

    merchant_pai = None
    if m := _RE_IOF.match(descricao_sem_cartao):
        tipo = TipoLancamento.IOF
        merchant_pai = normaliza_merchant(m.group(1))
    elif m := _RE_ESTORNO.match(descricao_sem_cartao):
        tipo = TipoLancamento.ESTORNO
        merchant_pai = normaliza_merchant(m.group(1))
    elif _RE_PAGAMENTO.match(descricao_sem_cartao):
        tipo = TipoLancamento.PAGAMENTO
    elif _RE_AJUSTE.match(descricao_sem_cartao):
        tipo = TipoLancamento.AJUSTE
    elif parcela_atual is not None:
        tipo = TipoLancamento.PARCELA
    else:
        tipo = TipoLancamento.COMPRA

    # Para IOF e estorno o merchant que importa e o pai: e nele que o valor
    # tem de aterrissar na hora de somar por merchant e por categoria.
    merchant = merchant_pai or normaliza_merchant(descricao_sem_cartao)
    moeda, valor_origem, taxa = _le_conversao(linha.conversao)

    return Transacao(
        data=_resolve_ano_transacao(linha.dia, MESES[linha.mes_abrev], ini, fim),
        descricao_raw=descricao,
        merchant_norm=merchant,
        valor=linha.valor,
        tipo=tipo,
        cartao_final=cartao,
        parcela_atual=parcela_atual,
        parcela_total=parcela_total,
        merchant_pai=merchant_pai,
        moeda_origem=moeda,
        valor_origem=valor_origem,
        taxa_cambio=taxa,
    )


def normaliza(raw: RawFatura) -> Fatura:
    """RawFatura -> Fatura com tipos de dominio e datas resolvidas."""
    vencimento = _data_com_mes_abrev(raw.vencimento_txt, ano=0)
    ini, fim = _resolve_periodo(raw.periodo_txt, vencimento)

    return Fatura(
        competencia=f"{vencimento.year:04d}-{vencimento.month:02d}",
        vencimento=vencimento,
        periodo_ini=ini,
        periodo_fim=fim,
        fatura_anterior=raw.resumo["fatura_anterior"],
        pagamentos=raw.resumo["pagamentos"],
        compras=raw.resumo["compras"],
        iof=raw.resumo["iof"],
        outros_lancamentos=raw.resumo["outros_lancamentos"],
        total_a_pagar=raw.resumo["total_a_pagar"],
        saldo_aberto_proximo=raw.saldo_aberto_proximo,
        saldo_aberto_total=raw.saldo_aberto_total,
        pdf_sha256=raw.pdf_sha256,
        arquivo=raw.caminho,
        transacoes=[_transacao(l, ini, fim) for l in raw.linhas],
    )
