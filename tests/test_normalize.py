from datetime import date
from decimal import Decimal

import pytest

from nubank.models import TipoLancamento
from nubank.normalize import normaliza, normaliza_merchant


@pytest.mark.parametrize(
    "descricao, esperado",
    [
        ("•••• 5449 Alda Lanches", "alda lanches"),
        ("•••• 5449 Bombahia - Parcela 2/6", "bombahia"),
        ("Shopee *Utilshome", "shopee*utilshome"),
        ("Shopee*Utilshome", "shopee*utilshome"),
        ("•••• 5449 Dl*99 Ride", "dl*99 ride"),
        ("•••• 5449 Dl *99 Ride", "dl*99 ride"),
        ("Conversão Café", "conversao cafe"),
        ("•••• 3018 Render.Com", "render.com"),
        ("Github, Inc.", "github, inc"),
    ],
)
def test_normaliza_merchant(descricao, esperado):
    assert normaliza_merchant(descricao) == esperado


def test_parcela_vira_campo_estruturado(raw):
    fatura = normaliza(raw)
    parcela = next(t for t in fatura.transacoes if t.parcela_total)
    assert (parcela.parcela_atual, parcela.parcela_total) == (2, 2)
    assert parcela.tipo is TipoLancamento.PARCELA
    # o merchant fica limpo: sem o sufixo, as 6 parcelas de uma compra colapsam
    # num merchant so em vez de virarem 6 merchants diferentes.
    assert parcela.merchant_norm == "ebn*playstation"


def test_iof_aponta_para_o_merchant_pai(raw):
    """Sem isso o IOF fica sem dono e o custo do merchant sai errado pra baixo."""
    fatura = normaliza(raw)
    iof = next(t for t in fatura.transacoes if t.tipo is TipoLancamento.IOF)
    assert iof.merchant_pai == "render.com"
    assert iof.merchant_norm == "render.com"

    render = [t for t in fatura.transacoes if t.merchant_norm == "render.com"]
    assert sum(t.valor for t in render) == Decimal("38.97")  # 37,65 + 1,32


def test_estorno_e_negativo_e_herda_o_pai(raw):
    fatura = normaliza(raw)
    estorno = next(t for t in fatura.transacoes if t.tipo is TipoLancamento.ESTORNO)
    assert estorno.valor < 0
    assert estorno.merchant_norm == "uber - nupay"


def test_cambio_com_uma_moeda(raw):
    fatura = normaliza(raw)
    render = next(t for t in fatura.transacoes if t.merchant_norm == "render.com")
    assert render.moeda_origem == "USD"
    assert render.valor_origem == Decimal("7.00")
    assert render.taxa_cambio == Decimal("5.37")


def test_cambio_com_duas_moedas(raw):
    """'BRL 110.00 = USD 21.57' e o formato de assinatura cobrada em real."""
    fatura = normaliza(raw)
    tx = next(t for t in fatura.transacoes if t.merchant_norm == "anthropic*claude sub")
    assert tx.moeda_origem == "BRL"
    assert tx.valor_origem == Decimal("110.00")
    assert tx.taxa_cambio == Decimal("5.29")


def test_pagamento_nao_e_gasto(raw):
    fatura = normaliza(raw)
    pagamentos = [t for t in fatura.transacoes if t.tipo is TipoLancamento.PAGAMENTO]
    assert len(pagamentos) == 2
    assert all(not p.eh_gasto for p in pagamentos)


def test_ano_inferido_dentro_do_periodo(raw):
    fatura = normaliza(raw)
    assert fatura.competencia == "2026-08"
    assert (fatura.periodo_ini, fatura.periodo_fim) == (
        date(2026, 7, 3),
        date(2026, 8, 3),
    )
    assert all(fatura.periodo_ini <= t.data <= fatura.periodo_fim for t in fatura.transacoes)


def test_virada_de_ano(raw_virada_ano):
    """Fatura de janeiro/2026: '29 NOV' e novembro de 2025, nao de 2026."""
    fatura = normaliza(raw_virada_ano)
    assert fatura.competencia == "2026-01"
    assert fatura.periodo_ini == date(2025, 11, 29)
    assert fatura.periodo_fim == date(2025, 12, 30)
    assert all(t.data.year == 2025 for t in fatura.transacoes)


def test_competencia_vem_do_vencimento(raw):
    """A aba Cartao e indexada por mes de vencimento, nao por periodo vigente."""
    fatura = normaliza(raw)
    assert fatura.competencia == "2026-08"
    assert fatura.periodo_ini.month == 7
