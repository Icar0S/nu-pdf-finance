from dataclasses import replace
from decimal import Decimal

import pytest

from nubank.errors import ErroReconciliacao
from nubank.normalize import normaliza
from nubank.reconcile import TOLERANCIA, invariantes, reconcilia


def test_fatura_integra_passa(raw):
    resultado = reconcilia(normaliza(raw))
    assert len(resultado) == 3
    assert all(i.ok for i in resultado)


def test_centavo_de_arredondamento_passa(raw):
    """O proprio Nubank erra 1 centavo em metade das faturas do corpus."""
    resumo = dict(raw.resumo)
    resumo["compras"] += Decimal("0.01")
    resumo["total_a_pagar"] += Decimal("0.01")
    fatura = normaliza(replace(raw, resumo=resumo))
    assert all(i.ok for i in invariantes(fatura))


def test_transacao_perdida_e_rejeitada(raw):
    """O cenario que o portao existe para pegar: o parser comeu uma linha."""
    fatura = normaliza(raw)
    fatura.transacoes.pop(1)
    with pytest.raises(ErroReconciliacao, match="nao fecha"):
        reconcilia(fatura)


def test_valor_errado_e_rejeitado(raw):
    fatura = normaliza(raw)
    fatura.transacoes[1].valor += Decimal("10.00")
    with pytest.raises(ErroReconciliacao):
        reconcilia(fatura)


def test_pagamento_perdido_e_rejeitado(raw):
    fatura = normaliza(raw)
    fatura.transacoes = [t for t in fatura.transacoes if "Pagamento" not in t.descricao_raw]
    with pytest.raises(ErroReconciliacao):
        reconcilia(fatura)


def test_erro_maior_que_a_tolerancia_e_rejeitado(raw):
    fatura = normaliza(raw)
    fatura.transacoes[1].valor += TOLERANCIA + Decimal("0.01")
    with pytest.raises(ErroReconciliacao):
        reconcilia(fatura)


def test_ajuste_de_saldo_e_a_diferenca_entre_total_e_gasto(raw):
    """Total a pagar e gasto do periodo sao grandezas diferentes.

    A diferenca e sempre fatura_anterior - pagamentos, e e o residuo que sobra
    na coluna F da planilha.
    """
    fatura = normaliza(raw)
    gasto = fatura.compras + fatura.iof + fatura.outros_lancamentos
    assert fatura.total_a_pagar - gasto == fatura.ajuste_saldo
