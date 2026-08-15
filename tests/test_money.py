from decimal import Decimal

import pytest

from nubank.money import encontra_valor, format_brl, parse_brl


@pytest.mark.parametrize(
    "texto, esperado",
    [
        ("1.234,56", "1234.56"),
        ("0,00", "0.00"),
        ("8,37", "8.37"),
        ("10.250,00", "10250.00"),
        ("R$ 3.276,42", "3276.42"),
        ("-R$ 2.788,70", "-2788.70"),
        ("−R$ 22,95", "-22.95"),  # MINUS SIGN, o que o Nubank usa no estorno
    ],
)
def test_parse_brl(texto, esperado):
    assert parse_brl(texto) == Decimal(esperado)


def test_parse_brl_nao_usa_float():
    """0,1 + 0,2 em float da 0.30000000000000004 e quebraria a reconciliacao."""
    assert parse_brl("0,10") + parse_brl("0,20") == parse_brl("0,30")


@pytest.mark.parametrize(
    "linha, esperado",
    [
        ("Total a pagar R$ 3.276,42", "3276.42"),
        ("Pagamento recebido −R$ 2.788,70", "-2788.70"),
        ("06 JUL Estorno de Uber - NuPay −R$ 22,95", "-22.95"),
        ("linha sem valor nenhum", None),
    ],
)
def test_encontra_valor(linha, esperado):
    resultado = encontra_valor(linha)
    assert resultado == (Decimal(esperado) if esperado else None)


def test_encontra_valor_pega_o_ultimo():
    """A tabela de parcelamento tem dois valores na mesma linha."""
    assert encontra_valor("Total a pagar R$ 3.532,40 R$ 3.726,31") == Decimal("3726.31")


@pytest.mark.parametrize(
    "valor, esperado",
    [("1234.56", "R$ 1.234,56"), ("-50.00", "-R$ 50,00"), ("0.00", "R$ 0,00")],
)
def test_format_brl(valor, esperado):
    assert format_brl(Decimal(valor)) == esperado
