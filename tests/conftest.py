"""Fixtures compartilhadas.

Os PDFs reais ficam fora do git (faturas/ esta no .gitignore), entao o grosso
dos testes roda sobre uma RawFatura sintetica montada aqui. Os testes que
precisam dos PDFs de verdade pulam sozinhos quando a pasta nao existe - eles
sao o teste de regressao local, nao o de CI.
"""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest

from nubank.models import LinhaRaw, RawFatura

RAIZ = Path(__file__).resolve().parent.parent
PASTA_FATURAS = RAIZ / "faturas"


def pdfs_disponiveis() -> list[Path]:
    return sorted(PASTA_FATURAS.glob("*.pdf")) if PASTA_FATURAS.is_dir() else []


def d(valor: str) -> Decimal:
    return Decimal(valor)


@pytest.fixture
def linhas_exemplo() -> tuple[LinhaRaw, ...]:
    """Um recorte que cobre todos os formatos de linha que a fatura produz."""
    return (
        LinhaRaw(3, "JUL", "•••• 5449 Ebn *Playstation - Parcela 2/2", d("37.47"), 4),
        LinhaRaw(3, "JUL", "Uber - NuPay", d("8.83"), 4),
        LinhaRaw(
            5, "JUL", "•••• 3018 Render.Com", d("37.65"), 4,
            conversao=("USD 7.00", "Conversão: USD 1 = R$ 5,37"),
        ),
        LinhaRaw(5, "JUL", 'IOF de "Render.Com"', d("1.32"), 4),
        LinhaRaw(6, "JUL", "Estorno de Uber - NuPay", d("-22.95"), 4),
        LinhaRaw(6, "JUL", "•••• 5449 Ifd*Sorveteria 50 Sabo", d("70.35"), 4),
        LinhaRaw(
            21, "JUL", "•••• 5449 Anthropic* Claude Sub", d("114.17"), 6,
            conversao=("BRL 110.00 = USD 21.57", "Conversão: BRL 5.29 = USD 1 = R$ 5,29"),
        ),
        LinhaRaw(22, "JUL", "•••• 5449 Alda Lanches", d("8.00"), 6),
        LinhaRaw(2, "AGO", "•••• 5449 Super Santiago", d("23.16"), 7),
        LinhaRaw(3, "JUL", "Pagamento em 03 JUL", d("-217.46"), 7),
        LinhaRaw(7, "JUL", "Pagamento em 07 JUL", d("-100.00"), 7),
    )


@pytest.fixture
def raw(linhas_exemplo) -> RawFatura:
    """Fatura sintetica cujas tres invariantes fecham exatas."""
    gastos = sum(l.valor for l in linhas_exemplo if not l.descricao.startswith("Pagamento"))
    pagamentos = -sum(l.valor for l in linhas_exemplo if l.descricao.startswith("Pagamento"))
    anterior = d("2738.70")
    iof = d("1.32")
    outros = d("8.83")
    compras = gastos - iof - outros

    return RawFatura(
        caminho="sintetica.pdf",
        pdf_sha256="0" * 64,
        vencimento_txt="10 AGO 2026",
        periodo_txt="03 JUL a 03 AGO",
        resumo={
            "fatura_anterior": anterior,
            "pagamentos": pagamentos,
            "compras": compras,
            "iof": iof,
            "outros_lancamentos": outros,
            "total_a_pagar": anterior - pagamentos + compras + iof + outros,
        },
        fechamento_proximo_txt="03 SET 2026",
        saldo_aberto_proximo=d("1230.48"),
        saldo_aberto_total=d("1729.71"),
        linhas=linhas_exemplo,
    )


@pytest.fixture
def raw_virada_ano(linhas_exemplo) -> RawFatura:
    """Fatura de janeiro: o periodo vigente cai no ano anterior."""
    linhas = (
        LinhaRaw(29, "NOV", "•••• 5449 Alda Lanches", d("10.00"), 4),
        LinhaRaw(30, "DEZ", "Uber - NuPay", d("20.00"), 4),
        LinhaRaw(2, "DEZ", "Pagamento em 02 DEZ", d("-30.00"), 5),
    )
    return RawFatura(
        caminho="virada.pdf",
        pdf_sha256="1" * 64,
        vencimento_txt="06 JAN 2026",
        periodo_txt="29 NOV a 30 DEZ",
        resumo={
            "fatura_anterior": d("0.00"),
            "pagamentos": d("30.00"),
            "compras": d("30.00"),
            "iof": d("0.00"),
            "outros_lancamentos": d("0.00"),
            "total_a_pagar": d("0.00"),
        },
        fechamento_proximo_txt="30 JAN 2026",
        saldo_aberto_proximo=None,
        saldo_aberto_total=None,
        linhas=linhas,
    )
