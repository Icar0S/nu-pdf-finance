"""Valores monetarios em Decimal, nunca float.

Toda a aritmetica da fatura roda em centavos exatos. float quebraria as
invariantes de reconciliacao por erro de representacao antes mesmo de o
parser errar alguma coisa.
"""

from __future__ import annotations

import re
from decimal import Decimal, ROUND_HALF_UP

CENTAVO = Decimal("0.01")

# Nubank usa MINUS SIGN (U+2212) nos estornos/pagamentos dentro da lista de
# transacoes e HYPHEN-MINUS comum nos totais de secao. Aceitamos os dois.
SINAIS_NEGATIVOS = "-−–—"

_RE_VALOR = re.compile(
    r"(?P<sinal>[" + SINAIS_NEGATIVOS + r"])?\s*R\$\s*(?P<num>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})"
)


def parse_brl(texto: str) -> Decimal:
    """'1.234,56' -> Decimal('1234.56'). Aceita sinal negativo a esquerda."""
    t = texto.strip()
    negativo = bool(t) and t[0] in SINAIS_NEGATIVOS
    if negativo:
        t = t[1:].strip()
    t = t.replace("R$", "").strip().replace(".", "").replace(",", ".")
    valor = Decimal(t).quantize(CENTAVO, rounding=ROUND_HALF_UP)
    return -valor if negativo else valor


def encontra_valor(linha: str) -> Decimal | None:
    """Extrai o ultimo 'R$ x' da linha, respeitando o sinal. None se nao houver."""
    achados = list(_RE_VALOR.finditer(linha))
    if not achados:
        return None
    m = achados[-1]
    valor = parse_brl(m.group("num"))
    return -valor if m.group("sinal") else valor


def format_brl(valor: Decimal) -> str:
    """Decimal('1234.56') -> 'R$ 1.234,56'."""
    negativo = valor < 0
    inteiro = f"{abs(valor):,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return f"{'-' if negativo else ''}R$ {inteiro}"


def to_float(valor: Decimal) -> float:
    """Conversao explicita para a fronteira com o Excel, que so fala float."""
    return float(valor)
