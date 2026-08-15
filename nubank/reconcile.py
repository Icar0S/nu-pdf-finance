"""Estagio 3: o portao.

Cada PDF do Nubank carrega um oraculo de correcao embutido: o resumo e a lista
de transacoes sao dois caminhos independentes para os mesmos numeros. Se os
dois nao fecham, o parser errou, e nenhum dado entra na base.

Nao existe --force. Uma ferramenta que voce precisa auditar linha a linha nao
economiza trabalho nenhum; o valor inteiro disso aqui esta em poder confiar no
resultado sem reconferir.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal

from .errors import ErroReconciliacao
from .models import Fatura, TipoLancamento
from .money import format_brl

# O proprio Nubank arredonda: em 4 das 8 faturas do corpus inicial a soma das
# transacoes bate no resumo com 1 centavo de diferenca. A tolerancia cobre o
# arredondamento deles sem esconder erro de parsing, que erra por muito mais.
TOLERANCIA = Decimal("0.02")


@dataclass(frozen=True)
class Invariante:
    nome: str
    esquerda: Decimal
    direita: Decimal
    descricao: str

    @property
    def diferenca(self) -> Decimal:
        return self.esquerda - self.direita

    @property
    def ok(self) -> bool:
        return abs(self.diferenca) <= TOLERANCIA

    def __str__(self) -> str:
        marca = "ok" if self.ok else "FALHOU"
        return (
            f"[{marca}] {self.nome}: {format_brl(self.esquerda)} vs "
            f"{format_brl(self.direita)} (dif {format_brl(self.diferenca)})"
        )


def invariantes(fatura: Fatura) -> list[Invariante]:
    """As tres somas que precisam fechar para a fatura ser considerada lida."""
    gastos = sum(
        (t.valor for t in fatura.transacoes if t.eh_gasto), start=Decimal("0.00")
    )
    pagamentos_listados = sum(
        (
            -t.valor
            for t in fatura.transacoes
            if t.tipo is TipoLancamento.PAGAMENTO
        ),
        start=Decimal("0.00"),
    )
    lancamentos = fatura.compras + fatura.iof + fatura.outros_lancamentos

    return [
        Invariante(
            nome="composicao do total",
            esquerda=fatura.fatura_anterior - fatura.pagamentos + lancamentos,
            direita=fatura.total_a_pagar,
            descricao="fatura anterior - pagamentos + compras + IOF + outros = total a pagar",
        ),
        Invariante(
            nome="soma das transacoes",
            esquerda=gastos,
            direita=lancamentos,
            descricao="soma das transacoes do periodo = compras + IOF + outros lancamentos",
        ),
        Invariante(
            nome="soma dos pagamentos",
            esquerda=pagamentos_listados,
            direita=fatura.pagamentos,
            descricao="soma dos pagamentos listados = pagamento recebido no resumo",
        ),
    ]


def reconcilia(fatura: Fatura) -> list[Invariante]:
    """Roda as invariantes. Levanta ErroReconciliacao se alguma falhar."""
    resultado = invariantes(fatura)
    falhas = [i for i in resultado if not i.ok]
    if falhas:
        detalhe = "\n".join(f"  {i}\n    ({i.descricao})" for i in falhas)
        raise ErroReconciliacao(
            f"{fatura.arquivo}: a fatura nao fecha, import rejeitado.\n{detalhe}"
        )
    return resultado
