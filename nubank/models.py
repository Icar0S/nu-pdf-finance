"""Estruturas de dados do pipeline.

extract -> RawFatura -> normalize -> Fatura -> reconcile -> classify -> export

Os quatro primeiros estagios sao funcoes puras (bytes -> dataclass). Só o
export toca disco. Isso e o que torna o teste trivial e permite reclassificar
o historico inteiro sem reparsear PDF nenhum.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from enum import Enum


class TipoLancamento(str, Enum):
    """O que a linha e, estruturalmente. Nada a ver com categoria de gasto."""

    COMPRA = "compra"
    PARCELA = "parcela"
    IOF = "iof"
    ESTORNO = "estorno"
    PAGAMENTO = "pagamento"
    AJUSTE = "ajuste"  # 'Saldo restante da fatura anterior' e afins


class Bucket(str, Enum):
    """As tres colunas da aba Cartao. Todo gasto cai em exatamente uma."""

    PARCELAS = "parcelas"  # coluna C - Parcelas e assinaturas
    ESSENCIAL = "essencial"  # coluna D - mercado, farmacia, transporte
    SUPERFLUO = "superfluo"  # coluna E - delivery, compras, saidas
    NAO_GASTO = "nao_gasto"  # pagamentos e ajustes: nao sao gasto do periodo


class OrigemCategoria(str, Enum):
    REGRA = "regra"  # regra estrutural (parcela, IOF herda do pai, pagamento)
    TABELA = "tabela"  # match exato na tabela de merchants
    PADRAO = "padrao"  # match por padrao/prefixo (ifd*, shopee*, ...)
    PENDENTE = "pendente"  # merchant desconhecido, aguardando revisao


@dataclass(frozen=True)
class LinhaRaw:
    """Linha de transacao como saiu do PDF, sem interpretacao nenhuma."""

    dia: int
    mes_abrev: str
    descricao: str
    valor: Decimal
    pagina: int
    conversao: tuple[str, ...] = ()  # linhas de cambio que seguem a transacao


@dataclass(frozen=True)
class RawFatura:
    """Saida do extract: header + linhas, ainda sem tipos de dominio."""

    caminho: str
    pdf_sha256: str
    vencimento_txt: str
    periodo_txt: str
    resumo: dict[str, Decimal]
    fechamento_proximo_txt: str | None
    saldo_aberto_proximo: Decimal | None
    saldo_aberto_total: Decimal | None
    linhas: tuple[LinhaRaw, ...]


@dataclass
class Transacao:
    data: date
    descricao_raw: str
    merchant_norm: str
    valor: Decimal
    tipo: TipoLancamento
    cartao_final: str | None = None
    parcela_atual: int | None = None
    parcela_total: int | None = None
    merchant_pai: str | None = None  # para IOF e estorno: de quem veio
    moeda_origem: str | None = None
    valor_origem: Decimal | None = None
    taxa_cambio: Decimal | None = None
    categoria: str | None = None
    bucket: Bucket | None = None
    origem_categoria: OrigemCategoria | None = None

    @property
    def eh_gasto(self) -> bool:
        """Pagamentos e ajustes nao sao gasto do periodo."""
        return self.tipo not in (TipoLancamento.PAGAMENTO, TipoLancamento.AJUSTE)


@dataclass
class Fatura:
    competencia: str  # 'YYYY-MM' do vencimento - e a chave da aba Cartao
    vencimento: date
    periodo_ini: date
    periodo_fim: date
    fatura_anterior: Decimal
    pagamentos: Decimal  # positivo; e o quanto foi pago
    compras: Decimal
    iof: Decimal
    outros_lancamentos: Decimal
    total_a_pagar: Decimal
    saldo_aberto_proximo: Decimal | None
    saldo_aberto_total: Decimal | None
    pdf_sha256: str
    arquivo: str
    transacoes: list[Transacao] = field(default_factory=list)

    @property
    def ajuste_saldo(self) -> Decimal:
        """fatura_anterior - pagamentos.

        E a diferenca entre 'total a pagar' e 'o que voce gastou no periodo'.
        Quando voce paga a mais, isso vira credito e aparece negativo aqui.
        E exatamente o residuo que sobra na coluna F da planilha.
        """
        return self.fatura_anterior - self.pagamentos

    def total_bucket(self, bucket: Bucket) -> Decimal:
        return sum(
            (t.valor for t in self.transacoes if t.bucket is bucket),
            start=Decimal("0.00"),
        )

    @property
    def pendentes(self) -> list[Transacao]:
        return [
            t
            for t in self.transacoes
            if t.eh_gasto and t.origem_categoria is OrigemCategoria.PENDENTE
        ]
