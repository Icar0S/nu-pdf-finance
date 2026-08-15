"""Estagio 4: categoria e bucket.

Tres camadas, todas deterministicas:

  1. regras estruturais  - parcela, IOF, estorno, pagamento. Vem da forma da
                           linha, nao do nome do merchant.
  2. tabela de merchants - match exato na chave normalizada. Cobre a grande
                           maioria, porque gasto de pessoa fisica e repetitivo.
  3. padroes             - prefixo de agregador (ifd*, shopee*, ...) para o
                           merchant novo que ja nasce dentro de um marketplace
                           conhecido.

O que nao casa em nenhuma vira PENDENTE e trava o export ate voce classificar.
Isso e proposital: um merchant desconhecido caindo em algum bucket por padrao
e exatamente o erro silencioso que a planilha nao te deixaria ver.

Rodar o mesmo PDF duas vezes produz sempre o mesmo resultado.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml

from .errors import ErroFatura
from .models import Bucket, Fatura, OrigemCategoria, TipoLancamento, Transacao

TABELA_PADRAO = Path(__file__).resolve().parent.parent / "merchants.yml"

# Categorias sinteticas das regras estruturais. Nao moram no YAML porque nao
# sao editaveis: mudam o significado do pipeline, nao a opiniao sobre um gasto.
CATEGORIA_PARCELAMENTO = "parcelamento"
CATEGORIA_PAGAMENTO = "pagamento"
CATEGORIA_AJUSTE = "ajuste"

# Nao aparecem no menu do review: sao atribuidas pela forma da linha, e escolher
# uma delas a mao para um merchant nao faria sentido nenhum.
CATEGORIAS_ESTRUTURAIS = frozenset(
    {CATEGORIA_PARCELAMENTO, CATEGORIA_PAGAMENTO, CATEGORIA_AJUSTE}
)


@dataclass
class Tabela:
    categorias: dict[str, Bucket]
    merchants: dict[str, str]
    padroes: list[tuple[str, str]]  # (prefixo, categoria), mais longo primeiro

    def bucket_de(self, categoria: str) -> Bucket:
        try:
            return self.categorias[categoria]
        except KeyError:
            raise ErroFatura(
                f"categoria '{categoria}' nao existe em merchants.yml. "
                f"Categorias validas: {', '.join(sorted(self.categorias))}"
            ) from None

    def busca(self, merchant: str) -> tuple[str, OrigemCategoria] | None:
        if merchant in self.merchants:
            return self.merchants[merchant], OrigemCategoria.TABELA
        for prefixo, categoria in self.padroes:
            if merchant.startswith(prefixo):
                return categoria, OrigemCategoria.PADRAO
        return None


def carrega_tabela(caminho: str | Path | None = None) -> Tabela:
    caminho = Path(caminho) if caminho else TABELA_PADRAO
    if not caminho.exists():
        raise ErroFatura(f"tabela de merchants nao encontrada: {caminho}")

    dados = yaml.safe_load(caminho.read_text(encoding="utf-8")) or {}

    categorias: dict[str, Bucket] = {}
    for nome, bucket in (dados.get("categorias") or {}).items():
        try:
            categorias[nome] = Bucket(bucket)
        except ValueError:
            raise ErroFatura(
                f"merchants.yml: categoria '{nome}' aponta para bucket invalido "
                f"'{bucket}'. Use um de: parcelas, essencial, superfluo."
            ) from None

    # Regras estruturais precisam de bucket tambem.
    categorias.setdefault(CATEGORIA_PARCELAMENTO, Bucket.PARCELAS)
    categorias.setdefault(CATEGORIA_PAGAMENTO, Bucket.NAO_GASTO)
    categorias.setdefault(CATEGORIA_AJUSTE, Bucket.NAO_GASTO)

    merchants = {str(k): str(v) for k, v in (dados.get("merchants") or {}).items()}

    padroes = [
        (str(p["prefixo"]), str(p["categoria"])) for p in (dados.get("padroes") or [])
    ]
    padroes.sort(key=lambda p: len(p[0]), reverse=True)

    desconhecidas = {c for c in merchants.values() if c not in categorias}
    desconhecidas |= {c for _, c in padroes if c not in categorias}
    if desconhecidas:
        raise ErroFatura(
            "merchants.yml usa categorias que nao estao declaradas em "
            f"'categorias': {', '.join(sorted(desconhecidas))}"
        )

    return Tabela(categorias=categorias, merchants=merchants, padroes=padroes)


def _aplica(t: Transacao, categoria: str, origem: OrigemCategoria, tabela: Tabela) -> None:
    t.categoria = categoria
    t.origem_categoria = origem
    t.bucket = tabela.bucket_de(categoria)


def classifica_transacao(t: Transacao, tabela: Tabela) -> None:
    """Preenche categoria/bucket/origem de uma transacao, in place."""
    if t.tipo is TipoLancamento.PAGAMENTO:
        _aplica(t, CATEGORIA_PAGAMENTO, OrigemCategoria.REGRA, tabela)
        return
    if t.tipo is TipoLancamento.AJUSTE:
        _aplica(t, CATEGORIA_AJUSTE, OrigemCategoria.REGRA, tabela)
        return

    # Parcelamento e uma decisao estrutural e ganha do merchant: a coluna C da
    # planilha existe justamente para separar compromisso ja assumido de gasto
    # do mes. Uma geladeira em 6x nao e 'compras' em nenhum dos 6 meses.
    if t.tipo is TipoLancamento.PARCELA:
        _aplica(t, CATEGORIA_PARCELAMENTO, OrigemCategoria.REGRA, tabela)
        return

    # IOF e estorno ja carregam o merchant do pai em merchant_norm, entao a
    # busca abaixo os coloca automaticamente no mesmo bucket do lancamento que
    # os originou. E o que impede o IOF de flutuar sem dono.
    achado = tabela.busca(t.merchant_norm)
    if achado:
        categoria, origem = achado
        _aplica(t, categoria, origem, tabela)
        return

    t.categoria = None
    t.bucket = None
    t.origem_categoria = OrigemCategoria.PENDENTE


def classifica(fatura: Fatura, tabela: Tabela) -> Fatura:
    for t in fatura.transacoes:
        classifica_transacao(t, tabela)
    return fatura
