from decimal import Decimal

import pytest

from nubank.classify import carrega_tabela, classifica
from nubank.errors import ErroFatura
from nubank.models import Bucket, OrigemCategoria, TipoLancamento
from nubank.normalize import normaliza

TABELA_YML = """
categorias:
  assinatura: parcelas
  transporte: essencial
  delivery: superfluo
padroes:
  - prefixo: "ifd*"
    categoria: delivery
merchants:
  "uber - nupay": transporte
  "render.com": assinatura
"""


@pytest.fixture
def tabela(tmp_path):
    caminho = tmp_path / "merchants.yml"
    caminho.write_text(TABELA_YML, encoding="utf-8")
    return carrega_tabela(caminho)


def test_parcela_ganha_do_merchant(raw, tabela):
    """Compra em 6x nao e 'compras' em nenhum dos 6 meses: e compromisso."""
    fatura = classifica(normaliza(raw), tabela)
    parcela = next(t for t in fatura.transacoes if t.tipo is TipoLancamento.PARCELA)
    assert parcela.bucket is Bucket.PARCELAS
    assert parcela.origem_categoria is OrigemCategoria.REGRA


def test_iof_cai_no_bucket_do_pai(raw, tabela):
    fatura = classifica(normaliza(raw), tabela)
    iof = next(t for t in fatura.transacoes if t.tipo is TipoLancamento.IOF)
    assert iof.categoria == "assinatura"
    assert iof.bucket is Bucket.PARCELAS


def test_estorno_cai_no_bucket_do_pai(raw, tabela):
    """Estorno tem de abater no mesmo balde onde o gasto entrou."""
    fatura = classifica(normaliza(raw), tabela)
    estorno = next(t for t in fatura.transacoes if t.tipo is TipoLancamento.ESTORNO)
    assert estorno.categoria == "transporte"
    assert estorno.bucket is Bucket.ESSENCIAL
    assert estorno.valor < 0


def test_padrao_de_agregador(raw, tabela):
    """Restaurante inedito no iFood ja nasce delivery, sem aprovacao manual."""
    fatura = classifica(normaliza(raw), tabela)
    tx = next(t for t in fatura.transacoes if t.merchant_norm.startswith("ifd*"))
    assert tx.categoria == "delivery"
    assert tx.origem_categoria is OrigemCategoria.PADRAO


def test_merchant_desconhecido_fica_pendente(raw, tabela):
    """Nunca cair num bucket por padrao: erro silencioso e o que se quer evitar."""
    fatura = classifica(normaliza(raw), tabela)
    pendentes = {t.merchant_norm for t in fatura.pendentes}
    assert "alda lanches" in pendentes
    assert all(t.bucket is None for t in fatura.pendentes)


def test_pagamento_nao_entra_em_bucket_de_gasto(raw, tabela):
    fatura = classifica(normaliza(raw), tabela)
    pagamento = next(t for t in fatura.transacoes if t.tipo is TipoLancamento.PAGAMENTO)
    assert pagamento.bucket is Bucket.NAO_GASTO
    assert pagamento not in fatura.pendentes


def test_classificacao_e_deterministica(raw, tabela):
    """Rodar duas vezes tem de dar o mesmo resultado."""
    a = [(t.categoria, t.bucket) for t in classifica(normaliza(raw), tabela).transacoes]
    b = [(t.categoria, t.bucket) for t in classifica(normaliza(raw), tabela).transacoes]
    assert a == b


def test_bucket_invalido_no_yaml(tmp_path):
    caminho = tmp_path / "m.yml"
    caminho.write_text("categorias:\n  x: inventado\n", encoding="utf-8")
    with pytest.raises(ErroFatura, match="bucket invalido"):
        carrega_tabela(caminho)


def test_categoria_nao_declarada_no_yaml(tmp_path):
    caminho = tmp_path / "m.yml"
    caminho.write_text(
        'categorias:\n  a: essencial\nmerchants:\n  "x": inexistente\n', encoding="utf-8"
    )
    with pytest.raises(ErroFatura, match="nao estao declaradas"):
        carrega_tabela(caminho)


def test_tabela_do_projeto_carrega():
    """merchants.yml versionado precisa estar sempre valido."""
    tabela = carrega_tabela()
    assert tabela.merchants
    assert all(c in tabela.categorias for c in tabela.merchants.values())


def test_buckets_cobrem_o_gasto_do_periodo(raw, tabela):
    """C + D + E tem de somar exatamente compras + IOF + outros lancamentos.

    Quando isso vale e nao ha pendente, o residuo da coluna F da planilha e
    exatamente o ajuste de saldo, e nada mais.
    """
    fatura = classifica(normaliza(raw), tabela)
    for t in fatura.pendentes:  # a fixture tem merchants fora da tabela minima
        t.bucket = Bucket.SUPERFLUO

    soma = sum(
        (fatura.total_bucket(b) for b in (Bucket.PARCELAS, Bucket.ESSENCIAL, Bucket.SUPERFLUO)),
        start=Decimal("0.00"),
    )
    assert soma == fatura.compras + fatura.iof + fatura.outros_lancamentos
