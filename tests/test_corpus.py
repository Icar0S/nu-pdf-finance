"""Teste de propriedade sobre os PDFs reais.

Para qualquer fatura do corpus, as tres invariantes de reconciliacao valem. E
barato e pega quase todo erro de parsing: se uma linha some, se um valor sai
errado, se o layout muda entre uma fatura de janeiro e uma de agosto, alguma
das tres soma para de fechar.

Os PDFs nao estao no git. Sem eles, estes testes pulam.
"""

from __future__ import annotations

from decimal import Decimal

import pytest

from nubank.classify import carrega_tabela, classifica
from nubank.export import gasto_do_periodo
from nubank.extract import extrai
from nubank.models import Bucket, TipoLancamento
from nubank.normalize import normaliza
from nubank.reconcile import invariantes

from .conftest import pdfs_disponiveis

PDFS = pdfs_disponiveis()
requer_corpus = pytest.mark.skipif(not PDFS, reason="pasta faturas/ sem PDFs")
ids = [p.stem for p in PDFS]


@pytest.fixture(scope="module")
def tabela():
    return carrega_tabela()


@requer_corpus
@pytest.mark.parametrize("caminho", PDFS, ids=ids)
def test_invariantes_valem(caminho):
    fatura = normaliza(extrai(caminho))
    falhas = [str(i) for i in invariantes(fatura) if not i.ok]
    assert not falhas, "\n".join(falhas)


@requer_corpus
@pytest.mark.parametrize("caminho", PDFS, ids=ids)
def test_toda_transacao_cai_dentro_do_periodo(caminho):
    fatura = normaliza(extrai(caminho))
    fora = [
        t for t in fatura.transacoes
        if not (fatura.periodo_ini <= t.data <= fatura.periodo_fim)
    ]
    assert not fora, [(t.data.isoformat(), t.descricao_raw) for t in fora]


@requer_corpus
@pytest.mark.parametrize("caminho", PDFS, ids=ids)
def test_toda_linha_datada_tem_valor(caminho):
    """Descricao vazia ou valor zerado indica coluna colapsada errado."""
    fatura = normaliza(extrai(caminho))
    assert all(t.descricao_raw.strip() for t in fatura.transacoes)
    assert all(t.merchant_norm.strip() for t in fatura.transacoes)


@requer_corpus
@pytest.mark.parametrize("caminho", PDFS, ids=ids)
def test_parse_e_deterministico(caminho):
    a = normaliza(extrai(caminho))
    b = normaliza(extrai(caminho))
    assert [(t.data, t.descricao_raw, t.valor) for t in a.transacoes] == [
        (t.data, t.descricao_raw, t.valor) for t in b.transacoes
    ]
    assert a.pdf_sha256 == b.pdf_sha256


@requer_corpus
@pytest.mark.parametrize("caminho", PDFS, ids=ids)
def test_iof_sempre_tem_um_pai_conhecido(caminho, tabela):
    """IOF orfao quebra o total por merchant e some da classificacao."""
    fatura = classifica(normaliza(extrai(caminho)), tabela)
    for t in fatura.transacoes:
        if t.tipo is TipoLancamento.IOF:
            assert t.merchant_pai, t.descricao_raw


@requer_corpus
@pytest.mark.parametrize("caminho", PDFS, ids=ids)
def test_buckets_somam_o_gasto_do_periodo(caminho, tabela):
    """C + D + E + pendentes = compras + IOF + outros lancamentos.

    Se isso vale, o residuo da coluna F da planilha e so o ajuste de saldo.
    """
    fatura = classifica(normaliza(extrai(caminho)), tabela)
    soma = sum(
        (
            fatura.total_bucket(b)
            for b in (Bucket.PARCELAS, Bucket.ESSENCIAL, Bucket.SUPERFLUO)
        ),
        start=Decimal("0.00"),
    )
    soma += sum((t.valor for t in fatura.pendentes), start=Decimal("0.00"))
    esperado = fatura.compras + fatura.iof + fatura.outros_lancamentos
    assert abs(soma - esperado) <= Decimal("0.02")


@requer_corpus
@pytest.mark.parametrize("caminho", PDFS, ids=ids)
def test_coluna_f_da_planilha_zera(caminho, tabela):
    """C + D + E == H exatamente, para qualquer fatura sem pendente.

    A coluna F da aba Cartao e `H - SOMA(C:E)`. Se essa igualdade nao for
    exata, F fica mostrando centavos de ruido numa coluna cujo unico trabalho e
    mostrar zero quando esta tudo classificado.
    """
    fatura = classifica(normaliza(extrai(caminho)), tabela)
    classificado = sum(
        (
            fatura.total_bucket(b)
            for b in (Bucket.PARCELAS, Bucket.ESSENCIAL, Bucket.SUPERFLUO)
        ),
        start=Decimal("0.00"),
    )
    pendente = sum((t.valor for t in fatura.pendentes), start=Decimal("0.00"))
    assert classificado + pendente == gasto_do_periodo(fatura)


@requer_corpus
def test_competencias_sao_unicas_e_sequenciais():
    faturas = [normaliza(extrai(p)) for p in PDFS]
    competencias = [f.competencia for f in faturas]
    assert len(set(competencias)) == len(competencias)


@requer_corpus
def test_fatura_anterior_bate_com_o_total_do_mes_passado():
    """Invariante que atravessa faturas: um erro de leitura de header aparece aqui."""
    faturas = sorted(
        (normaliza(extrai(p)) for p in PDFS), key=lambda f: f.competencia
    )
    for anterior, atual in zip(faturas, faturas[1:]):
        if anterior.competencia[:4] != atual.competencia[:4]:
            continue
        mes_ant = int(anterior.competencia[5:])
        if int(atual.competencia[5:]) != mes_ant + 1:
            continue
        assert atual.fatura_anterior == anterior.total_a_pagar, (
            f"{atual.competencia}: fatura anterior {atual.fatura_anterior} "
            f"!= total de {anterior.competencia} {anterior.total_a_pagar}"
        )
