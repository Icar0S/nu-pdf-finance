from datetime import datetime
from decimal import Decimal

import openpyxl
import pytest

from nubank.classify import carrega_tabela, classifica
from nubank.errors import ErroExport
from nubank.export import aplica, exporta_csv, observacao, planeja
from nubank.models import Bucket
from nubank.normalize import normaliza

from .conftest import RAIZ

PLANILHA_REAL = RAIZ / "planejamento-avancado-financeiro-icaro26.xlsx"


@pytest.fixture
def planilha(tmp_path):
    """Copia minima da aba Cartao, com a formula da coluna F no lugar."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cartao"
    ws["A3"] = "Mes"
    ws["B3"] = "Fatura total"
    ws["C3"] = "Parcelas e assinaturas"
    ws["D3"] = "Essencial (mercado, farmacia, transporte)"
    ws["E3"] = "Superfluo (delivery, compras, saidas)"
    ws["F3"] = "Nao classificado"
    ws["G3"] = "Observacoes"
    for i, mes in enumerate(range(1, 13)):
        linha = 4 + i
        ws[f"A{linha}"] = datetime(2026, mes, 1)
        ws[f"F{linha}"] = f'=IF($B{linha}="","",$B{linha}-SUM($C{linha}:$E{linha}))'
    wb.create_sheet("Outra Aba")["A1"] = "nao pode ser tocada"
    caminho = tmp_path / "planilha.xlsx"
    wb.save(caminho)
    return caminho


@pytest.fixture
def fatura(raw):
    return classifica(normaliza(raw), carrega_tabela())


def test_dry_run_nao_escreve(planilha, fatura):
    antes = planilha.read_bytes()
    mudancas = planeja(planilha, [fatura])
    assert planilha.read_bytes() == antes
    assert any(m.mudou for m in mudancas)


def test_escreve_na_linha_certa(planilha, fatura):
    aplica(planilha, [fatura])
    ws = openpyxl.load_workbook(planilha)["Cartao"]
    linha = 4 + 7  # competencia 2026-08 -> agosto
    assert ws[f"B{linha}"].value == pytest.approx(float(fatura.total_a_pagar))
    assert ws[f"C{linha}"].value == pytest.approx(
        float(fatura.total_bucket(Bucket.PARCELAS))
    )
    assert ws["B4"].value is None  # janeiro fica intacto


def test_nao_escreve_na_coluna_de_formula(planilha, fatura):
    aplica(planilha, [fatura])
    ws = openpyxl.load_workbook(planilha)["Cartao"]
    assert str(ws["F11"].value).startswith("=IF(")


def test_backup_e_criado_antes_de_escrever(planilha, fatura):
    original = planilha.read_bytes()
    backup, _ = aplica(planilha, [fatura])
    assert backup.exists()
    assert backup.read_bytes() == original
    assert planilha.read_bytes() != original


def test_outras_abas_sobrevivem(planilha, fatura):
    aplica(planilha, [fatura])
    wb = openpyxl.load_workbook(planilha)
    assert wb["Outra Aba"]["A1"].value == "nao pode ser tocada"


def test_observacao_explica_o_residuo_da_coluna_f(fatura):
    texto = observacao(fatura)
    assert "gasto do periodo" in texto
    assert "ajuste de saldo" in texto  # a fixture tem fatura_anterior != pagamentos


def test_segunda_aplicacao_nao_muda_nada(planilha, fatura):
    aplica(planilha, [fatura])
    assert not [m for m in planeja(planilha, [fatura]) if m.mudou]


def test_planilha_sem_a_aba_cartao(tmp_path, fatura):
    caminho = tmp_path / "vazia.xlsx"
    openpyxl.Workbook().save(caminho)
    with pytest.raises(ErroExport, match="aba 'Cartao'"):
        planeja(caminho, [fatura])


def test_cabecalho_inesperado(planilha, fatura):
    wb = openpyxl.load_workbook(planilha)
    wb["Cartao"]["B3"] = "Outra coisa"
    wb.save(planilha)
    with pytest.raises(ErroExport, match="formato esperado"):
        planeja(planilha, [fatura])


def test_competencia_sem_linha_na_planilha(planilha, fatura):
    wb = openpyxl.load_workbook(planilha)
    wb["Cartao"]["A11"] = datetime(2027, 8, 1)
    wb.save(planilha)
    with pytest.raises(ErroExport, match="2026-08"):
        planeja(planilha, [fatura])


def test_csv_de_detalhe(tmp_path, fatura):
    destino = exporta_csv(tmp_path / "detalhe.csv", [fatura])
    linhas = destino.read_text(encoding="utf-8-sig").splitlines()
    assert linhas[0].startswith("competencia;data;descricao")
    assert len(linhas) == 1 + len(fatura.transacoes)


@pytest.mark.skipif(not PLANILHA_REAL.exists(), reason="planilha real ausente")
def test_round_trip_na_planilha_real_preserva_estrutura(tmp_path, fatura):
    """openpyxl reescreve o arquivo inteiro; garante que nada se perde no caminho."""
    import shutil

    copia = tmp_path / PLANILHA_REAL.name
    shutil.copy2(PLANILHA_REAL, copia)

    antes = openpyxl.load_workbook(copia)
    abas_antes = antes.sheetnames
    formatos_antes = {
        ws.title: len(ws.conditional_formatting._cf_rules) for ws in antes.worksheets
    }
    formula_antes = antes["Controle Mensal"]["F11"].value
    parametro_antes = antes["Parametros"]["C6"].value

    aplica(copia, [fatura])

    depois = openpyxl.load_workbook(copia)
    assert depois.sheetnames == abas_antes
    assert {
        ws.title: len(ws.conditional_formatting._cf_rules) for ws in depois.worksheets
    } == formatos_antes
    assert depois["Controle Mensal"]["F11"].value == formula_antes
    assert depois["Parametros"]["C6"].value == parametro_antes
