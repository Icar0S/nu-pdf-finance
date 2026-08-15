"""Estagio 5: o unico estagio que toca disco.

Escreve na aba Cartao da planilha de planejamento. Tres cuidados, todos por
motivo concreto:

  1. backup timestampado antes de qualquer escrita;
  2. dry-run por padrao, mostrando celula a celula o que muda;
  3. a coluna F nunca e escrita - ela e formula (=B-SOMA(C:E)) e o residuo
     dela tem significado, explicado abaixo.

Sobre a coluna F: 'Fatura total' (B) e 'quanto voce gastou no periodo' sao
grandezas diferentes. B inclui o saldo que veio da fatura anterior menos o que
voce pagou; C+D+E cobrem so os lancamentos do periodo. A diferenca e
exatamente `fatura_anterior - pagamentos`, e e por isso que F fica em -49,99
no mes em que voce pagou R$ 50 a mais. Nao e gasto perdido, e ajuste de saldo,
e a observacao na coluna G diz isso em cada linha.
"""

from __future__ import annotations

import csv
import shutil
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import openpyxl

from .errors import ErroExport
from .models import Bucket, Fatura
from .money import format_brl, to_float

ABA = "Cartao"
LINHA_CABECALHO = 3
COLUNA_MES = "A"

# Coluna da planilha <- bucket. F fica de fora de proposito: e formula.
COLUNAS = {
    "B": None,  # Fatura total: total a pagar, nao soma de bucket
    "C": Bucket.PARCELAS,
    "D": Bucket.ESSENCIAL,
    "E": Bucket.SUPERFLUO,
}

CABECALHO_ESPERADO = {
    "A": "Mes",
    "B": "Fatura total",
    "C": "Parcelas e assinaturas",
}


@dataclass(frozen=True)
class Mudanca:
    celula: str
    rotulo: str
    antes: object
    depois: object

    @property
    def mudou(self) -> bool:
        if isinstance(self.antes, float) and isinstance(self.depois, float):
            return abs(self.antes - self.depois) > 1e-9
        return self.antes != self.depois

    def _fmt(self, valor: object) -> str:
        if valor is None or valor == "":
            return "(vazio)"
        if isinstance(valor, float):
            return format_brl(Decimal(str(valor)))
        return str(valor)

    def __str__(self) -> str:
        return f"{self.celula} {self.rotulo:<24} {self._fmt(self.antes)} -> {self._fmt(self.depois)}"


def observacao(fatura: Fatura) -> str:
    """Texto da coluna G. Existe para o residuo da coluna F nao virar caca ao bug."""
    gasto = fatura.compras + fatura.iof + fatura.outros_lancamentos
    periodo = f"{fatura.periodo_ini:%d/%m}-{fatura.periodo_fim:%d/%m}"
    n = sum(1 for t in fatura.transacoes if t.eh_gasto)
    texto = f"{periodo} | {n} lanc. | gasto do periodo {format_brl(gasto)}"
    ajuste = fatura.ajuste_saldo
    if ajuste:
        texto += f" | col.F = ajuste de saldo anterior {format_brl(ajuste)}"
    return texto


def _localiza_linha(ws, competencia: str) -> int:
    ano, mes = (int(p) for p in competencia.split("-"))
    for linha in range(LINHA_CABECALHO + 1, ws.max_row + 1):
        valor = ws[f"{COLUNA_MES}{linha}"].value
        if isinstance(valor, datetime) and valor.year == ano and valor.month == mes:
            return linha
    raise ErroExport(
        f"a aba '{ABA}' nao tem linha para a competencia {competencia}. "
        f"A coluna {COLUNA_MES} precisa ter uma data desse mes."
    )


def _valida_planilha(ws) -> None:
    for coluna, esperado in CABECALHO_ESPERADO.items():
        atual = ws[f"{coluna}{LINHA_CABECALHO}"].value
        if atual is None or not str(atual).strip().lower().startswith(esperado.lower()):
            raise ErroExport(
                f"a aba '{ABA}' nao tem o formato esperado: {coluna}{LINHA_CABECALHO} "
                f"deveria comecar com '{esperado}', mas contem {atual!r}."
            )


def planeja(planilha: str | Path, faturas: list[Fatura]) -> list[Mudanca]:
    """Calcula as mudancas sem escrever nada."""
    planilha = Path(planilha)
    if not planilha.exists():
        raise ErroExport(f"planilha nao encontrada: {planilha}")

    wb = openpyxl.load_workbook(planilha)
    if ABA not in wb.sheetnames:
        raise ErroExport(
            f"a planilha nao tem a aba '{ABA}'. Abas: {', '.join(wb.sheetnames)}"
        )
    ws = wb[ABA]
    _valida_planilha(ws)

    mudancas: list[Mudanca] = []
    for fatura in faturas:
        linha = _localiza_linha(ws, fatura.competencia)
        alvos = {
            "B": to_float(fatura.total_a_pagar),
            "C": to_float(fatura.total_bucket(Bucket.PARCELAS)),
            "D": to_float(fatura.total_bucket(Bucket.ESSENCIAL)),
            "E": to_float(fatura.total_bucket(Bucket.SUPERFLUO)),
            "G": observacao(fatura),
        }
        for coluna, novo in alvos.items():
            celula = f"{coluna}{linha}"
            rotulo = str(ws[f"{coluna}{LINHA_CABECALHO}"].value or coluna)
            mudancas.append(
                Mudanca(
                    celula=f"{fatura.competencia} {celula}",
                    rotulo=rotulo[:24],
                    antes=ws[celula].value,
                    depois=novo,
                )
            )
    wb.close()
    return mudancas


def faz_backup(planilha: Path, pasta: Path | None = None) -> Path:
    pasta = pasta or planilha.parent / "backups"
    pasta.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d-%H%M%S")
    destino = pasta / f"{planilha.stem}-{carimbo}{planilha.suffix}"
    shutil.copy2(planilha, destino)
    return destino


def aplica(planilha: str | Path, faturas: list[Fatura]) -> tuple[Path, list[Mudanca]]:
    """Escreve de verdade. Devolve (caminho do backup, mudancas aplicadas)."""
    planilha = Path(planilha)
    mudancas = [m for m in planeja(planilha, faturas) if m.mudou]
    backup = faz_backup(planilha)

    wb = openpyxl.load_workbook(planilha)
    ws = wb[ABA]
    for fatura in faturas:
        linha = _localiza_linha(ws, fatura.competencia)
        ws[f"B{linha}"] = to_float(fatura.total_a_pagar)
        ws[f"C{linha}"] = to_float(fatura.total_bucket(Bucket.PARCELAS))
        ws[f"D{linha}"] = to_float(fatura.total_bucket(Bucket.ESSENCIAL))
        ws[f"E{linha}"] = to_float(fatura.total_bucket(Bucket.SUPERFLUO))
        ws[f"G{linha}"] = observacao(fatura)
    wb.save(planilha)
    wb.close()
    return backup, mudancas


def exporta_csv(destino: str | Path, faturas: list[Fatura]) -> Path:
    """CSV de detalhe, uma linha por transacao. Para conferir a mao quando quiser."""
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    with open(destino, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "competencia", "data", "descricao", "merchant", "valor", "tipo",
                "categoria", "bucket", "parcela", "cartao", "moeda_origem",
                "valor_origem", "origem_categoria",
            ]
        )
        for fatura in faturas:
            for t in fatura.transacoes:
                parcela = (
                    f"{t.parcela_atual}/{t.parcela_total}" if t.parcela_total else ""
                )
                w.writerow(
                    [
                        fatura.competencia,
                        t.data.isoformat(),
                        t.descricao_raw,
                        t.merchant_norm,
                        str(t.valor).replace(".", ","),
                        t.tipo.value,
                        t.categoria or "",
                        t.bucket.value if t.bucket else "",
                        parcela,
                        t.cartao_final or "",
                        t.moeda_origem or "",
                        str(t.valor_origem).replace(".", ",") if t.valor_origem else "",
                        t.origem_categoria.value if t.origem_categoria else "",
                    ]
                )
    return destino
