"""Persistencia em SQLite.

Guardar o parse tem duas razoes praticas: reimportar o mesmo PDF vira no-op
(dedup por sha256), e mudar uma regra de categoria permite reclassificar os
oito meses de historico sem reparsear PDF nenhum, o que torna barato comparar
o antes e o depois de uma regra nova.

Decimal vai para o banco como TEXT. SQLite so tem REAL para numero, e REAL e
float: bastaria isso para quebrar as invariantes de reconciliacao na volta.
"""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from .models import Bucket, Fatura, OrigemCategoria, TipoLancamento, Transacao

BANCO_PADRAO = Path(__file__).resolve().parent.parent / "faturas.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS faturas (
    id                   INTEGER PRIMARY KEY,
    competencia          TEXT NOT NULL UNIQUE,
    vencimento           TEXT NOT NULL,
    periodo_ini          TEXT NOT NULL,
    periodo_fim          TEXT NOT NULL,
    fatura_anterior      TEXT NOT NULL,
    pagamentos           TEXT NOT NULL,
    compras              TEXT NOT NULL,
    iof                  TEXT NOT NULL,
    outros_lancamentos   TEXT NOT NULL,
    total_a_pagar        TEXT NOT NULL,
    saldo_aberto_proximo TEXT,
    saldo_aberto_total   TEXT,
    pdf_sha256           TEXT NOT NULL UNIQUE,
    arquivo              TEXT NOT NULL,
    importado_em         TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS transacoes (
    id               INTEGER PRIMARY KEY,
    fatura_id        INTEGER NOT NULL REFERENCES faturas(id) ON DELETE CASCADE,
    data             TEXT NOT NULL,
    descricao_raw    TEXT NOT NULL,
    merchant_norm    TEXT NOT NULL,
    valor            TEXT NOT NULL,
    tipo             TEXT NOT NULL,
    cartao_final     TEXT,
    parcela_atual    INTEGER,
    parcela_total    INTEGER,
    merchant_pai     TEXT,
    moeda_origem     TEXT,
    valor_origem     TEXT,
    taxa_cambio      TEXT,
    categoria        TEXT,
    bucket           TEXT,
    origem_categoria TEXT
);

CREATE INDEX IF NOT EXISTS ix_tx_fatura   ON transacoes(fatura_id);
CREATE INDEX IF NOT EXISTS ix_tx_merchant ON transacoes(merchant_norm);
"""

_COLUNAS_TX = (
    "data", "descricao_raw", "merchant_norm", "valor", "tipo", "cartao_final",
    "parcela_atual", "parcela_total", "merchant_pai", "moeda_origem",
    "valor_origem", "taxa_cambio", "categoria", "bucket", "origem_categoria",
)


def conecta(caminho: str | Path | None = None) -> sqlite3.Connection:
    caminho = Path(caminho) if caminho else BANCO_PADRAO
    conn = sqlite3.connect(caminho)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA)
    return conn


def _dec(valor: str | None) -> Decimal | None:
    return Decimal(valor) if valor is not None else None


def ja_importada(conn: sqlite3.Connection, pdf_sha256: str) -> bool:
    cur = conn.execute("SELECT 1 FROM faturas WHERE pdf_sha256 = ?", (pdf_sha256,))
    return cur.fetchone() is not None


def competencia_existente(conn: sqlite3.Connection, competencia: str) -> str | None:
    """Devolve o arquivo ja gravado nessa competencia, se houver."""
    cur = conn.execute(
        "SELECT arquivo FROM faturas WHERE competencia = ?", (competencia,)
    )
    linha = cur.fetchone()
    return linha["arquivo"] if linha else None


def grava(conn: sqlite3.Connection, fatura: Fatura) -> int:
    """Grava a fatura, substituindo o que houver na mesma competencia.

    Uma competencia tem exatamente uma fatura. Se voce rebaixar o PDF do mesmo
    mes, a versao nova manda.
    """
    with conn:
        conn.execute("DELETE FROM faturas WHERE competencia = ?", (fatura.competencia,))
        cur = conn.execute(
            """INSERT INTO faturas (
                   competencia, vencimento, periodo_ini, periodo_fim,
                   fatura_anterior, pagamentos, compras, iof, outros_lancamentos,
                   total_a_pagar, saldo_aberto_proximo, saldo_aberto_total,
                   pdf_sha256, arquivo, importado_em)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                fatura.competencia,
                fatura.vencimento.isoformat(),
                fatura.periodo_ini.isoformat(),
                fatura.periodo_fim.isoformat(),
                str(fatura.fatura_anterior),
                str(fatura.pagamentos),
                str(fatura.compras),
                str(fatura.iof),
                str(fatura.outros_lancamentos),
                str(fatura.total_a_pagar),
                str(fatura.saldo_aberto_proximo) if fatura.saldo_aberto_proximo is not None else None,
                str(fatura.saldo_aberto_total) if fatura.saldo_aberto_total is not None else None,
                fatura.pdf_sha256,
                fatura.arquivo,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        fatura_id = int(cur.lastrowid)
        conn.executemany(
            f"INSERT INTO transacoes (fatura_id, {', '.join(_COLUNAS_TX)}) "
            f"VALUES (?{', ?' * len(_COLUNAS_TX)})",
            [
                (
                    fatura_id,
                    t.data.isoformat(),
                    t.descricao_raw,
                    t.merchant_norm,
                    str(t.valor),
                    t.tipo.value,
                    t.cartao_final,
                    t.parcela_atual,
                    t.parcela_total,
                    t.merchant_pai,
                    t.moeda_origem,
                    str(t.valor_origem) if t.valor_origem is not None else None,
                    str(t.taxa_cambio) if t.taxa_cambio is not None else None,
                    t.categoria,
                    t.bucket.value if t.bucket else None,
                    t.origem_categoria.value if t.origem_categoria else None,
                )
                for t in fatura.transacoes
            ],
        )
    return fatura_id


def _monta_transacao(linha: sqlite3.Row) -> Transacao:
    return Transacao(
        data=date.fromisoformat(linha["data"]),
        descricao_raw=linha["descricao_raw"],
        merchant_norm=linha["merchant_norm"],
        valor=Decimal(linha["valor"]),
        tipo=TipoLancamento(linha["tipo"]),
        cartao_final=linha["cartao_final"],
        parcela_atual=linha["parcela_atual"],
        parcela_total=linha["parcela_total"],
        merchant_pai=linha["merchant_pai"],
        moeda_origem=linha["moeda_origem"],
        valor_origem=_dec(linha["valor_origem"]),
        taxa_cambio=_dec(linha["taxa_cambio"]),
        categoria=linha["categoria"],
        bucket=Bucket(linha["bucket"]) if linha["bucket"] else None,
        origem_categoria=(
            OrigemCategoria(linha["origem_categoria"])
            if linha["origem_categoria"]
            else None
        ),
    )


def carrega(
    conn: sqlite3.Connection, competencia: str | None = None
) -> list[Fatura]:
    """Le as faturas gravadas, em ordem de competencia."""
    if competencia:
        cur = conn.execute(
            "SELECT * FROM faturas WHERE competencia = ? ORDER BY competencia",
            (competencia,),
        )
    else:
        cur = conn.execute("SELECT * FROM faturas ORDER BY competencia")

    faturas = []
    for linha in cur.fetchall():
        transacoes = [
            _monta_transacao(t)
            for t in conn.execute(
                "SELECT * FROM transacoes WHERE fatura_id = ? ORDER BY id",
                (linha["id"],),
            )
        ]
        faturas.append(
            Fatura(
                competencia=linha["competencia"],
                vencimento=date.fromisoformat(linha["vencimento"]),
                periodo_ini=date.fromisoformat(linha["periodo_ini"]),
                periodo_fim=date.fromisoformat(linha["periodo_fim"]),
                fatura_anterior=Decimal(linha["fatura_anterior"]),
                pagamentos=Decimal(linha["pagamentos"]),
                compras=Decimal(linha["compras"]),
                iof=Decimal(linha["iof"]),
                outros_lancamentos=Decimal(linha["outros_lancamentos"]),
                total_a_pagar=Decimal(linha["total_a_pagar"]),
                saldo_aberto_proximo=_dec(linha["saldo_aberto_proximo"]),
                saldo_aberto_total=_dec(linha["saldo_aberto_total"]),
                pdf_sha256=linha["pdf_sha256"],
                arquivo=linha["arquivo"],
                transacoes=transacoes,
            )
        )
    return faturas


def atualiza_classificacao(conn: sqlite3.Connection, faturas: list[Fatura]) -> int:
    """Reescreve categoria/bucket/origem das transacoes ja gravadas.

    Casa por (competencia, ordem), que e estavel porque a ordem de leitura do
    PDF e deterministica.
    """
    total = 0
    with conn:
        for fatura in faturas:
            ids = [
                linha["id"]
                for linha in conn.execute(
                    "SELECT t.id FROM transacoes t "
                    "JOIN faturas f ON f.id = t.fatura_id "
                    "WHERE f.competencia = ? ORDER BY t.id",
                    (fatura.competencia,),
                )
            ]
            for tx_id, t in zip(ids, fatura.transacoes):
                conn.execute(
                    "UPDATE transacoes SET categoria = ?, bucket = ?, "
                    "origem_categoria = ? WHERE id = ?",
                    (
                        t.categoria,
                        t.bucket.value if t.bucket else None,
                        t.origem_categoria.value if t.origem_categoria else None,
                        tx_id,
                    ),
                )
                total += 1
    return total
