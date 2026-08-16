"""Interface de linha de comando.

    python -m nubank import faturas/*.pdf
    python -m nubank status
    python -m nubank review
    python -m nubank export            # dry-run
    python -m nubank export --apply
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from . import store
from .classify import CATEGORIAS_ESTRUTURAIS, carrega_tabela, classifica
from .conferir import confere, repara, valores_esperados
from .errors import ErroFatura, ErroReconciliacao
from .export import aplica, exporta_csv, faz_backup, planeja
from .extract import extrai
from .models import Bucket, Fatura, TipoLancamento
from .money import format_brl
from .normalize import normaliza
from .reconcile import reconcilia

RAIZ = Path(__file__).resolve().parent.parent

# Cabecalho do bloco que o review escreve no merchants.yml. Reaproveitado entre
# execucoes para nao acumular um cabecalho por mes.
MARCADOR_REVIEW = "  # -- adicionados via review --"


def _planilha_padrao() -> Path | None:
    preferida = RAIZ / "planejamento-avancado-financeiro-icaro26.xlsx"
    if preferida.exists():
        return preferida
    candidatas = [
        p for p in RAIZ.glob("*.xlsx") if not p.name.startswith("~$")
    ]
    return candidatas[0] if len(candidatas) == 1 else None


def _faturas_classificadas(conn, tabela, competencia=None) -> list[Fatura]:
    """Le do banco e reclassifica com o merchants.yml atual.

    A categoria e sempre derivada da tabela na hora da leitura, e nunca lida
    de volta do banco: corrigir uma linha do merchants.yml passa a valer no
    comando seguinte, sem reimportar PDF nenhum.

    As colunas categoria/bucket gravadas sao uma denormalizacao para consulta
    em SQL, e por isso sao reescritas aqui - senao o banco passaria a dizer
    uma coisa e a ferramenta outra depois de qualquer edicao na tabela.
    """
    faturas = [classifica(f, tabela) for f in store.carrega(conn, competencia)]
    store.atualiza_classificacao(conn, faturas)
    return faturas


# --------------------------------------------------------------------------- #
# import
# --------------------------------------------------------------------------- #

def cmd_import(args) -> int:
    tabela = carrega_tabela(args.tabela)
    conn = store.conecta(args.db)

    caminhos: list[Path] = []
    for alvo in args.pdfs:
        p = Path(alvo)
        caminhos.extend(sorted(p.glob("*.pdf")) if p.is_dir() else [p])

    if not caminhos:
        print("nenhum PDF encontrado.", file=sys.stderr)
        return 1

    importadas = puladas = rejeitadas = 0
    for caminho in caminhos:
        try:
            raw = extrai(caminho)
            if not args.forcar and store.ja_importada(conn, raw.pdf_sha256):
                print(f"  -- {caminho.name}: ja importada (mesmo sha256), pulando")
                puladas += 1
                continue

            fatura = classifica(normaliza(raw), tabela)
            reconcilia(fatura)

            anterior = store.competencia_existente(conn, fatura.competencia)
            store.grava(conn, fatura)
            importadas += 1

            aviso = ""
            if anterior and Path(anterior).name != caminho.name:
                aviso = f"  (substituiu {Path(anterior).name})"
            n_pend = len(fatura.pendentes)
            pend = f", {n_pend} a classificar" if n_pend else ""
            print(
                f"  ok {caminho.name}: {fatura.competencia}, "
                f"{len(fatura.transacoes)} lanc., "
                f"total {format_brl(fatura.total_a_pagar)}{pend}{aviso}"
            )
        except ErroReconciliacao as e:
            print(f"  REJEITADA {e}", file=sys.stderr)
            rejeitadas += 1
        except ErroFatura as e:
            print(f"  ERRO {e}", file=sys.stderr)
            rejeitadas += 1

    print(
        f"\n{importadas} importada(s), {puladas} pulada(s), {rejeitadas} rejeitada(s)."
    )
    pendentes = _agrupa_pendentes(_faturas_classificadas(conn, tabela))
    if pendentes:
        print(
            f"{len(pendentes)} merchant(s) sem categoria. "
            f"Rode `python -m nubank review` antes do export."
        )
    return 1 if rejeitadas else 0


# --------------------------------------------------------------------------- #
# status
# --------------------------------------------------------------------------- #

def cmd_status(args) -> int:
    tabela = carrega_tabela(args.tabela)
    conn = store.conecta(args.db)
    faturas = _faturas_classificadas(conn, tabela)
    if not faturas:
        print("nenhuma fatura importada. Rode `python -m nubank import faturas/`.")
        return 0

    cab = f"{'comp':8} {'total':>12} {'parcelas':>11} {'essencial':>11} {'superfluo':>11} {'ajuste':>11} {'pend':>5}"
    print(cab)
    print("-" * len(cab))
    for f in faturas:
        n_pend = len(f.pendentes)
        print(
            f"{f.competencia:8} {format_brl(f.total_a_pagar):>12} "
            f"{format_brl(f.total_bucket(Bucket.PARCELAS)):>11} "
            f"{format_brl(f.total_bucket(Bucket.ESSENCIAL)):>11} "
            f"{format_brl(f.total_bucket(Bucket.SUPERFLUO)):>11} "
            f"{format_brl(f.ajuste_saldo):>11} {n_pend or '-':>5}"
        )

    total_pend = sum(len(f.pendentes) for f in faturas)
    if total_pend:
        merchants = _agrupa_pendentes(faturas)
        print(
            f"\n{total_pend} lancamento(s) em {len(merchants)} merchant(s) sem "
            f"categoria: {format_brl(sum(v for _, v, _ in merchants.values()))}"
        )
    return 0


# --------------------------------------------------------------------------- #
# review
# --------------------------------------------------------------------------- #

def _categorias_escolhiveis(tabela) -> list[str]:
    """As categorias que fazem sentido escolher a mao no review."""
    return sorted(
        c
        for c, b in tabela.categorias.items()
        if b is not Bucket.NAO_GASTO and c not in CATEGORIAS_ESTRUTURAIS
    )


def _limpa(descricao: str) -> str:
    """Tira o prefixo de cartao da descricao exibida.

    O '••••' vira lixo no console do Windows e o final do cartao ja esta
    guardado em campo proprio.
    """
    return re.sub(r"^•+\s*\d{4}\s*", "", descricao).strip()


def _agrupa_pendentes(faturas: list[Fatura]) -> dict[str, tuple[int, Decimal, str]]:
    """merchant -> (quantidade, total, exemplo de descricao crua)."""
    agrupado: dict[str, list] = defaultdict(lambda: [0, Decimal("0.00"), ""])
    for f in faturas:
        for t in f.pendentes:
            item = agrupado[t.merchant_norm]
            item[0] += 1
            item[1] += t.valor
            item[2] = item[2] or t.descricao_raw
    return {k: tuple(v) for k, v in agrupado.items()}


def _insere_merchants(caminho: Path, novos: dict[str, str]) -> None:
    """Acrescenta entradas no bloco `merchants:` preservando comentarios.

    yaml.dump reescreveria o arquivo inteiro e apagaria os comentarios, que sao
    metade do valor dessa tabela.
    """
    linhas = caminho.read_text(encoding="utf-8").splitlines()
    try:
        inicio = next(i for i, l in enumerate(linhas) if l.startswith("merchants:"))
    except StopIteration:
        raise ErroFatura(f"{caminho}: bloco 'merchants:' nao encontrado.") from None

    fim = len(linhas)
    for i in range(inicio + 1, len(linhas)):
        linha = linhas[i]
        if linha.strip() and not linha[0].isspace() and not linha.startswith("#"):
            fim = i
            break

    # Um review por mes criaria um cabecalho por mes se ele fosse escrito de
    # novo toda vez. Como a insercao e sempre no fim do bloco, o marcador que
    # ja existe e o ultimo, e basta escrever as entradas embaixo dele.
    ja_marcado = MARCADOR_REVIEW in linhas[inicio:fim]
    bloco = [] if ja_marcado else ["", MARCADOR_REVIEW]
    bloco += [f'  "{m}": {c}' for m, c in sorted(novos.items())]

    while fim > inicio + 1 and not linhas[fim - 1].strip():
        fim -= 1
    linhas[fim:fim] = bloco
    caminho.write_text("\n".join(linhas) + "\n", encoding="utf-8")


def _nova_categoria(caminho: Path, nome: str, bucket: str) -> None:
    linhas = caminho.read_text(encoding="utf-8").splitlines()
    inicio = next(i for i, l in enumerate(linhas) if l.startswith("categorias:"))
    fim = next(
        (
            i
            for i in range(inicio + 1, len(linhas))
            if linhas[i].strip() and not linhas[i][0].isspace()
        ),
        len(linhas),
    )
    while fim > inicio + 1 and not linhas[fim - 1].strip():
        fim -= 1
    linhas.insert(fim, f"  {nome}: {bucket}")
    caminho.write_text("\n".join(linhas) + "\n", encoding="utf-8")


def cmd_review(args) -> int:
    caminho_tabela = Path(args.tabela) if args.tabela else None
    tabela = carrega_tabela(caminho_tabela)
    arquivo_tabela = caminho_tabela or (RAIZ / "merchants.yml")

    conn = store.conecta(args.db)
    faturas = _faturas_classificadas(conn, tabela)
    pendentes = _agrupa_pendentes(faturas)

    if not pendentes:
        print("nada a revisar: todos os merchants ja tem categoria.")
        return 0

    ordenados = sorted(pendentes.items(), key=lambda kv: -abs(kv[1][1]))
    print(f"{len(ordenados)} merchant(s) sem categoria, do maior valor para o menor.")
    print("Digite o numero da categoria, ENTER para pular, 'n' para criar uma "
          "categoria nova, 'q' para salvar e sair.\n")

    novos: dict[str, str] = {}
    categorias = _categorias_escolhiveis(tabela)

    for i, (merchant, (qtd, total, exemplo)) in enumerate(ordenados, 1):
        print(f"[{i}/{len(ordenados)}] {merchant}")
        print(f"     {qtd} lanc., {format_brl(total)}   ex.: {_limpa(exemplo)}")
        menu = "  ".join(
            f"{n}={c}({tabela.categorias[c].value[:4]})"
            for n, c in enumerate(categorias, 1)
        )
        print(f"     {menu}")
        try:
            resposta = input("  > ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if resposta == "q":
            break
        if not resposta:
            continue
        if resposta == "n":
            nome = input("    nome da categoria: ").strip().lower()
            bucket = input("    bucket (parcelas/essencial/superfluo): ").strip().lower()
            if not nome or bucket not in ("parcelas", "essencial", "superfluo"):
                print("    ignorado.")
                continue
            _nova_categoria(arquivo_tabela, nome, bucket)
            tabela = carrega_tabela(caminho_tabela)
            categorias = _categorias_escolhiveis(tabela)
            novos[merchant] = nome
            print(f"    {merchant} -> {nome}")
            continue
        if resposta.isdigit() and 1 <= int(resposta) <= len(categorias):
            escolhida = categorias[int(resposta) - 1]
            novos[merchant] = escolhida
            print(f"    {merchant} -> {escolhida}")
        else:
            print("    nao entendi, pulando.")

    if not novos:
        print("\nnada foi classificado.")
        return 0

    _insere_merchants(arquivo_tabela, novos)
    tabela = carrega_tabela(caminho_tabela)
    faturas = _faturas_classificadas(conn, tabela)
    n = store.atualiza_classificacao(conn, faturas)
    restantes = _agrupa_pendentes(faturas)

    print(f"\n{len(novos)} merchant(s) gravado(s) em {arquivo_tabela.name}.")
    print(f"{n} lancamento(s) reclassificado(s).")
    if restantes:
        print(f"ainda faltam {len(restantes)} merchant(s).")
    return 0


# --------------------------------------------------------------------------- #
# export
# --------------------------------------------------------------------------- #

def cmd_export(args) -> int:
    tabela = carrega_tabela(args.tabela)
    conn = store.conecta(args.db)
    faturas = _faturas_classificadas(conn, tabela, args.competencia)

    if not faturas:
        alvo = f" para {args.competencia}" if args.competencia else ""
        print(f"nenhuma fatura importada{alvo}.", file=sys.stderr)
        return 1

    pendentes = _agrupa_pendentes(faturas)
    if pendentes and not args.ignorar_pendentes:
        total = sum(v for _, v, _ in pendentes.values())
        print(
            f"export bloqueado: {len(pendentes)} merchant(s) sem categoria, "
            f"somando {format_brl(total)}.",
            file=sys.stderr,
        )
        print("\nOs maiores:", file=sys.stderr)
        for m, (qtd, valor, _) in sorted(
            pendentes.items(), key=lambda kv: -abs(kv[1][1])
        )[:10]:
            print(f"  {format_brl(valor):>12}  {qtd:>3}x  {m}", file=sys.stderr)
        print(
            "\nRode `python -m nubank review` para classificar. "
            "Se quiser exportar assim mesmo, use --ignorar-pendentes "
            "(o valor nao classificado vai cair na coluna F junto com o "
            "ajuste de saldo, e voce perde a leitura dessa coluna).",
            file=sys.stderr,
        )
        return 1

    planilha = Path(args.planilha) if args.planilha else _planilha_padrao()
    if planilha is None:
        print(
            "nao consegui achar a planilha. Passe --planilha caminho.xlsx",
            file=sys.stderr,
        )
        return 1

    if args.csv:
        destino = exporta_csv(args.csv, faturas)
        print(f"CSV de detalhe: {destino}")

    if not args.apply:
        mudancas = [m for m in planeja(planilha, faturas) if m.mudou]
        if not mudancas:
            print(f"{planilha.name}: nada a mudar, ja esta atualizada.")
            return 0
        print(f"dry-run em {planilha.name} - {len(mudancas)} celula(s) mudariam:\n")
        for m in mudancas:
            print(f"  {m}")
        print("\nnada foi escrito. Repita com --apply para aplicar.")
        return 0

    backup, mudancas = aplica(planilha, faturas)
    print(f"backup: {backup}")
    print(f"{len(mudancas)} celula(s) escrita(s) em {planilha.name}:\n")
    for m in mudancas:
        print(f"  {m}")
    return 0


# --------------------------------------------------------------------------- #
# conferir
# --------------------------------------------------------------------------- #

def _mostra_valores(planilha: Path) -> None:
    """Imprime o que a planilha deve exibir, para comparar com a tela.

    Se a tela mostrar outro numero com as formulas intactas, o Excel esta
    exibindo a copia que ele tem em memoria desde antes de a ferramenta
    escrever - e salvar por cima desfaz tudo.
    """
    esperados = valores_esperados(planilha)
    gastos = [v for v in esperados if v.celula.startswith("E")]
    fatura = {v.mes: v for v in esperados if v.celula.startswith("F")}

    print("Com essas formulas, a aba Controle Mensal tem de mostrar:\n")
    print(f"  {'mes':8} {'E gastos fixos':>16} {'F fatura cartao':>17}")
    print("  " + "-" * 43)
    for v in gastos:
        f = fatura.get(v.mes)
        valor_f = format_brl(Decimal(str(f.valor))) if f else "-"
        print(
            f"  {v.mes:8} {format_brl(Decimal(str(v.valor))):>16} {valor_f:>17}"
        )
    print(
        "\nSe a sua tela mostrar outro numero, o Excel esta exibindo a copia "
        "dele\nem memoria, nao o arquivo. Feche SEM SALVAR e abra de novo - "
        "salvar\npor cima desfaz o conserto."
    )


def cmd_conferir(args) -> int:
    """Verifica as formulas que ligam as abas, e opcionalmente repara."""
    planilha = Path(args.planilha) if args.planilha else _planilha_padrao()
    if planilha is None:
        print("nao achei a planilha. Passe --planilha caminho.xlsx", file=sys.stderr)
        return 1

    problemas = confere(planilha)
    if not problemas:
        print(f"{planilha.name}: todas as formulas entre abas estao no lugar.\n")
        _mostra_valores(planilha)
        return 0

    print(f"{planilha.name}: {len(problemas)} formula(s) quebrada(s).\n")
    for p in problemas:
        print(f"  {p}")

    if not args.reparar:
        print(
            "\nEssas celulas nao estao puxando o valor da outra aba, e o Excel "
            "nao acusa erro nisso: a conta segue com o numero errado.\n"
            "Para consertar:  python -m nubank conferir --reparar"
        )
        return 1

    backup = faz_backup(planilha)
    n = repara(planilha, problemas)
    print(f"\nbackup: {backup}")
    print(f"{n} formula(s) restaurada(s).")

    restantes = confere(planilha)
    if restantes:
        print(f"ATENCAO: {len(restantes)} ainda quebrada(s).", file=sys.stderr)
        return 1
    print("conferido de novo: tudo no lugar.")
    return 0


# --------------------------------------------------------------------------- #
# parcelas
# --------------------------------------------------------------------------- #

def cmd_parcelas(args) -> int:
    """Projeta as parcelas que ainda vao cair nos proximos meses.

    E a divida ja contratada que nao aparece em coluna nenhuma da planilha, e a
    explicacao mais provavel para a fatura oscilar entre 2.500 e 4.000.
    """
    tabela = carrega_tabela(args.tabela)
    conn = store.conecta(args.db)
    faturas = _faturas_classificadas(conn, tabela)
    if not faturas:
        print("nenhuma fatura importada.")
        return 0

    ultima = faturas[-1]
    ativos = [
        t
        for t in ultima.transacoes
        if t.tipo is TipoLancamento.PARCELA
        and t.parcela_atual is not None
        and t.parcela_atual < t.parcela_total
    ]
    if not ativos:
        print(f"nenhum parcelamento em aberto na fatura {ultima.competencia}.")
        return 0

    print(f"parcelamentos em aberto na fatura {ultima.competencia}:\n")
    futuro: dict[str, Decimal] = defaultdict(lambda: Decimal("0.00"))
    ano, mes = (int(p) for p in ultima.competencia.split("-"))
    total_restante = Decimal("0.00")

    for t in sorted(ativos, key=lambda t: -t.valor):
        restantes = t.parcela_total - t.parcela_atual
        soma = t.valor * restantes
        total_restante += soma
        print(
            f"  {format_brl(t.valor):>11} x {restantes:>2} restante(s)  "
            f"({t.parcela_atual}/{t.parcela_total})  {t.merchant_norm}"
        )
        for k in range(1, restantes + 1):
            m = mes + k
            futuro[f"{ano + (m - 1) // 12:04d}-{(m - 1) % 12 + 1:02d}"] += t.valor

    print(f"\n  total ja contratado: {format_brl(total_restante)}\n")
    print("por competencia futura:")
    for comp in sorted(futuro):
        print(f"  {comp}  {format_brl(futuro[comp])}")
    return 0


# --------------------------------------------------------------------------- #

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python -m nubank",
        description="Le faturas do cartao Nubank e alimenta a aba Cartao da planilha.",
    )
    p.add_argument("--db", help="caminho do SQLite (padrao: faturas.db na raiz)")
    p.add_argument("--tabela", help="caminho do merchants.yml")
    sub = p.add_subparsers(dest="comando", required=True)

    s = sub.add_parser("import", help="le PDFs, reconcilia e grava")
    s.add_argument("pdfs", nargs="+", help="arquivos .pdf ou uma pasta")
    s.add_argument(
        "--forcar", action="store_true", help="reimporta mesmo se o sha256 ja existir"
    )
    s.set_defaults(func=cmd_import)

    s = sub.add_parser("status", help="resumo do que ja foi importado")
    s.set_defaults(func=cmd_status)

    s = sub.add_parser("review", help="classifica os merchants pendentes")
    s.set_defaults(func=cmd_review)

    s = sub.add_parser("export", help="escreve na aba Cartao (dry-run por padrao)")
    s.add_argument("--apply", action="store_true", help="escreve de verdade")
    s.add_argument("--planilha", help="caminho do .xlsx")
    s.add_argument("--competencia", help="so uma competencia, ex: 2026-08")
    s.add_argument("--csv", help="tambem grava um CSV de detalhe nesse caminho")
    s.add_argument(
        "--ignorar-pendentes",
        dest="ignorar_pendentes",
        action="store_true",
        help="exporta mesmo com merchants sem categoria",
    )
    s.set_defaults(func=cmd_export)

    s = sub.add_parser(
        "conferir", help="verifica as formulas que ligam as abas da planilha"
    )
    s.add_argument("--planilha", help="caminho do .xlsx")
    s.add_argument(
        "--reparar", action="store_true", help="reescreve as formulas quebradas"
    )
    s.set_defaults(func=cmd_conferir)

    s = sub.add_parser("parcelas", help="projeta as parcelas futuras ja contratadas")
    s.set_defaults(func=cmd_parcelas)

    return p


def _console_tolerante_a_unicode() -> None:
    """Descricao de merchant tem acento, e o PDF usa '•' e o MINUS SIGN.

    No console do Windows em cp1252 isso levanta UnicodeEncodeError no meio de
    um print e derruba o comando. Trocar o caractere por '?' e melhor do que
    perder o import inteiro por causa de um bullet.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except (ValueError, OSError):
                pass


def main(argv: list[str] | None = None) -> int:
    _console_tolerante_a_unicode()
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except ErroFatura as e:
        print(f"erro: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
