"""Edicao cirurgica de .xlsx: altera as celulas pedidas e mais nada.

Por que nao usar openpyxl para escrever: o `save()` dele reescreve o arquivo
inteiro, e nesta planilha isso perde os atributos apply* dos estilos - 50
`applyNumberFormat`, 68 `applyFont`, 54 `applyBorder`, 7 `applyFill`. Sem
`applyNumberFormat="1"`, o Excel ignora o `numFmtId` da celula e herda o
formato do estilo nomeado, que aqui e `General`. Resultado visivel: a coluna de
mes mostra `46023` em vez de `jan/26`, e os valores em reais viram numero cru.

O openpyxl nao enxerga o proprio estrago porque le o `numFmtId` direto e ignora
a flag - conferir a saida dele com ele mesmo da sempre verde.

Aqui o arquivo e tratado como o zip que ele e: le todas as partes, troca so o
XML da aba alvo, e grava as outras byte a byte. O que nao foi pedido nao muda,
e isso e verificavel comparando as partes do zip.

Leitura continua com openpyxl, que nao danifica nada.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from .errors import ErroExport

CAMINHO_WORKBOOK = "xl/workbook.xml"
CAMINHO_RELS = "xl/_rels/workbook.xml.rels"
FIM_ROW = "</row>"


def indice_coluna(letras: str) -> int:
    """'A' -> 1, 'H' -> 8, 'AA' -> 27."""
    n = 0
    for ch in letras.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parte_ref(ref: str) -> tuple[str, int]:
    """'H16' -> ('H', 16)."""
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"referencia de celula invalida: {ref!r}")
    return m.group(1).upper(), int(m.group(2))


def atributos(tag: str) -> dict[str, str]:
    """Atributos de uma tag XML, por nome.

    A ordem nao e garantida e varia com quem gravou o arquivo: o Google Sheets
    escreve `<row r="1" ht="18.75">`, o Excel escreve
    `<row x14ac:dyDescent="0.25" r="1" ht="18.75">`. Procurar por `<row r="`
    funciona num e falha silenciosamente no outro.
    """
    return dict(re.findall(r'([\w:.-]+)="([^"]*)"', tag))


class Aba:
    """O XML de uma aba, com operacoes de celula sobre o texto cru."""

    def __init__(self, xml: str, nome: str) -> None:
        self.xml = xml
        self.nome = nome

    # -- leitura ----------------------------------------------------------- #

    def _localiza(self, ref: str) -> tuple[int, int, str, str | None] | None:
        """(inicio, fim, tag_de_abertura, corpo) da celula, ou None.

        Varre as tags e compara o atributo `r` ja parseado, em vez de procurar
        pelo texto `<c r="REF"`: a ordem dos atributos depende de quem gravou o
        arquivo, e assumir uma ordem falha calada.
        """
        for m in re.finditer(r"<c\b[^>]*>", self.xml):
            tag = m.group(0)
            if atributos(tag).get("r") != ref:
                continue
            if tag.endswith("/>"):
                return m.start(), m.end(), tag, None
            fecha = self.xml.index("</c>", m.end())
            return m.start(), fecha + 4, tag, self.xml[m.end() : fecha]
        return None

    def _localiza_linha(self, numero: int) -> tuple[int, int, str] | None:
        """(inicio, fim, xml_inteiro) da <row>, ou None."""
        for m in re.finditer(r"<row\b[^>]*>", self.xml):
            tag = m.group(0)
            if atributos(tag).get("r") != str(numero):
                continue
            if tag.endswith("/>"):
                return m.start(), m.end(), tag
            fecha = self.xml.index("</row>", m.end())
            return m.start(), fecha + 6, self.xml[m.start() : fecha + 6]
        return None

    def existe(self, ref: str) -> bool:
        return self._localiza(ref) is not None

    def estilo(self, ref: str) -> str | None:
        """O indice de estilo (`s="..."`) da celula, se houver."""
        achado = self._localiza(ref)
        if not achado:
            return None
        m = re.search(r'\ss="(\d+)"', achado[2])
        return m.group(1) if m else None

    def formula(self, ref: str) -> str | None:
        achado = self._localiza(ref)
        if not achado or achado[3] is None:
            return None
        m = re.search(r"<f[^>]*>(.*?)</f>", achado[3], re.S)
        return m.group(1) if m else None

    # -- escrita ----------------------------------------------------------- #

    def _substitui(self, ref: str, elemento: str) -> None:
        achado = self._localiza(ref)
        if achado:
            inicio, fim, _, _ = achado
            self.xml = self.xml[:inicio] + elemento + self.xml[fim:]
        else:
            self._insere_em_ordem(ref, elemento)

    def _insere_em_ordem(self, ref: str, elemento: str) -> None:
        """Insere a celula na posicao correta dentro da sua <row>.

        O Excel exige as celulas em ordem de coluna dentro da linha; jogar no
        fim funciona por acidente ate a primeira linha em que nao funciona.
        """
        coluna, linha = parte_ref(ref)
        alvo = indice_coluna(coluna)

        achado = self._localiza_linha(linha)
        if not achado:
            raise ErroExport(
                f"aba '{self.nome}': linha {linha} nao existe, nao sei onde por {ref}."
            )
        inicio, fim, row_xml = achado

        if row_xml.endswith("/>"):  # linha vazia e auto-fechada
            nova = row_xml[:-2] + ">" + elemento + FIM_ROW
            self.xml = self.xml[:inicio] + nova + self.xml[fim:]
            return

        posicao = None
        for c in re.finditer(r"<c\b[^>]*>", row_xml):
            ref_atual = atributos(c.group(0)).get("r", "")
            m_col = re.match(r"([A-Z]+)", ref_atual)
            if m_col and indice_coluna(m_col.group(1)) > alvo:
                posicao = c.start()
                break
        if posicao is None:
            posicao = row_xml.rindex(FIM_ROW)

        nova = row_xml[:posicao] + elemento + row_xml[posicao:]
        self.xml = self.xml[:inicio] + nova + self.xml[fim:]

    def _monta(self, ref: str, corpo: str, tipo: str | None, estilo: str | None) -> str:
        estilo = estilo if estilo is not None else self.estilo(ref)
        attrs = f' r="{ref}"'
        if estilo is not None:
            attrs += f' s="{estilo}"'
        if tipo:
            attrs += f' t="{tipo}"'
        return f"<c{attrs}>{corpo}</c>"

    def define_numero(self, ref: str, valor: float, estilo: str | None = None) -> None:
        self._substitui(ref, self._monta(ref, f"<v>{valor:.2f}</v>", None, estilo))

    def define_texto(self, ref: str, texto: str, estilo: str | None = None) -> None:
        """String inline, para nao precisar mexer em sharedStrings.xml."""
        corpo = f'<is><t xml:space="preserve">{escape(texto)}</t></is>'
        self._substitui(ref, self._monta(ref, corpo, "inlineStr", estilo))

    def define_formula(
        self,
        ref: str,
        formula: str,
        estilo: str | None = None,
        valor: float | None = None,
    ) -> None:
        """Escreve a formula e, quando dado, o valor em cache dela.

        O cache importa: quem abre o arquivo sem avaliar formula nenhuma - o
        preview do editor, um visualizador leve, o pandas - mostra o `<v>` e so.
        Sem ele, essas ferramentas tentam avaliar `INDEX(...)` por conta propria
        e exibem lixo ou erro, ainda que o Excel abra a mesma planilha certa.

        Nao ha risco de o cache ficar velho: `recalcular_ao_abrir` marca
        fullCalcOnLoad, entao o Excel recalcula tudo e sobrescreve.

        Se a celula for a mestra de uma formula compartilhada, os atributos
        `t="shared"`, `ref` e `si` sao preservados: as celulas irmas derivam do
        texto novo sozinhas, deslocando as referencias por linha.
        """
        achado = self._localiza(ref)
        compartilhada = ""
        if achado and achado[3]:
            m = re.search(r"<f([^>]*)>", achado[3])
            if m and 't="shared"' in m.group(1):
                compartilhada = m.group(1)
        # O conteudo de <f> no OOXML nao leva o '=' inicial; quem usa isso e a
        # barra de formulas do Excel. Com o '=' o arquivo abre corrompido.
        corpo = f"<f{compartilhada}>{escape(formula.lstrip('='))}</f>"
        if valor is not None:
            corpo += f"<v>{valor:.2f}</v>"
        self._substitui(ref, self._monta(ref, corpo, None, estilo))

    def largura_coluna(self, coluna: str, largura: float) -> None:
        idx = indice_coluna(coluna)
        novo = f'<col customWidth="1" min="{idx}" max="{idx}" width="{largura}"/>'

        m_cols = re.search(r"<cols>.*?</cols>", self.xml, re.S)
        if not m_cols:
            # Sem bloco <cols>, ele entra logo antes de <sheetData>.
            m_dados = re.search(r"<sheetData[ >]", self.xml)
            if not m_dados:
                raise ErroExport(f"aba '{self.nome}': <sheetData> nao encontrado.")
            self.xml = (
                self.xml[: m_dados.start()]
                + f"<cols>{novo}</cols>"
                + self.xml[m_dados.start() :]
            )
            return

        cols = m_cols.group(0)

        # Um <col> costuma cobrir um intervalo ('min=8 max=26'). Deixar dois
        # <col> sobrepostos e arquivo invalido, entao o intervalo que contem a
        # coluna alvo e partido em ate tres: o pedaco antes, ela sozinha com a
        # largura nova, e o pedaco depois.
        for atual in re.finditer(r"<col[^>]*/>", cols):
            tag = atual.group(0)
            mn = re.search(r'min="(\d+)"', tag)
            mx = re.search(r'max="(\d+)"', tag)
            if not mn or not mx:
                continue
            minimo, maximo = int(mn.group(1)), int(mx.group(1))
            if not (minimo <= idx <= maximo):
                continue

            atual_largura = re.search(r'width="([\d.]+)"', tag)
            if minimo == maximo == idx:
                if atual_largura and float(atual_largura.group(1)) >= largura:
                    return
                substituto = novo
            else:
                pedacos = []
                if minimo < idx:
                    pedacos.append(
                        re.sub(r'max="\d+"', f'max="{idx - 1}"', tag)
                    )
                pedacos.append(novo)
                if idx < maximo:
                    pedacos.append(
                        re.sub(r'min="\d+"', f'min="{idx + 1}"', tag)
                    )
                substituto = "".join(pedacos)

            novos_cols = cols.replace(tag, substituto, 1)
            self.xml = (
                self.xml[: m_cols.start()] + novos_cols + self.xml[m_cols.end() :]
            )
            return

        novos_cols = cols.replace("</cols>", novo + "</cols>")
        self.xml = self.xml[: m_cols.start()] + novos_cols + self.xml[m_cols.end() :]


class Planilha:
    """O .xlsx como zip. So as partes explicitamente trocadas sao reescritas."""

    def __init__(self, caminho: str | Path) -> None:
        self.caminho = Path(caminho)
        if not self.caminho.exists():
            raise ErroExport(f"planilha nao encontrada: {self.caminho}")
        with zipfile.ZipFile(self.caminho) as z:
            self.ordem = z.namelist()
            self.partes = {nome: z.read(nome) for nome in self.ordem}
        self._abas: dict[str, Aba] = {}

    def _caminho_da_aba(self, nome: str) -> str:
        wb = self.partes[CAMINHO_WORKBOOK].decode("utf-8")
        rels = self.partes[CAMINHO_RELS].decode("utf-8")

        rid = None
        for tag in re.findall(r"<sheet\b[^>]*/?>", wb):
            attrs = atributos(tag)
            if attrs.get("name") == nome:
                rid = attrs.get("r:id") or attrs.get("id")
                break
        if rid is None:
            raise ErroExport(f"a planilha nao tem a aba '{nome}'.")

        for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
            attrs = atributos(tag)
            if attrs.get("Id") == rid:
                alvo = attrs.get("Target", "").lstrip("/")
                return alvo if alvo.startswith("xl/") else f"xl/{alvo}"
        raise ErroExport(f"relacionamento da aba '{nome}' nao encontrado.")

    def aba(self, nome: str) -> Aba:
        if nome not in self._abas:
            caminho = self._caminho_da_aba(nome)
            self._abas[nome] = Aba(self.partes[caminho].decode("utf-8"), nome)
            self._abas[nome]._parte = caminho  # type: ignore[attr-defined]
        return self._abas[nome]

    def recalcular_ao_abrir(self) -> None:
        """Marca fullCalcOnLoad.

        As formulas guardam o ultimo valor calculado. Trocamos os numeros de que
        elas dependem, entao sem isso o Excel mostraria o cache velho ate alguem
        editar alguma celula.
        """
        wb = self.partes[CAMINHO_WORKBOOK].decode("utf-8")
        if "fullCalcOnLoad" in wb:
            return
        if "<calcPr" in wb:
            wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, 1)
        else:
            wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self.partes[CAMINHO_WORKBOOK] = wb.encode("utf-8")

    def partes_alteradas(self, original: Planilha) -> list[str]:
        return [n for n in self.ordem if self.partes[n] != original.partes.get(n)]

    def salva(self, destino: str | Path | None = None) -> Path:
        for aba in self._abas.values():
            self.partes[aba._parte] = aba.xml.encode("utf-8")  # type: ignore[attr-defined]

        destino = Path(destino) if destino else self.caminho
        temporario = destino.with_suffix(destino.suffix + ".tmp")
        with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED) as z:
            for nome in self.ordem:
                z.writestr(nome, self.partes[nome])
        shutil.move(str(temporario), str(destino))
        return destino
