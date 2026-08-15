"""O review edita o merchants.yml como texto, nao com yaml.dump.

yaml.dump reescreveria o arquivo inteiro e apagaria os comentarios, que sao
metade do valor dessa tabela. Em troca disso, a edicao textual precisa destes
testes para nao corromper o arquivo com o uso.
"""

from __future__ import annotations

import yaml

from nubank.cli import MARCADOR_REVIEW, _insere_merchants, _nova_categoria

BASE = """\
# comentario do topo que precisa sobreviver
categorias:
  mercado: essencial
  compras: superfluo

padroes:
  - prefixo: "ifd*"
    categoria: compras

merchants:
  # -- transporte --
  "uber - nupay": mercado
"""


def _escreve(tmp_path, conteudo=BASE):
    caminho = tmp_path / "merchants.yml"
    caminho.write_text(conteudo, encoding="utf-8")
    return caminho


def test_insere_e_mantem_os_comentarios(tmp_path):
    caminho = _escreve(tmp_path)
    _insere_merchants(caminho, {"padaria do ze": "compras"})

    texto = caminho.read_text(encoding="utf-8")
    assert "# comentario do topo que precisa sobreviver" in texto
    assert "# -- transporte --" in texto

    dados = yaml.safe_load(texto)
    assert dados["merchants"]["padaria do ze"] == "compras"
    assert dados["merchants"]["uber - nupay"] == "mercado"


def test_reviews_seguidos_nao_acumulam_cabecalho(tmp_path):
    """Um review por mes criaria um cabecalho por mes."""
    caminho = _escreve(tmp_path)
    _insere_merchants(caminho, {"loja a": "compras"})
    _insere_merchants(caminho, {"loja b": "mercado"})
    _insere_merchants(caminho, {"loja c": "compras"})

    texto = caminho.read_text(encoding="utf-8")
    assert texto.count(MARCADOR_REVIEW) == 1

    dados = yaml.safe_load(texto)
    for merchant in ("loja a", "loja b", "loja c"):
        assert merchant in dados["merchants"]


def test_padroes_nao_sao_engolidos(tmp_path):
    """A insercao vai no fim do bloco merchants, que aqui e o ultimo do arquivo."""
    caminho = _escreve(tmp_path)
    _insere_merchants(caminho, {"loja nova": "compras"})

    dados = yaml.safe_load(caminho.read_text(encoding="utf-8"))
    assert dados["padroes"] == [{"prefixo": "ifd*", "categoria": "compras"}]
    assert dados["categorias"] == {"mercado": "essencial", "compras": "superfluo"}


def test_categoria_nova_entra_no_bloco_certo(tmp_path):
    caminho = _escreve(tmp_path)
    _nova_categoria(caminho, "transferencia", "superfluo")
    _insere_merchants(caminho, {"fulano da silva": "transferencia"})

    dados = yaml.safe_load(caminho.read_text(encoding="utf-8"))
    assert dados["categorias"]["transferencia"] == "superfluo"
    assert dados["merchants"]["fulano da silva"] == "transferencia"
    # o bloco de categorias nao pode ter comido o de padroes
    assert dados["padroes"][0]["prefixo"] == "ifd*"


def test_arquivo_continua_yaml_valido_apos_varias_edicoes(tmp_path):
    caminho = _escreve(tmp_path)
    _nova_categoria(caminho, "saude", "essencial")
    _insere_merchants(caminho, {"drogaria x": "saude"})
    _nova_categoria(caminho, "pet", "essencial")
    _insere_merchants(caminho, {"petshop y": "pet", "loja z": "compras"})

    dados = yaml.safe_load(caminho.read_text(encoding="utf-8"))
    assert len(dados["categorias"]) == 4
    assert len(dados["merchants"]) == 4
    assert dados["merchants"]["petshop y"] == "pet"
