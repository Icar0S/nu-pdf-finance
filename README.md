# nu-pdf-finance

Le as faturas do cartao Nubank em PDF e preenche a aba **Cartao** da planilha
`planejamento-avancado-financeiro-icaro26.xlsx`.

Nem os PDFs nem a planilha nem o banco vao para o git.

## Instalacao

```bash
pip install -r requirements.txt
```

## Uso normal (uma vez por mes)

```bash
python -m nubank import faturas/          # le, reconcilia e grava
python -m nubank review                   # classifica merchants novos
python -m nubank export                   # dry-run: mostra o que mudaria
python -m nubank export --apply           # escreve na planilha
```

Outros comandos:

```bash
python -m nubank status                   # resumo do que ja foi importado
python -m nubank parcelas                 # parcelas futuras ja contratadas
python -m nubank export --csv detalhe.csv # CSV linha a linha
```

## Como funciona

```
extract  ->  normalize  ->  reconcile  ->  classify  ->  export
  PDF        Decimal,        as contas     categoria    aba Cartao
             datas,          fecham?       e bucket     + backup
             parcelas
```

Os quatro primeiros estagios sao funcoes puras. So o `export` toca disco.

### A fatura se verifica sozinha

Cada PDF traz o resumo e a lista de transacoes: dois caminhos independentes
para os mesmos numeros. O `reconcile` confere tres somas:

```
fatura anterior - pagamentos + compras + IOF + outros = total a pagar
soma das transacoes do periodo                        = compras + IOF + outros
soma dos pagamentos listados                          = pagamento recebido
```

Se qualquer uma nao fechar (tolerancia de R$ 0,02, porque o proprio Nubank
arredonda um centavo em metade das faturas), o import e **rejeitado inteiro**.
Nao existe `--force`. E o que permite importar oito meses de historico sem
reconferir linha a linha.

### Classificacao

Tres camadas, todas deterministicas — rodar o mesmo PDF duas vezes da sempre o
mesmo resultado:

1. **Regras estruturais** — parcelamento, IOF, estorno e pagamento saem da
   forma da linha, nao do nome do merchant. Uma compra em 6x vai para
   "Parcelas e assinaturas" nos 6 meses, independente da loja. O IOF de
   `Render.Com` entra no mesmo balde que o `Render.Com` que o gerou.
2. **Tabela de merchants** (`merchants.yml`) — match exato na descricao
   normalizada. Gasto de pessoa fisica e repetitivo, entao isso cobre quase
   tudo.
3. **Padroes de agregador** — `ifd*` vira delivery, `shopee*` vira compras. Um
   restaurante inedito no iFood ja nasce classificado.

O que nao casa em nenhuma camada fica **pendente e trava o export**. Um
merchant desconhecido caindo em algum balde por padrao seria exatamente o erro
silencioso que a planilha nao te deixaria enxergar.

### A coluna F da planilha

`Fatura total` (coluna B) e "quanto voce gastou no periodo" **nao sao a mesma
coisa**. B inclui o saldo que veio da fatura anterior menos o que voce pagou.
C+D+E cobrem so os lancamentos do periodo. A diferenca e sempre:

```
coluna F  =  fatura anterior - pagamentos
```

Por isso F fica em `-49,99` no mes em que voce pagou R$ 50 a mais: nao e gasto
perdido nem bug, e ajuste de saldo. A coluna G (Observacoes) escreve isso em
cada linha para voce nao caçar um bug que nao existe.

### Escrita na planilha

- `export` sem `--apply` e dry-run: mostra celula a celula o que mudaria e nao
  escreve nada.
- `export --apply` copia a planilha para `backups/` com carimbo de data e hora
  **antes** de escrever.
- A coluna F nunca e escrita: ela e formula.
- Rodar o export duas vezes seguidas nao muda nada na segunda.

## Arquivos

| Arquivo | O que e |
|---|---|
| `merchants.yml` | Merchant -> categoria -> coluna. Versionado: o diff mostra toda mudanca de criterio. |
| `faturas.db` | SQLite. Dedup por sha256 do PDF e permite reclassificar o historico sem reparsear. |
| `backups/` | Copias da planilha, uma por `export --apply`. |

## Testes

```bash
python -m pytest
```

Os testes de `tests/test_corpus.py` rodam sobre os PDFs reais e pulam sozinhos
se a pasta `faturas/` estiver vazia (ela nao vai para o git). Sao eles que
pegam mudanca de layout entre uma fatura de janeiro e uma de agosto.
