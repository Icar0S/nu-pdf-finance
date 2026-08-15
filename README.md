# nu-pdf-finance

Le as faturas do cartao Nubank em PDF e preenche a aba **Cartao** da planilha
`planejamento-avancado-financeiro-icaro26.xlsx`.

Nem os PDFs nem a planilha nem o banco vao para o git.

## Uso normal (uma vez por mes)

Coloque os PDFs novos em `faturas/` e rode:

```powershell
.\executar.ps1            # prepara tudo, importa, revisa e mostra o dry-run
.\executar.ps1 -Apply     # idem, e escreve na planilha apos confirmar
```

O script cuida da sequencia inteira: cria o venv se nao existir, instala as
dependencias e entao roda

```
import  ->  status  ->  review  ->  export (dry-run)  ->  export --apply
```

Sem `-Apply` ele para no dry-run e **nao escreve nada**. Se alguma fatura for
rejeitada na reconciliacao, ou se sobrar merchant sem categoria, ele para ali
e diz o porque.

Opcoes uteis:

| Flag | Efeito |
|---|---|
| `-Apply` | escreve na planilha (pede confirmacao) |
| `-Force` | com `-Apply`, nao pede confirmacao |
| `-SkipInstall` | pula venv e pip (~8s em vez do setup completo) |
| `-SkipReview` | nao abre a revisao de merchants |
| `-Recreate` | apaga o venv e cria de novo |
| `-Csv detalhe.csv` | tambem grava o CSV linha a linha |
| `-Planilha caminho.xlsx` | usa outra planilha |

Se o PowerShell recusar a execucao do script, rode uma vez:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

ou chame com `powershell -ExecutionPolicy Bypass -File .\executar.ps1`.

### Na mao, sem o script

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

python -m nubank import faturas/          # le, reconcilia e grava
python -m nubank review                   # classifica merchants novos
python -m nubank export                   # dry-run: mostra o que mudaria
python -m nubank export --apply           # escreve na planilha
```

Outros comandos:

```powershell
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

### Corrigir uma categoria errada

Errou no `review`, ou mudou de ideia? Edite a linha no `merchants.yml`:

```yaml
  # -- adicionados via review --
  "loja fisica lj 16 jac": mercado    # estava 'compras'
```

E so isso. **Nao precisa reimportar nada**: a categoria e sempre derivada da
tabela no momento da leitura, entao o proximo `status` ou `export` ja usa o
valor novo, retroativo a todos os meses onde aquele merchant aparece.

Confira o efeito antes de escrever na planilha:

```powershell
python -m nubank status     # ve as colunas mudarem
python -m nubank export     # dry-run, celula a celula
```

### As colunas B e H

`Fatura total` (B) e `Gasto do periodo` (H) **nao sao a mesma coisa**:

```
B  Fatura total      = total a pagar, o que sai do bolso no vencimento
H  Gasto do periodo  = tudo que a fatura lancou entre o fechamento anterior e o atual
B - H                = saldo que veio da fatura anterior (anterior - pagamentos)
```

No mes em que voce paga R$ 50 a mais, B fica R$ 50 abaixo de H. Isso nao e
gasto perdido nem bug: e credito carregado. A coluna G escreve a conta em cada
linha.

**Quem C+D+E decompoem e H, nao B.** Por isso a coluna F (`Nao classificado`)
confere contra H:

```
F  =  H - SOMA(C:E)      ->  zero quando esta tudo classificado
```

Se F fosse contra B, ela mostraria o saldo carregado — um numero grande e
vermelho numa coluna cujo nome promete outra coisa. Como a planilha nasceu sem
a coluna H, o `export` cuida disso: escreve o cabecalho de H e reescreve as
formulas de F na primeira vez que roda, e depois e idempotente.

H sai da **soma das transacoes**, nao de `compras + IOF + outros` do resumo do
PDF. Os dois divergem em 1 centavo em metade das faturas (arredondamento do
Nubank), e como C/D/E vem das transacoes, usar o resumo faria F oscilar entre
`0,01` e `-0,01` em vez de mostrar zero limpo.

### Escrita na planilha

- `export` sem `--apply` e dry-run: mostra celula a celula o que mudaria e nao
  escreve nada.
- `export --apply` copia a planilha para `backups/` com carimbo de data e hora
  **antes** de escrever.
- Escreve `B`, `C`, `D`, `E`, `G` e `H` das linhas de mes, mais o cabecalho de
  `H`, o total de `H` e as formulas de `F`. Nada fora disso e tocado.
- Rodar o export duas vezes seguidas nao muda nada na segunda.

### Por que a escrita nao usa openpyxl

O `openpyxl.save()` reescreve o arquivo inteiro, e nesta planilha isso **perde
os atributos `apply*` dos estilos**: 50 `applyNumberFormat`, 68 `applyFont`, 54
`applyBorder`, 7 `applyFill`. Sem `applyNumberFormat="1"`, o Excel ignora o
`numFmtId` da celula e herda o formato do estilo nomeado, que aqui e `General`.
Na pratica: a coluna de mes passa a mostrar `46023` em vez de `jan/26`, e os
valores em reais viram numero cru.

Pior, o openpyxl nao enxerga o proprio estrago — ele le o `numFmtId` direto e
ignora a flag. Conferir a saida dele com ele mesmo da sempre verde.

Entao `nubank/xlsx.py` trata o .xlsx como o zip que ele e: le todas as partes,
altera **so** o XML da aba Cartao (mais o `calcPr` do workbook, para o Excel
recalcular ao abrir) e grava as outras byte a byte. Na planilha real isso
significa 30 das 32 partes intactas, `styles.xml` e `sharedStrings.xml`
inclusos.

O openpyxl continua sendo usado para ler, o que nao danifica nada.

Os testes comparam o **zip cru**, nunca a leitura do openpyxl.

## Arquivos

| Arquivo | O que e |
|---|---|
| `executar.ps1` | Roda a pipeline inteira na ordem certa, incluindo venv e dependencias. |
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
