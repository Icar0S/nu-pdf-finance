<#
.SYNOPSIS
    Prepara o ambiente e roda a pipeline inteira, na ordem certa.

.DESCRIPTION
    Cria o venv (se nao existir), instala as dependencias e executa:

        import  ->  status  ->  review  ->  export (dry-run)  ->  export --apply

    Sem -Apply o script para no dry-run e nao escreve nada na planilha.

    O arquivo e ASCII puro de proposito: o Windows PowerShell 5.1 le .ps1 sem
    BOM usando a codepage ANSI, e acento aqui viraria lixo no console.

.PARAMETER Apply
    Escreve na planilha depois de mostrar o dry-run. Pede confirmacao antes.

.PARAMETER Force
    Com -Apply, nao pede confirmacao.

.PARAMETER SkipReview
    Nao abre a revisao de merchants pendentes.

.PARAMETER SkipInstall
    Pula venv e pip. Util quando o ambiente ja esta pronto.

.PARAMETER Recreate
    Apaga o venv e cria de novo do zero.

.PARAMETER Faturas
    Pasta com os PDFs. Padrao: faturas

.PARAMETER Planilha
    Caminho do .xlsx. Padrao: o que o proprio comando encontrar na raiz.

.PARAMETER Csv
    Tambem grava um CSV de detalhe nesse caminho.

.EXAMPLE
    .\executar.ps1
    Prepara o ambiente, importa, revisa e mostra o dry-run.

.EXAMPLE
    .\executar.ps1 -Apply
    Idem, e escreve na planilha depois de confirmar.

.NOTES
    Se o PowerShell recusar a execucao do script, rode uma vez:
        Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
    ou chame assim:
        powershell -ExecutionPolicy Bypass -File .\executar.ps1
#>

[CmdletBinding()]
param(
    [switch]$Apply,
    [switch]$Force,
    [switch]$SkipReview,
    [switch]$SkipInstall,
    [switch]$Recreate,
    [string]$Faturas = "faturas",
    [string]$Planilha,
    [string]$Csv
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #

$script:Etapa = 0
$script:TotalEtapas = 7

function Write-Etapa {
    param([string]$Texto)
    $script:Etapa++
    Write-Host ""
    Write-Host "[$script:Etapa/$script:TotalEtapas] $Texto" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Texto)
    Write-Host "      $Texto" -ForegroundColor Green
}

function Write-Aviso {
    param([string]$Texto)
    Write-Host "      $Texto" -ForegroundColor Yellow
}

function Write-Falha {
    param([string]$Texto)
    Write-Host ""
    Write-Host "ERRO: $Texto" -ForegroundColor Red
}

function Assert-ExitCode {
    <#
        $ErrorActionPreference nao pega falha de executavel nativo: python que
        sai com codigo 1 passaria batido. Por isso a checagem e explicita.
    #>
    param([string]$Mensagem)
    if ($LASTEXITCODE -ne 0) {
        Write-Falha $Mensagem
        exit $LASTEXITCODE
    }
}

function Find-PythonBase {
    <#
        Procura um Python 3.10+ para criar o venv. O py.exe launcher e a via
        mais confiavel no Windows; python.exe do PATH e o plano B.
    #>
    $candidatos = @(
        @{ Exe = "py";     Args = @("-3") },
        @{ Exe = "python"; Args = @() }
    )
    foreach ($c in $candidatos) {
        $cmd = Get-Command $c.Exe -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        # Splatting precisa de variavel: '@(...)' seria array subexpression, e
        # array virando argumento de executavel nativo e ambiguo no PS 5.1.
        $argsVersao = $c.Args + @(
            "-c", "import sys; print('%d.%d' % sys.version_info[:2])"
        )
        $versao = & $c.Exe @argsVersao 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $versao) { continue }
        $partes = $versao.Trim().Split(".")
        if ([int]$partes[0] -eq 3 -and [int]$partes[1] -ge 10) {
            return @{ Exe = $c.Exe; Args = $c.Args; Versao = $versao.Trim() }
        }
        Write-Aviso "$($c.Exe) e Python $($versao.Trim()); o projeto pede 3.10+."
    }
    return $null
}

# --------------------------------------------------------------------------- #

$Raiz = $PSScriptRoot
Push-Location $Raiz
try {
    Write-Host ""
    Write-Host "nu-pdf-finance" -ForegroundColor White
    Write-Host "faturas do Nubank -> aba Cartao da planilha" -ForegroundColor DarkGray
    Write-Host ("-" * 60) -ForegroundColor DarkGray

    $VenvDir = Join-Path $Raiz ".venv"
    $VenvPython = Join-Path $VenvDir "Scripts\python.exe"
    $Activate = Join-Path $VenvDir "Scripts\Activate.ps1"

    # ---------------------------------------------------------------- 1. venv
    Write-Etapa "Ambiente virtual"

    if ($SkipInstall) {
        if (-not (Test-Path $VenvPython)) {
            Write-Falha "-SkipInstall foi passado, mas nao existe venv em $VenvDir."
            exit 1
        }
        Write-Aviso "pulado (-SkipInstall)."
    }
    else {
        if ($Recreate -and (Test-Path $VenvDir)) {
            Write-Host "      removendo venv anterior..."
            Remove-Item -Recurse -Force $VenvDir
        }

        if (Test-Path $VenvPython) {
            Write-Ok "ja existe em .venv"
        }
        else {
            $base = Find-PythonBase
            if (-not $base) {
                Write-Falha @"
nao encontrei Python 3.10 ou superior.
Instale em https://www.python.org/downloads/ marcando
'Add python.exe to PATH', abra um PowerShell novo e rode o script de novo.
"@
                exit 1
            }
            Write-Host "      criando venv com Python $($base.Versao)..."
            $argsVenv = $base.Args + @("-m", "venv", $VenvDir)
            & $base.Exe @argsVenv
            Assert-ExitCode "falhou ao criar o venv em $VenvDir."
            Write-Ok "criado em .venv"
        }
    }

    # Ativa para a sessao. A pipeline abaixo chama $VenvPython pelo caminho
    # completo de qualquer jeito, entao um bloqueio de ExecutionPolicy aqui
    # atrapalha o prompt, nao o funcionamento.
    if (Test-Path $Activate) {
        try {
            . $Activate
            Write-Ok "venv ativado nesta sessao"
        }
        catch {
            Write-Aviso "nao consegui ativar o venv (ExecutionPolicy?); seguindo com o python do venv direto."
        }
    }

    # -------------------------------------------------------- 2. dependencias
    Write-Etapa "Dependencias"

    if ($SkipInstall) {
        Write-Aviso "pulado (-SkipInstall)."
    }
    else {
        $req = Join-Path $Raiz "requirements.txt"
        if (-not (Test-Path $req)) {
            Write-Falha "requirements.txt nao encontrado em $Raiz."
            exit 1
        }
        & $VenvPython -m pip install --upgrade pip --quiet --disable-pip-version-check
        Assert-ExitCode "falhou ao atualizar o pip."
        & $VenvPython -m pip install -r $req --quiet --disable-pip-version-check
        Assert-ExitCode "falhou ao instalar as dependencias de requirements.txt."
        Write-Ok "pdfplumber, openpyxl, PyYAML e pytest prontos"
    }

    # ------------------------------------------------------------- 3. import
    Write-Etapa "Import das faturas"

    $PastaFaturas = if ([System.IO.Path]::IsPathRooted($Faturas)) {
        $Faturas
    }
    else {
        Join-Path $Raiz $Faturas
    }

    if (-not (Test-Path $PastaFaturas)) {
        Write-Falha "pasta de faturas nao encontrada: $PastaFaturas"
        exit 1
    }

    $pdfs = @(Get-ChildItem -Path $PastaFaturas -Filter *.pdf -File)
    if ($pdfs.Count -eq 0) {
        Write-Falha "nenhum PDF em $PastaFaturas. Baixe as faturas do app do Nubank e coloque ali."
        exit 1
    }
    Write-Host "      $($pdfs.Count) PDF(s) em $Faturas"
    Write-Host ""

    & $VenvPython -m nubank import $PastaFaturas
    # O import so devolve != 0 quando alguma fatura foi REJEITADA na
    # reconciliacao. Nesse caso parar e o comportamento certo: os numeros
    # daquele PDF nao fecham e nada dele deve chegar na planilha.
    Assert-ExitCode "alguma fatura foi rejeitada na reconciliacao. Nada foi gravado dela."

    # ------------------------------------------------------------- 4. status
    Write-Etapa "Situacao atual"
    Write-Host ""
    & $VenvPython -m nubank status
    Assert-ExitCode "falhou ao ler o status."

    # ------------------------------------------------------------- 5. review
    Write-Etapa "Revisao de merchants"

    if ($SkipReview) {
        Write-Aviso "pulado (-SkipReview). O export vai travar se houver pendente."
    }
    else {
        Write-Host ""
        & $VenvPython -m nubank review
        Assert-ExitCode "falhou na revisao de merchants."
    }

    # ----------------------------------------------------------- 6. conferir
    Write-Etapa "Formulas da planilha"

    $argsConferir = @("-m", "nubank", "conferir")
    if ($Planilha) { $argsConferir += @("--planilha", $Planilha) }

    & $VenvPython @argsConferir
    if ($LASTEXITCODE -ne 0) {
        # Nao e fatal: o export escreve valores, e as formulas quebradas estao
        # em outra aba. Mas seguir sem avisar deixaria o Painel mentindo.
        Write-Host ""
        Write-Aviso "formula(s) quebrada(s) acima. Reparando..."
        Write-Host ""
        & $VenvPython @($argsConferir + @("--reparar"))
        Assert-ExitCode "nao consegui reparar as formulas da planilha."
    }

    # ------------------------------------------------------------- 7. export
    Write-Etapa "Export para a planilha"

    $argsExport = @("-m", "nubank", "export")
    if ($Planilha) { $argsExport += @("--planilha", $Planilha) }
    if ($Csv) { $argsExport += @("--csv", $Csv) }

    Write-Host ""
    & $VenvPython @argsExport
    if ($LASTEXITCODE -ne 0) {
        Write-Falha @"
o export foi bloqueado (veja o motivo acima).
O caso comum e merchant sem categoria: rode
    .\executar.ps1 -SkipInstall
e classifique os pendentes na etapa de revisao.
"@
        exit $LASTEXITCODE
    }

    if (-not $Apply) {
        Write-Host ""
        Write-Host ("-" * 60) -ForegroundColor DarkGray
        Write-Host "Dry-run concluido. Nada foi escrito na planilha." -ForegroundColor Green
        Write-Host "Para aplicar de verdade:  .\executar.ps1 -Apply" -ForegroundColor White
        exit 0
    }

    if (-not $Force) {
        Write-Host ""
        $resposta = Read-Host "      Aplicar essas mudancas na planilha? (s/N)"
        if ($resposta.Trim().ToLower() -ne "s") {
            Write-Host ""
            Write-Aviso "cancelado. Nada foi escrito."
            exit 0
        }
    }

    $argsExport += "--apply"
    Write-Host ""
    & $VenvPython @argsExport
    Assert-ExitCode "falhou ao escrever na planilha."

    Write-Host ""
    Write-Host ("-" * 60) -ForegroundColor DarkGray
    Write-Host "Planilha atualizada. O backup esta em backups\." -ForegroundColor Green
    exit 0
}
finally {
    Pop-Location
}
