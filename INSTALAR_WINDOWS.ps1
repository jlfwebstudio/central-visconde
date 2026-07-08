# Configura o ambiente da Central Visconde no Windows: cria o .venv,
# instala as dependencias e o Chromium do Playwright, e cria as pastas
# usadas pelas automacoes.
#
# Rode a partir da pasta raiz do projeto (onde este arquivo esta), depois
# de clonar o repositorio com git clone. Nao precisa copiar a pasta pra
# outro lugar -- diferente do instalador do Mac, aqui o git ja cuida de
# manter a pasta certa.

$ErrorActionPreference = "Stop"

Write-Host "================================================================"
Write-Host "       INSTALACAO DA CENTRAL VISCONDE - Windows"
Write-Host "================================================================"
Write-Host ""

Set-Location $PSScriptRoot

# Localiza um Python utilizavel (prefere o launcher "py", que normalmente
# ja aponta pra versao mais recente instalada).
$PythonCmd = $null

if (Get-Command py -ErrorAction SilentlyContinue) {
    $teste = & py -c "import sys, tkinter; print(sys.version_info >= (3,9))" 2>$null
    if ($teste -match "True") {
        $PythonCmd = "py"
    }
}

if (-not $PythonCmd -and (Get-Command python -ErrorAction SilentlyContinue)) {
    $teste = & python -c "import sys, tkinter; print(sys.version_info >= (3,9))" 2>$null
    if ($teste -match "True") {
        $PythonCmd = "python"
    }
}

if (-not $PythonCmd) {
    Write-Host "[ERRO] Nao encontrei Python 3.9+ com Tkinter instalado." -ForegroundColor Red
    Write-Host "Instale o Python em https://python.org/downloads (marque a opcao"
    Write-Host "'tcl/tk and IDLE' durante a instalacao) e rode este script de novo."
    Read-Host "Pressione Enter para fechar"
    exit 1
}

Write-Host "Python encontrado: $PythonCmd"
& $PythonCmd --version

Write-Host ""
Write-Host "[1/4] Criando ambiente virtual (.venv)..."
if (-not (Test-Path ".venv")) {
    & $PythonCmd -m venv .venv
} else {
    Write-Host "Ja existe um .venv, mantendo o atual."
}

$VenvPython = ".\.venv\Scripts\python.exe"

Write-Host ""
Write-Host "[2/4] Atualizando pip e instalando as dependencias..."
& $VenvPython -m pip install --upgrade pip setuptools wheel
& $VenvPython -m pip install -r requirements.txt

Write-Host ""
Write-Host "[3/4] Instalando o Chromium do Playwright..."
& $VenvPython -m playwright install chromium

Write-Host ""
Write-Host "[4/4] Criando pastas de trabalho..."
$Pastas = @(
    "downloads\relatorios_completos",
    "downloads\roteirizacao\mobyan",
    "downloads\roteirizacao\ogea",
    "logs\roteirizacao",
    "logs\pdfs",
    "logs\abonos_ogea",
    "outputs\pendencias_do_dia\backups",
    "outputs\por_prestador_imagens",
    "outputs\whatsapp_temp",
    "outputs\roteirizacao",
    "outputs\pdfs",
    "outputs\abonos_ogea",
    "bases\backups_roteirizacao",
    "whatsapp_profile"
)
foreach ($pasta in $Pastas) {
    New-Item -ItemType Directory -Force -Path $pasta | Out-Null
}

if (-not (Test-Path ".env")) {
    if (Test-Path ".env.exemplo") {
        Copy-Item ".env.exemplo" ".env"
    }
    Write-Host ""
    Write-Host "================================================================"
    Write-Host "INSTALACAO TECNICA CONCLUIDA, MAS FALTA A CONFIGURACAO."
    Write-Host "================================================================"
    Write-Host "O arquivo .env foi criado sem usuario e senha."
    Write-Host "Copie o .env de outra maquina ja configurada, ou preencha manualmente."
    Read-Host "Pressione Enter para fechar"
    exit 0
}

Write-Host ""
Write-Host "Rodando diagnostico..."
& $VenvPython app\diagnostico_mac.py
$status = $LASTEXITCODE

if ($status -ne 0) {
    Write-Host ""
    Write-Host "A instalacao terminou, mas o diagnostico encontrou um problema." -ForegroundColor Yellow
    Read-Host "Pressione Enter para fechar"
    exit $status
}

Write-Host ""
Write-Host "================================================================"
Write-Host "        CENTRAL VISCONDE INSTALADA COM SUCESSO NO WINDOWS"
Write-Host "================================================================"
Write-Host "Use o INICIAR_CENTRAL_WINDOWS.bat para abrir a Central."
Read-Host "Pressione Enter para fechar"
