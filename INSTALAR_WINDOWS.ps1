# Configura o ambiente da Central Visconde no Windows: cria o .venv,
# instala as dependencias e o Chromium do Playwright, e cria as pastas
# usadas pelas automacoes.
#
# Rode a partir da pasta raiz do projeto (onde este arquivo esta), depois
# de clonar o repositorio com git clone. Nao precisa copiar a pasta pra
# outro lugar -- diferente do instalador do Mac, aqui o git ja cuida de
# manter a pasta certa.

# Nao usamos $ErrorActionPreference = "Stop" aqui: no PowerShell 5.1 (o do
# Windows), isso faz qualquer saida no stderr de um programa externo (py.exe,
# pip, etc.) virar erro fatal, mesmo redirecionada com "2>$null". Os passos
# criticos abaixo checam o resultado explicitamente (throw / exit) em vez de
# depender dessa preferencia.
$ErrorActionPreference = "Continue"

Set-Location $PSScriptRoot

New-Item -ItemType Directory -Force -Path "logs" | Out-Null
$ArquivoLog = "logs\instalacao_windows_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
Start-Transcript -Path $ArquivoLog | Out-Null

try {

Write-Host "================================================================"
Write-Host "       INSTALACAO DA CENTRAL VISCONDE - Windows"
Write-Host "================================================================"
Write-Host ""

# Localiza um Python utilizavel. Prefere versoes especificas (3.13 ate 3.9)
# via "py -3.X" em vez de confiar no "py" generico, que pega a versao mais
# nova instalada -- e uma versao Python muito recente pode ainda nao ter
# wheel pronto de pandas no PyPI, obrigando o pip a compilar do zero (o que
# falha sem as ferramentas de build do Visual Studio instaladas).
$PythonExe = $null
$PythonArgs = @()

foreach ($versao in @("3.13", "3.12", "3.11", "3.10", "3.9")) {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $teste = & py "-$versao" -c "import sys, tkinter; print('OK')" 2>$null
        if ($teste -match "OK") {
            $PythonExe = "py"
            $PythonArgs = @("-$versao")
            break
        }
    }
}

if (-not $PythonExe -and (Get-Command py -ErrorAction SilentlyContinue)) {
    $teste = & py -c "import sys, tkinter; print(sys.version_info >= (3,9))" 2>$null
    if ($teste -match "True") {
        $PythonExe = "py"
    }
}

if (-not $PythonExe -and (Get-Command python -ErrorAction SilentlyContinue)) {
    $teste = & python -c "import sys, tkinter; print(sys.version_info >= (3,9))" 2>$null
    if ($teste -match "True") {
        $PythonExe = "python"
    }
}

if (-not $PythonExe) {
    Write-Host "[ERRO] Nao encontrei Python 3.9+ com Tkinter instalado." -ForegroundColor Red
    Write-Host "Instale o Python 3.12 em https://python.org/downloads (marque a opcao"
    Write-Host "'tcl/tk and IDLE' durante a instalacao) e rode este script de novo."
    Write-Host "Evite instalar so a versao mais nova do Python: pacotes como"
    Write-Host "pandas as vezes demoram a publicar suporte pronto (wheel) pra"
    Write-Host "versoes recem-lancadas, e a instalacao pode falhar por isso."
    Read-Host "Pressione Enter para fechar"
    exit 1
}

Write-Host "Python encontrado: $PythonExe $PythonArgs"
& $PythonExe @PythonArgs --version

Write-Host ""
Write-Host "[1/4] Criando ambiente virtual (.venv)..."
if (-not (Test-Path ".venv")) {
    & $PythonExe @PythonArgs -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "Criacao do .venv falhou (codigo $LASTEXITCODE)." }
} else {
    Write-Host "Ja existe um .venv, mantendo o atual."
}

$VenvPython = ".\.venv\Scripts\python.exe"

Write-Host ""
Write-Host "[2/4] Atualizando pip e instalando as dependencias..."
& $VenvPython -m pip install --upgrade pip setuptools wheel
if ($LASTEXITCODE -ne 0) { throw "Atualizacao do pip falhou (codigo $LASTEXITCODE)." }
& $VenvPython -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "Instalacao das dependencias (requirements.txt) falhou (codigo $LASTEXITCODE)." }

Write-Host ""
Write-Host "[3/4] Instalando o Chromium do Playwright..."
& $VenvPython -m playwright install chromium
if ($LASTEXITCODE -ne 0) { throw "Instalacao do Chromium do Playwright falhou (codigo $LASTEXITCODE)." }

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
    "outputs\por_prestador_imagens_estoque",
    "outputs\whatsapp_temp",
    "outputs\roteirizacao",
    "outputs\pdfs",
    "outputs\abonos_ogea",
    "bases\backups_roteirizacao",
    "whatsapp_profile",
    "onedrive_profile"
)
foreach ($pasta in $Pastas) {
    New-Item -ItemType Directory -Force -Path $pasta | Out-Null
}

Write-Host "Criando atalho na Area de Trabalho..."
try {
    $CaminhoDesktop = [Environment]::GetFolderPath("Desktop")
    $WshShell = New-Object -ComObject WScript.Shell
    $Atalho = $WshShell.CreateShortcut("$CaminhoDesktop\ViscondeApp.lnk")
    $Atalho.TargetPath = "$PSScriptRoot\INICIAR_CENTRAL_WINDOWS.bat"
    $Atalho.WorkingDirectory = "$PSScriptRoot"
    $IconePath = "$PSScriptRoot\assets\logo_visconde.ico"
    if (Test-Path $IconePath) {
        $Atalho.IconLocation = $IconePath
    }
    $Atalho.Save()
} catch {
    Write-Host "[AVISO] Nao consegui criar o atalho na Area de Trabalho automaticamente." -ForegroundColor Yellow
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

}
catch {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host "ERRO DURANTE A INSTALACAO" -ForegroundColor Red
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host $_.ScriptStackTrace
    Write-Host ""
    Write-Host "Detalhes completos foram salvos em: $ArquivoLog"
    Write-Host "Envie esse arquivo para quem estiver te ajudando com a instalacao."
    Read-Host "Pressione Enter para fechar"
}
finally {
    Stop-Transcript | Out-Null
}
