@echo off
REM Clique duas vezes neste arquivo para instalar a Central Visconde no Windows.
REM Ele so chama o INSTALAR_WINDOWS.ps1 ignorando a politica de execucao do
REM PowerShell (que por padrao bloqueia rodar scripts .ps1 com duplo clique).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALAR_WINDOWS.ps1"
if errorlevel 1 (
    echo.
    echo O PowerShell terminou com erro antes de conseguir pausar sozinho.
    echo Verifique a pasta "logs" para o arquivo de instalacao.
)
pause
