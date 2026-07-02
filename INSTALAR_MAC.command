#!/bin/bash
set -u

SOURCE_DIR="$(cd "$(dirname "$0")" && pwd -P)"
TARGET_DIR="$HOME/mobyan-automacoes"

echo "================================================================"
echo "       INSTALAÇÃO DA CENTRAL MOBYAN - macOS / Apple Silicon"
echo "================================================================"
echo

# Instala/atualiza sempre em uma pasta fixa para o atalho continuar funcionando.
if [ "$SOURCE_DIR" != "$TARGET_DIR" ]; then
    echo "Copiando a Central para: $TARGET_DIR"
    mkdir -p "$TARGET_DIR"
    rsync -a \
        --exclude '.venv' \
        --exclude '.env' \
        --exclude 'outputs' \
        --exclude 'downloads' \
        --exclude 'logs' \
        --exclude 'whatsapp_profile' \
        --exclude '.DS_Store' \
        "$SOURCE_DIR/" "$TARGET_DIR/"

    # Se o usuário colocou seu próprio .env junto do pacote, migra sem exibir.
    if [ ! -f "$TARGET_DIR/.env" ] && [ -f "$SOURCE_DIR/.env" ]; then
        cp -p "$SOURCE_DIR/.env" "$TARGET_DIR/.env"
        chmod 600 "$TARGET_DIR/.env" 2>/dev/null || true
    fi
fi

cd "$TARGET_DIR" || exit 1
chmod +x ./*.command 2>/dev/null || true

# Localiza uma instalação de Python compatível e com Tkinter.
PYTHON_EXE=""
CANDIDATOS=(
    "/opt/homebrew/bin/python3.12"
    "/opt/homebrew/bin/python3.11"
    "/usr/local/bin/python3.12"
    "/usr/local/bin/python3.11"
    "python3.12"
    "python3.11"
    "python3.10"
    "python3"
)

for candidato in "${CANDIDATOS[@]}"; do
    if command -v "$candidato" >/dev/null 2>&1 || [ -x "$candidato" ]; then
        caminho="$candidato"
        if command -v "$candidato" >/dev/null 2>&1; then
            caminho="$(command -v "$candidato")"
        fi
        HOST_ARCH_CANDIDATO="$(uname -m)"
        PYTHON_ARCH="$("$caminho" -c 'import platform; print(platform.machine())' 2>/dev/null || true)"

        if [ "$HOST_ARCH_CANDIDATO" = "arm64" ] && [ "$PYTHON_ARCH" != "arm64" ]; then
            continue
        fi

        if "$caminho" -c 'import sys, tkinter; raise SystemExit(0 if sys.version_info >= (3,9) else 1)' >/dev/null 2>&1; then
            PYTHON_EXE="$caminho"
            break
        fi
    fi
done

if [ -z "$PYTHON_EXE" ]; then
    echo "[ERRO] Não encontrei Python 3.9+ com Tkinter."
    echo
    echo "Instale o Python 3.12 para macOS pelo instalador oficial do Python."
    echo "Depois execute novamente INSTALAR_MAC.command."
    echo
    read -r -p "Pressione Enter para fechar..."
    exit 1
fi

echo "Python encontrado: $PYTHON_EXE"
"$PYTHON_EXE" --version

HOST_ARCH="$(uname -m)"
RECRIAR_VENV=0

if [ -x ".venv/bin/python" ]; then
    VENV_ARCH="$(.venv/bin/python -c 'import platform; print(platform.machine())' 2>/dev/null || true)"
    if [ "$VENV_ARCH" != "$HOST_ARCH" ]; then
        echo "O ambiente virtual existente é de outra arquitetura. Recriando..."
        RECRIAR_VENV=1
    fi
else
    RECRIAR_VENV=1
fi

if [ "$RECRIAR_VENV" -eq 1 ]; then
    rm -rf .venv
    "$PYTHON_EXE" -m venv .venv || exit 1
fi

echo
echo "[1/4] Atualizando o pip..."
.venv/bin/python -m pip install --upgrade pip setuptools wheel || exit 1

echo
echo "[2/4] Instalando as dependências..."
.venv/bin/python -m pip install -r requirements.txt || exit 1

echo
echo "[3/4] Instalando o Chromium das automações..."
.venv/bin/python -m playwright install chromium || exit 1

echo
echo "[4/4] Criando pastas e atalhos..."
mkdir -p \
    downloads/relatorios_completos \
    downloads/roteirizacao/mobyan \
    downloads/roteirizacao/ogea \
    logs/roteirizacao \
    logs/pdfs \
    outputs/pendencias_do_dia/backups \
    outputs/por_prestador_imagens \
    outputs/whatsapp_temp \
    outputs/roteirizacao \
    outputs/pdfs \
    whatsapp_profile

# Cria um aplicativo local sem assinatura, evitando depender do Terminal no uso diário.
APP_DIR="$HOME/Applications/Central Mobyan.app"
mkdir -p "$APP_DIR/Contents/MacOS"
cat > "$APP_DIR/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>Central Mobyan</string>
    <key>CFBundleIdentifier</key>
    <string>br.local.centralmobyan</string>
    <key>CFBundleName</key>
    <string>Central Mobyan</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>11.0</string>
</dict>
</plist>
PLIST
cat > "$APP_DIR/Contents/MacOS/Central Mobyan" <<'APP'
#!/bin/bash
PROJECT_DIR="$HOME/mobyan-automacoes"
cd "$PROJECT_DIR" || exit 1
mkdir -p logs
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
nohup "$PROJECT_DIR/.venv/bin/python" "$PROJECT_DIR/app/central_mobyan.py" >> "$PROJECT_DIR/logs/central_mobyan.log" 2>&1 &
APP
chmod +x "$APP_DIR/Contents/MacOS/Central Mobyan"
mkdir -p "$HOME/Applications"
if [ -d "$HOME/Desktop" ]; then
    ln -sfn "$APP_DIR" "$HOME/Desktop/Central Mobyan.app"
fi

if [ ! -f ".env" ]; then
    cp .env.exemplo .env
    echo
    echo "================================================================"
    echo "INSTALAÇÃO TÉCNICA CONCLUÍDA, MAS FALTA A CONFIGURAÇÃO."
    echo "================================================================"
    echo "O arquivo .env foi criado sem usuário e senha."
    echo "Copie o .env da Central do Windows ou execute:"
    echo "MIGRAR_CONFIGURACAO_DO_WINDOWS.command"
    echo
    read -r -p "Pressione Enter para fechar..."
    exit 0
fi

export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
.venv/bin/python app/diagnostico_mac.py
STATUS=$?

if [ "$STATUS" -ne 0 ]; then
    echo
    echo "A instalação terminou, mas o diagnóstico encontrou um problema."
    read -r -p "Pressione Enter para fechar..."
    exit "$STATUS"
fi

echo
echo "================================================================"
echo "        CENTRAL MOBYAN INSTALADA COM SUCESSO NO MAC"
echo "================================================================"
echo "Abra pelo aplicativo Central Mobyan na Mesa ou em Aplicativos."
echo
open "$APP_DIR"
read -r -p "Pressione Enter para fechar o instalador..."
