#!/bin/bash
# Instala o ViscondeApp em distros baseadas em Debian/Ubuntu (Zorin OS incluída).
#
# Diferente do instalador do Mac, este roda direto de onde o repositório foi
# clonado com "git clone" (mesma filosofia do instalador do Windows) — não
# copia a pasta pra outro lugar. Rode a partir da raiz do projeto:
#
#   chmod +x INSTALAR_LINUX.sh && ./INSTALAR_LINUX.sh
#
set -u

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd -P)"
cd "$PROJECT_DIR" || exit 1

echo "================================================================"
echo "       INSTALAÇÃO DO VISCONDEAPP - Linux"
echo "================================================================"
echo
echo "Pasta do projeto: $PROJECT_DIR"
echo

# --- Localiza um Python 3.9+ com Tkinter -----------------------------------
PYTHON_EXE=""
CANDIDATOS=(python3.12 python3.11 python3.10 python3)

for candidato in "${CANDIDATOS[@]}"; do
    if command -v "$candidato" >/dev/null 2>&1; then
        caminho="$(command -v "$candidato")"
        if "$caminho" -c 'import sys, tkinter; raise SystemExit(0 if sys.version_info >= (3,9) else 1)' >/dev/null 2>&1; then
            PYTHON_EXE="$caminho"
            break
        fi
    fi
done

if [ -z "$PYTHON_EXE" ]; then
    echo "[ERRO] Não encontrei Python 3.9+ com Tkinter."
    echo
    echo "Em distros baseadas em Ubuntu/Debian (Zorin incluída), instale com:"
    echo
    echo "    sudo apt update"
    echo "    sudo apt install python3 python3-venv python3-tk"
    echo
    echo "Depois execute novamente ./INSTALAR_LINUX.sh"
    echo
    read -r -p "Pressione Enter para fechar..."
    exit 1
fi

echo "Python encontrado: $PYTHON_EXE"
"$PYTHON_EXE" --version

# --- Ambiente virtual --------------------------------------------------------
if [ ! -x ".venv/bin/python" ]; then
    echo
    echo "Criando ambiente virtual (.venv)..."
    if ! "$PYTHON_EXE" -m venv .venv; then
        echo
        echo "[ERRO] Não consegui criar o .venv. Em distros baseadas em"
        echo "Ubuntu/Debian isso costuma faltar o pacote python3-venv:"
        echo
        echo "    sudo apt install python3-venv"
        echo
        read -r -p "Pressione Enter para fechar..."
        exit 1
    fi
else
    echo
    echo "Já existe um .venv, mantendo o atual."
fi

echo
echo "[1/5] Atualizando o pip..."
.venv/bin/python -m pip install --upgrade pip setuptools wheel || exit 1

echo
echo "[2/5] Instalando as dependências..."
.venv/bin/python -m pip install -r requirements.txt || exit 1

echo
echo "[3/5] Instalando o Chromium das automações..."
.venv/bin/python -m playwright install chromium || exit 1

echo
echo "[4/5] Instalando as bibliotecas de sistema que o Chromium precisa..."
echo "(pode pedir sua senha do sudo)"
if ! .venv/bin/python -m playwright install-deps chromium; then
    echo
    echo "[AVISO] Não consegui instalar as bibliotecas de sistema automaticamente."
    echo "O diagnóstico no final deste instalador vai confirmar se o Chromium"
    echo "abre mesmo assim. Se não abrir, rode manualmente:"
    echo
    echo "    sudo .venv/bin/python -m playwright install-deps chromium"
fi

echo
echo "[5/5] Criando pastas de trabalho e o atalho..."
mkdir -p \
    downloads/relatorios_completos \
    downloads/roteirizacao/mobyan \
    downloads/roteirizacao/ogea \
    logs/roteirizacao \
    logs/pdfs \
    logs/abonos_ogea \
    outputs/pendencias_do_dia/backups \
    outputs/por_prestador_imagens \
    outputs/whatsapp_temp \
    outputs/roteirizacao \
    outputs/pdfs \
    outputs/abonos_ogea \
    bases/backups_roteirizacao \
    whatsapp_profile

# --- Atalho no menu de aplicativos (padrão XDG, não precisa de root) --------
chmod +x INICIAR_CENTRAL_LINUX.sh DIAGNOSTICO_LINUX.sh 2>/dev/null || true

LANCADOR="$PROJECT_DIR/.viscondeapp_lancador.sh"
cat > "$LANCADOR" <<LAUNCHER
#!/bin/bash
cd "$PROJECT_DIR" || exit 1
mkdir -p logs
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
exec "$PROJECT_DIR/.venv/bin/python" "$PROJECT_DIR/app/central_mobyan.py" >> "$PROJECT_DIR/logs/central_mobyan.log" 2>&1
LAUNCHER
chmod +x "$LANCADOR"

DESKTOP_DIR="$HOME/.local/share/applications"
mkdir -p "$DESKTOP_DIR"
DESKTOP_FILE="$DESKTOP_DIR/viscondeapp.desktop"
cat > "$DESKTOP_FILE" <<DESKTOP
[Desktop Entry]
Type=Application
Name=ViscondeApp
Comment=Central Visconde - automações do dia a dia
Exec="$LANCADOR"
Icon=$PROJECT_DIR/assets/logo_visconde_app.png
Terminal=false
Categories=Office;
DESKTOP
chmod +x "$DESKTOP_FILE"

if [ -d "$HOME/Desktop" ]; then
    cp -f "$DESKTOP_FILE" "$HOME/Desktop/ViscondeApp.desktop"
    chmod +x "$HOME/Desktop/ViscondeApp.desktop"
fi

if command -v update-desktop-database >/dev/null 2>&1; then
    update-desktop-database "$DESKTOP_DIR" >/dev/null 2>&1 || true
fi

if [ ! -f ".env" ]; then
    cp .env.exemplo .env
    echo
    echo "================================================================"
    echo "INSTALAÇÃO TÉCNICA CONCLUÍDA, MAS FALTA A CONFIGURAÇÃO."
    echo "================================================================"
    echo "O arquivo .env foi criado sem usuário e senha."
    echo "Copie o .env de outra máquina já configurada, ou preencha manualmente."
    echo
    read -r -p "Pressione Enter para fechar..."
    exit 0
fi

echo
echo "Rodando diagnóstico..."
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
echo "        VISCONDEAPP INSTALADO COM SUCESSO NO LINUX"
echo "================================================================"
echo "Abra pelo ícone \"ViscondeApp\" no menu de aplicativos (ou na Área de"
echo "Trabalho, se copiado). Na primeira vez, talvez seja preciso clicar com"
echo "o botão direito nele e escolher \"Permitir execução\"/\"Confiar\" — é uma"
echo "proteção padrão do GNOME/Nautilus contra atalhos baixados da internet."
echo
echo "Também dá pra abrir pelo terminal com: ./INICIAR_CENTRAL_LINUX.sh"
read -r -p "Pressione Enter para fechar o instalador..."
