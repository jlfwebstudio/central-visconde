#!/bin/bash
PROJECT_DIR="$HOME/mobyan-automacoes"
cd "$PROJECT_DIR" || {
    echo "Não encontrei $PROJECT_DIR"
    read -r -p "Pressione Enter para fechar..."
    exit 1
}
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
"$PROJECT_DIR/.venv/bin/python" "$PROJECT_DIR/app/diagnostico_mac.py"
STATUS=$?
echo
read -r -p "Pressione Enter para fechar..."
exit "$STATUS"
