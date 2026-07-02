#!/bin/bash
PROJECT_DIR="$HOME/mobyan-automacoes"
cd "$PROJECT_DIR" || {
    echo "Não encontrei $PROJECT_DIR"
    echo "Execute INSTALAR_MAC.command primeiro."
    read -r -p "Pressione Enter para fechar..."
    exit 1
}
mkdir -p logs
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
nohup "$PROJECT_DIR/.venv/bin/python" "$PROJECT_DIR/app/central_mobyan.py" >> "$PROJECT_DIR/logs/central_mobyan.log" 2>&1 &
echo "Central Mobyan iniciada."
sleep 1
