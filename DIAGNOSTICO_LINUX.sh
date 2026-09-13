#!/bin/bash
cd "$(dirname "$0")" || exit 1
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
.venv/bin/python app/diagnostico_mac.py
STATUS=$?
echo
read -r -p "Pressione Enter para fechar..."
exit "$STATUS"
