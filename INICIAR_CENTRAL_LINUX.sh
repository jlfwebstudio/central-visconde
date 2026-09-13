#!/bin/bash
cd "$(dirname "$0")" || exit 1
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
.venv/bin/python app/central_mobyan.py
STATUS=$?
echo
if [ "$STATUS" -ne 0 ]; then
    echo "A Central terminou com erro (código $STATUS)."
fi
read -r -p "Pressione Enter para fechar..."
