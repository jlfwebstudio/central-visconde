#!/bin/bash
cd "$(dirname "$0")"
source .venv/bin/activate
python app/central_mobyan.py
STATUS=$?
echo
if [ $STATUS -ne 0 ]; then
  echo "A Central terminou com erro (código $STATUS)."
fi
read -r -p "Pressione Enter para fechar..."
