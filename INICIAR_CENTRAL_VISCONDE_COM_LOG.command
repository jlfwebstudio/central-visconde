#!/bin/bash
cd "$HOME/mobyan-automacoes"
"$HOME/mobyan-automacoes/.venv/bin/python" "$HOME/mobyan-automacoes/app/central_mobyan.py"
echo
read -n 1 -s -r -p "Pressione qualquer tecla para fechar..."
