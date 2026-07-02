#!/bin/bash
TARGET_DIR="$HOME/mobyan-automacoes"
mkdir -p "$TARGET_DIR"

echo "================================================================"
echo "MIGRAR CONFIGURAÇÃO DA CENTRAL DO WINDOWS PARA O MAC"
echo "================================================================"
echo
echo "Arraste para esta janela a pasta mobyan-automacoes do Windows"
echo "(de um pendrive, HD externo ou pasta compartilhada) e pressione Enter."
echo
read -r -p "Pasta de origem: " ORIGEM

# Remove aspas e espaços adicionados ao arrastar pelo Finder.
ORIGEM="${ORIGEM#\'}"
ORIGEM="${ORIGEM%\'}"
ORIGEM="${ORIGEM#\"}"
ORIGEM="${ORIGEM%\"}"
ORIGEM="${ORIGEM//\\ / }"

if [ ! -d "$ORIGEM" ]; then
    echo "[ERRO] Pasta não encontrada: $ORIGEM"
    read -r -p "Pressione Enter para fechar..."
    exit 1
fi

if [ ! -f "$ORIGEM/.env" ]; then
    echo "[ERRO] A pasta selecionada não contém o arquivo .env."
    read -r -p "Pressione Enter para fechar..."
    exit 1
fi

cp -p "$ORIGEM/.env" "$TARGET_DIR/.env"
chmod 600 "$TARGET_DIR/.env" 2>/dev/null || true

echo
echo "Configuração copiada com sucesso, sem mostrar as credenciais."
echo "Execute DIAGNOSTICO_MAC.command ou abra a Central Mobyan."
read -r -p "Pressione Enter para fechar..."
