"""Autoriza o acesso da automação ao OneDrive da conta (fluxo de device code do
MSAL) e confirma que o arquivo configurado em ONEDRIVE_ARQUIVO_CAMINHO é
alcançável. Rodar uma vez pelo terminal:

    .venv/bin/python app/conectar_onedrive.py

Depois disso, "Gerar Roteirização" já atualiza a planilha do OneDrive sozinho,
reaproveitando o token em cache — só é preciso rodar este script de novo se a
sessão expirar ou o acesso for revogado.
"""

import os

from dotenv import load_dotenv

from caminho_base import BASE_DIR

load_dotenv(BASE_DIR / ".env")

import onedrive_excel


def main():
    if not onedrive_excel.esta_configurado():
        print(
            "ONEDRIVE_CLIENT_ID e/ou ONEDRIVE_ARQUIVO_CAMINHO não estão configurados "
            "para esta conta. Preencha os dois em 'Configurações da conta' antes de "
            "rodar este script."
        )
        return 1

    print("Fazendo login no OneDrive (pode abrir o navegador automaticamente)...")
    token = onedrive_excel.obter_token(interativo=True)
    print("Login concluído com sucesso.")

    nome_arquivo = os.getenv("ONEDRIVE_ARQUIVO_CAMINHO", "").strip()
    print(f"Verificando acesso ao arquivo '{nome_arquivo}' em 'Compartilhado comigo'...")
    drive_id, item_id = onedrive_excel.verificar_conexao(token)
    print(f"Arquivo encontrado (drive: {drive_id}, item: {item_id}). Conexão com o OneDrive está pronta.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
