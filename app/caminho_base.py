"""Resolve os caminhos-base usados por todos os scripts do ViscondeApp.

Rodando a partir do código-fonte (`.venv/bin/python app/algum_script.py`),
BASE_DIR e RECURSOS_DIR são idênticos e iguais à raiz do repositório — exatamente
o que `Path(__file__).resolve().parent.parent` já dava antes deste módulo existir.

Dentro de um executável empacotado (PyInstaller), os dois passam a apontar pra
lugares diferentes:
- BASE_DIR: pasta persistente por usuário (sobrevive a atualizações do app) onde
  moram `.env`, `config/`, `outputs/`, `downloads/`, `logs/`, `bases/` e
  `whatsapp_profile/` — dados de execução, não arquivos do próprio app.
- RECURSOS_DIR: pasta só-leitura que vem junto do executável (assets/ícones e os
  próprios módulos .py usados pelo shim de despacho de subprocessos) — muda só
  quando o app é atualizado.
"""

import os
import sys
from pathlib import Path

_RAIZ_REPO = Path(__file__).resolve().parent.parent

FROZEN = getattr(sys, "frozen", False)

if FROZEN:
    # Nome da pasta mantido igual ao já publicado mesmo após o rebrand pra
    # "ViscondeApp" — trocar isso faria toda instalação já em uso "perder" seus
    # dados (login, credenciais, regras, sessão do WhatsApp) na próxima
    # atualização, já que o app abriria olhando pra uma pasta nova e vazia.
    if sys.platform == "darwin":
        BASE_DIR = Path.home() / "Library" / "Application Support" / "Central Visconde"
    elif sys.platform == "win32":
        BASE_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "Central Visconde"
    else:
        BASE_DIR = Path.home() / ".central_visconde"

    RECURSOS_DIR = Path(getattr(sys, "_MEIPASS", None) or Path(sys.executable).resolve().parent)
else:
    BASE_DIR = _RAIZ_REPO
    RECURSOS_DIR = _RAIZ_REPO

BASE_DIR.mkdir(parents=True, exist_ok=True)

if FROZEN:
    # Sem isso, o Playwright resolve "onde procurar o Chromium já instalado"
    # de um jeito e "onde o instalador baixou" de outro dentro de um
    # executável empacotado (visto em teste real no Windows: o download foi
    # pra %LOCALAPPDATA%\ms-playwright, mas o Python procurava dentro da
    # própria pasta do app). Fixando esse caminho explicitamente, os dois
    # lados usam sempre o mesmo lugar.
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(BASE_DIR / "navegadores_playwright"))
