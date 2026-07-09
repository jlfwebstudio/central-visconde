# -*- mode: python ; coding: utf-8 -*-
"""Build (onedir) do ViscondeApp para distribuição a clientes, sem
depender de Python instalado na máquina de destino.

Uso: .venv/bin/pyinstaller central_visconde.spec (limpa builds antigas com
--noconfirm se necessário). Requer requirements-build.txt instalado.
"""

import sys
from pathlib import Path

RAIZ = Path(SPECPATH)
NOME_APP = "ViscondeApp"

# Módulos hoje só referenciados por caminho de arquivo (subprocess), nunca por
# `import` estático — o PyInstaller não os descobre sozinho, então precisam
# entrar aqui manualmente. São os mesmos nomes usados pelo shim de despacho em
# central_mobyan.py (--rodar-automacao <nome_do_modulo>).
HIDDEN_IMPORTS = [
    "exportador_mobyan",
    "enviar_whatsapp",
    "gerar_roteirizacao",
    "gerar_pdfs",
    "analisar_abonos_ogea",
    "gerar_os_visconde_teste",
    "baixar_relatorios_roteirizacao",
    "configurar_base_justificativas",
    "configurar_contatos_prestadores",
]

ICONE = str(RAIZ / "assets" / ("logo_visconde.icns" if sys.platform == "darwin" else "logo_visconde.ico"))

a = Analysis(
    [str(RAIZ / "app" / "central_mobyan.py")],
    pathex=[str(RAIZ / "app")],
    binaries=[],
    datas=[(str(RAIZ / "assets"), "assets")],
    hiddenimports=HIDDEN_IMPORTS,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=NOME_APP,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=ICONE,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name=NOME_APP,
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name=f"{NOME_APP}.app",
        icon=ICONE,
        # Mantido igual ao já publicado (não segue o rebrand pra "ViscondeApp"):
        # trocar isso faz o macOS tratar como um app diferente do já instalado,
        # perdendo permissões concedidas em máquinas que já rodam o app.
        bundle_identifier="com.centralvisconde.app",
        info_plist={
            "CFBundleShortVersionString": "1.1.0",
            "LSMinimumSystemVersion": "11.0",
            "NSHighResolutionCapable": True,
        },
    )
