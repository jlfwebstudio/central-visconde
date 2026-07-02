#!/usr/bin/env python3
import os
import platform
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
ERROS = []
AVISOS = []


def ok(texto):
    print(f"[OK] {texto}")


def erro(texto):
    ERROS.append(texto)
    print(f"[ERRO] {texto}")


def aviso(texto):
    AVISOS.append(texto)
    print(f"[AVISO] {texto}")


print("=" * 70)
print("DIAGNÓSTICO DA CENTRAL MOBYAN - macOS")
print("=" * 70)
print(f"Python: {sys.version.split()[0]}")
print(f"Executável: {sys.executable}")
print(f"macOS: {platform.mac_ver()[0] or 'não identificado'}")
print(f"Arquitetura: {platform.machine()}")
print(f"Projeto: {BASE_DIR}")
print()

if sys.platform != "darwin":
    aviso("Este diagnóstico foi preparado para macOS.")

if sys.version_info < (3, 9):
    erro("É necessário Python 3.9 ou superior.")
else:
    ok("Versão do Python compatível.")

modulos = [
    ("tkinter", "tkinter"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("Pillow", "PIL"),
    ("Playwright", "playwright"),
    ("python-dotenv", "dotenv"),
    ("pypdf", "pypdf"),
]

for nome, modulo in modulos:
    try:
        __import__(modulo)
        ok(f"Dependência disponível: {nome}")
    except Exception as exc:
        erro(f"Dependência ausente ou com erro: {nome} ({exc})")

arquivos = [
    "app/central_mobyan.py",
    "app/exportador_mobyan.py",
    "app/enviar_whatsapp.py",
    "app/baixar_relatorios_roteirizacao.py",
    "app/gerar_roteirizacao.py",
    "app/gerar_pdfs.py",
    "bases/base_justificativas.xlsx",
    "bases/contatos_prestadores.xlsx",
    "bases/regras_roteirizacao.xlsx",
]

for relativo in arquivos:
    caminho = BASE_DIR / relativo
    if caminho.exists():
        ok(f"Arquivo encontrado: {relativo}")
    else:
        erro(f"Arquivo não encontrado: {relativo}")

env_path = BASE_DIR / ".env"
if not env_path.exists():
    erro("Arquivo .env não encontrado. Migre-o do Windows ou configure-o.")
else:
    valores = {}
    for linha in env_path.read_text(encoding="utf-8", errors="replace").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        valores[chave.strip()] = valor.strip()

    obrigatorias = [
        "MOBYAN_URL", "MOBYAN_USUARIO", "MOBYAN_SENHA",
        "OGEA_URL", "OGEA_USUARIO", "OGEA_SENHA",
    ]
    faltantes = [chave for chave in obrigatorias if not valores.get(chave)]
    if faltantes:
        erro("Configurações vazias no .env: " + ", ".join(faltantes))
    else:
        ok("Arquivo .env configurado (valores ocultos).")

pastas = [
    "downloads/relatorios_completos",
    "downloads/roteirizacao/mobyan",
    "downloads/roteirizacao/ogea",
    "logs/roteirizacao",
    "logs/pdfs",
    "outputs/pendencias_do_dia/backups",
    "outputs/por_prestador_imagens",
    "outputs/whatsapp_temp",
    "outputs/roteirizacao",
    "outputs/pdfs",
    "whatsapp_profile",
]

for relativo in pastas:
    caminho = BASE_DIR / relativo
    try:
        caminho.mkdir(parents=True, exist_ok=True)
        teste = caminho / ".teste_escrita"
        teste.write_text("ok", encoding="utf-8")
        teste.unlink()
        ok(f"Pasta gravável: {relativo}")
    except Exception as exc:
        erro(f"Sem permissão para gravar em {relativo}: {exc}")

try:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        pagina = navegador.new_page()
        pagina.set_content("<title>Central Mobyan</title><h1>OK</h1>")
        titulo = pagina.title()
        navegador.close()
    if titulo == "Central Mobyan":
        ok("Chromium do Playwright abriu corretamente.")
    else:
        erro("O Chromium abriu, mas o teste não respondeu corretamente.")
except Exception as exc:
    erro(f"Falha ao abrir o Chromium do Playwright: {exc}")

print()
print("=" * 70)
if ERROS:
    print(f"DIAGNÓSTICO CONCLUÍDO COM {len(ERROS)} ERRO(S).")
    for item in ERROS:
        print(f"- {item}")
    raise SystemExit(1)

print("DIAGNÓSTICO CONCLUÍDO COM SUCESSO.")
if AVISOS:
    print(f"Avisos: {len(AVISOS)}")
print("=" * 70)
