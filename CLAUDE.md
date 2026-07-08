# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Tkinter desktop app ("Central Mobyan" / "Central Visconde" — the product is mid-rebrand from Mobyan to Visconde, so both names appear in code, filenames, and UI strings) that automates the daily field-service dispatch workflow for a Brazilian company: scraping pending service orders from a portal, routing them to technicians, generating printable PDF route sheets, and notifying providers over WhatsApp. All identifiers, comments, and UI text are in Portuguese (pt-BR) — keep new code consistent with that.

The app drives two external systems with Playwright:
- **Mobyan** — the primary service-order portal (`MOBYAN_URL`).
- **OGEA / Workfinity** (`tefti.workfinity.com.br`) — a second service-order portal, referred to as "OGEA" throughout.

It is distributed as a self-contained macOS bundle (`.command` launcher scripts + a local `.venv` + a generated `~/Applications/Central Mobyan.app` wrapper), though a fair amount of Windows-compatibility code remains (`os.name == "nt"` branches, `winsound`, `WINDIR` font lookups) from an earlier Windows-only version.

## Setup and running

There is no package manager beyond pip + a plain `venv`; no test suite, linter, or CI config exists in this repo.

```bash
# First-time setup (macOS): creates .venv, installs deps + Playwright Chromium, creates folders
./INSTALAR_MAC.command

# Manual equivalent
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m playwright install chromium

# Run the GUI
.venv/bin/python app/central_mobyan.py
# or: ./INICIAR_CENTRAL_MAC.command (foreground) / INICIAR_CENTRAL.command (background, logs to logs/central_mobyan.log)

# Environment/dependency health check (also run automatically at the end of install)
.venv/bin/python app/diagnostico_mac.py

# Run any automation script standalone from the CLI (most support this — see "Entry points" below)
.venv/bin/python app/gerar_roteirizacao.py --sem-abrir
```

Configuration lives in `.env` (copy from `.env.exemplo`), read independently by every entry-point script via `load_dotenv(BASE_DIR / ".env")` — there is no shared config module. Key vars: `MOBYAN_URL/USUARIO/SENHA`, `OGEA_URL/USUARIO/SENHA/ORDEM_SERVICO_URL/DIAS_ABERTURA/TIMEOUT_CARREGAMENTO`, `PDF_TIMEOUT_SEGUNDOS`. `.env`, `.venv/`, `whatsapp_profile/`, `downloads/`, `logs/`, `outputs/`, `bases/`, and `backups/` are all gitignored — they hold runtime data/credentials, not source.

`app/gerar_os_visconde_teste.py` additionally requires `reportlab`, which is **not** in `requirements.txt` — it's an optional dependency only needed for that one pilot feature.

## Architecture

### GUI hub and process model

`app/central_mobyan.py` is the Tkinter entry point. It does **not** import most automation modules directly — it launches each as a separate subprocess (`subprocess.Popen([PYTHON_AUTOMACOES, script_path, *args], cwd=BASE_DIR, ...)`), streams stdout into the GUI's log panel via a `queue.Queue` fed by a background thread, and surfaces success/failure based on the child's exit code. The one exception is `app/gestao_rotas.py`, which is imported directly (`from gestao_rotas import abrir_gestao_rotas, obter_resumo_rotas`) so it can open as an embedded Tkinter sub-window and share state with the main GUI.

Buttons in the GUI map to these scripts (`central_mobyan.py`'s `rodar_script`/`SCRIPT_*` constants):

| GUI action | Script |
|---|---|
| Gerar Pendências | `app/exportador_mobyan.py` |
| Enviar WhatsApp | `app/enviar_whatsapp.py` |
| Gerar Roteirização | `app/gerar_roteirizacao.py --sem-abrir` |
| Gerar Roteiros PDF | `app/gerar_pdfs.py` |
| Gerar OS Visconde - TESTE | `app/gerar_os_visconde_teste.py` |
| Analisar Abonos OGEA | `app/analisar_abonos_ogea.py --arquivo <path>` |

### Daily pipeline and data flow

The documented daily routine (`LEIA-ME_MAC.txt`) is: **Gerar Roteirização → Gerar PDFs → print the per-technician PDFs**. Full data flow:

1. **`exportador_mobyan.py`** — logs into the Mobyan portal via Playwright, scrapes pending service orders, cross-references `bases/base_justificativas.xlsx` (justification rules) and `bases/contatos_prestadores.xlsx` (provider contacts) → writes `outputs/pendencias_do_dia/pendencias_do_dia_atual.xlsx` (with an `Envios` sheet consumed by the WhatsApp step) and provider-summary images to `outputs/por_prestador_imagens/`.
2. **`enviar_whatsapp.py`** — reads the `Envios` sheet of that same workbook, sends WhatsApp messages via Playwright against a **persistent browser profile at `whatsapp_profile/`** (so the WhatsApp Web session/QR-login survives between runs). Message template ("manhã" vs "acompanhamento") is chosen via the `MODELO_MENSAGEM` env var, which `central_mobyan.py` sets per-run based on a GUI prompt.
3. **`baixar_relatorios_roteirizacao.py`** — downloads Mobyan + OGEA CSV reports via Playwright into `downloads/roteirizacao/{mobyan,ogea}/`. Not launched directly from the GUI; imported by both scripts below.
4. **`gerar_roteirizacao.py`** — imports `baixar_relatorios_automaticamente` from step 3, merges the two reports against `bases/regras_roteirizacao.xlsx` (per-city/technician routing rules) and any pending resolutions, and writes `outputs/roteirizacao/roteirizacao_atual.xlsx` plus a dated copy under `outputs/roteirizacao/historico/`. CLI flags: `--mobyan`, `--ogea` (use existing CSVs instead of downloading), `--sem-abrir`, `--reprocessar-atual`, `--usar-resolucoes-temporarias`.
5. **Route resolution** — if the generated routing has orders with no matching rule or conflicting rules, `central_mobyan.py` (`verificar_pendencias_rotas_automaticamente` / `exigir_rotas_resolvidas`) blocks the PDF step and opens `gestao_rotas.py`'s embedded window so the user can resolve them manually. `gestao_rotas.py` writes decisions back into `bases/regras_roteirizacao.xlsx` (backing up the previous version to `bases/backups_roteirizacao/`) and `outputs/roteirizacao/resolucoes_temporarias.json`, and can re-invoke `gerar_roteirizacao.py` as a subprocess to regenerate the route.
6. **`gerar_pdfs.py`** — imports `fazer_login`/`validar_configuracoes` from `exportador_mobyan.py` and the OGEA helpers (`fazer_login_ogea`, `abrir_ordem_servico_ogea`, etc.) from `baixar_relatorios_roteirizacao.py` **directly** (not subprocess), drives Playwright to render/download the individual Mobyan + OGEA service-order PDFs referenced in `roteirizacao_atual.xlsx`, then merges them per technician with `pypdf` into `outputs/pdfs/Roteiro - <Nome>.pdf`. CLI flags: `--sem-abrir`, `--somente-ogea`/`--somente-mobyan` (mutually exclusive), `--tecnico <nome>`.
7. **`gerar_os_visconde_teste.py`** — a parallel/pilot output track using the same `roteirizacao_atual.xlsx` but generating an alternate Visconde-branded PDF layout via `reportlab` into `outputs/os_visconde_teste/`. Explicitly non-destructive: it never touches the official `roteirizacao_atual.xlsx` or `outputs/pdfs/` (confirmed by its own GUI confirmation dialog).
8. **`analisar_abonos_ogea.py`** — a standalone tool (not part of the linear pipeline): the user picks a downloaded OGEA report via a file dialog, the script applies regex patterns (Saturday-service keyword detection) plus `bases/regras_abonos_ogea.json` to decide "abono" (allowance) eligibility per order, and writes `outputs/abonos_ogea/analise_abonos_ogea_atual.xlsx` + a timestamped history copy + a run log under `logs/abonos_ogea/`.

### Scripts not wired into the GUI

These exist under `app/` but are run manually from the terminal for setup/maintenance, not launched by `central_mobyan.py`:
- `configurar_base_justificativas.py`, `configurar_contatos_prestadores.py` — build/edit the `bases/*.xlsx` reference workbooks used by `exportador_mobyan.py`.
- `migrar_base_rotas.py` — one-off migration helper for `bases/regras_roteirizacao.xlsx`.
- `teste_navegador.py` — minimal Playwright smoke test.
- `diagnostico_mac.py` — environment diagnostic (its own launcher, `DIAGNOSTICO_MAC.command`); checks Python version, required packages, presence of core files/`.env` vars, folder write permissions, and that headless Chromium launches.
- `gerar_imagens_prestadores.py` — an older/alternate generator for the `outputs/por_prestador_imagens/` images; the same logic now also lives inline inside `exportador_mobyan.py` (near-duplicate column/style constants) — check which one is actually current before modifying either.
- `gerar_pdfs_antes_v8.py` — the pre-v8 version of `gerar_pdfs.py`, kept for reference/rollback only.

### Conventions to follow

- Every script computes `BASE_DIR = Path(__file__).resolve().parent.parent` and reads/writes paths relative to it (`bases/`, `outputs/`, `downloads/`, `logs/`, `assets/`) — don't hardcode paths or assume CWD.
- Playwright is used via the **sync API** everywhere; follow that rather than introducing async.
- On automation failure, scripts save a screenshot via a local `salvar_screenshot_erro(page, nome)` helper to `logs/<feature>/erro_<nome>_<timestamp>.png` — replicate this pattern in new Playwright flows rather than failing silently.
- Text/column normalization for matching against scraped data consistently strips accents and uppercases via a local `normalizar_texto()` helper (NFKD decompose + drop combining marks + uppercase) — each module defines its own copy; match that behavior if you add new matching logic instead of doing naive string comparison.
- Scripts intended to be launched from the GUI take `argparse` flags (see tables above) and must still work when run standalone from the CLI with no args — preserve this dual usage.
- The `backups/` directory contains manual timestamped snapshots of key files (`central_mobyan.py`, `analisar_abonos_ogea.py`, etc.) made outside of git by the previous (non-version-controlled) workflow. This repo currently has a single commit; treat `backups/` as legacy reference, not something to keep updating — git history is now the source of truth for change tracking.
