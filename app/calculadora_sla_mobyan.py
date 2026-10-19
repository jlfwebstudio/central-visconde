import argparse
import os
import re
import shutil
import subprocess
import sys
import traceback
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

from caminho_base import BASE_DIR
load_dotenv(BASE_DIR / ".env")

import exportador_mobyan as mobyan

PASTA_SAIDA = BASE_DIR / "outputs" / "calculadora_sla"
PASTA_HISTORICO = PASTA_SAIDA / "historico"
PASTA_LOGS = BASE_DIR / "logs" / "calculadora_sla"
PASTA_DOWNLOADS = BASE_DIR / "downloads" / "calculadora_sla"
ARQUIVO_ATUAL = PASTA_SAIDA / "sla_mobyan_atual.xlsx"

# Colunas finais, na ordem em que devem sair na planilha (a coluna "SLA" é
# calculada à parte e entra por último — ver processar_relatorio_sla).
COLUNAS_SAIDA = [
    "Chamado",
    "Contratante",
    "Data Limite",
    "Data Atendimento",
    "Prestador",
    "Cidade",
    "Grupo de Serviço",
    "Serviço",
    "Status",
    "Justificativa do Abono",
    "Observação do Abono",
    "Abono Paytech",
    "Abono Contratante",
]

# Cada coluna de saída pode aparecer no relatório baixado com grafias
# ligeiramente diferentes (ex: "Data de Atendimento" vs "Data Atendimento",
# "Paytec" vs "Paytech") — todas comparadas já normalizadas (sem acento,
# maiúsculas, espaço único).
CANDIDATOS_COLUNAS = {
    "Chamado": ["CHAMADO"],
    "Contratante": ["CONTRATANTE"],
    "Data Limite": ["DATA LIMITE"],
    "Data Atendimento": ["DATA DE ATENDIMENTO", "DATA ATENDIMENTO"],
    "Prestador": ["PRESTADOR"],
    "Cidade": ["CIDADE"],
    "Grupo de Serviço": ["GRUPO DE SERVICO", "GRUPO SERVICO"],
    "Serviço": ["SERVICO"],
    "Status": ["STATUS"],
    "Justificativa do Abono": ["JUSTIFICATIVA DO ABONO", "JUSTIFICATIVA ABONO"],
    "Observação do Abono": ["OBSERVACAO DO ABONO", "OBSERVACAO ABONO"],
    "Abono Paytech": ["ABONO PAYTECH", "ABONO PAYTEC", "ABONADO PAYTECH", "ABONADO PAYTEC"],
    "Abono Contratante": ["ABONO CONTRATANTE", "ABONADO CONTRATANTE"],
}


def normalizar_cabecalho(valor):
    texto = "" if valor is None else str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = texto.upper()
    return re.sub(r"\s+", " ", texto).strip()


def abrir_caminho(caminho):
    caminho = str(Path(caminho).resolve())
    if os.name == "nt":
        os.startfile(caminho)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", caminho])
    else:
        subprocess.Popen(["xdg-open", caminho])


def salvar_screenshot_erro(page, nome):
    try:
        PASTA_LOGS.mkdir(parents=True, exist_ok=True)
        caminho = PASTA_LOGS / f"{nome}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        page.screenshot(path=str(caminho), full_page=True)
        print(f"Screenshot do erro salvo em: {caminho}")
    except Exception:
        pass


def calcular_periodo(hoje=None):
    """Data inicial = último dia do mês anterior; Data final = hoje. Sempre
    calculado na hora, nunca fixo."""
    hoje = hoje or date.today()
    primeiro_dia_mes_atual = hoje.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
    return ultimo_dia_mes_anterior, hoje


def montar_criteria(data_inicial, data_final):
    clausula_prestadores = " OR ".join(
        f"'Prestador' = \"{prestador}\"" for prestador in mobyan.MOBYAN_PRESTADORES
    )
    return (
        f"({clausula_prestadores}) "
        f"AND 'Data Atendimento' >= \"{data_inicial:%d/%m/%Y}\" "
        f"AND 'Data Atendimento' <= \"{data_final:%d/%m/%Y}\""
    )


def obter_frame_ativo(page, prefixo):
    """Iframe cujo name começa com "loadext-" (formulário) ou "listext-"
    (grade de resultados) — mesmo padrão usado em toda tela deste portal."""
    return page.locator(f'iframe[name^="{prefixo}"]').content_frame


def abrir_menu_oppay(page):
    """A barra de menu (Logistica/Relatórios/Roteirização/Serviços/Relatórios
    Salvos) fica recolhida numa bandeja "OPPAY MENU" no canto inferior
    esquerdo — mesmo clique por coordenada já usado em
    exportador_mobyan.abrir_relatorio_ordem_servico_sla pra abrir o menu de
    Relatórios na geração de pendências."""
    page.wait_for_timeout(3000)
    page.mouse.click(55, 748)
    page.wait_for_timeout(1500)


def baixar_relatorio_sla(data_inicial, data_final):
    """Navegação gravada ao vivo com playwright codegen em 2026-08-06
    (Serviços > Ordem de Serviço > Pesquisar > Advanced Search > #VALUECRITERIA
    > Pesquisar > Exportar CSV) — o botão de exportação real dessa tela é
    "Exportar CSV" (mesmo usado no resto do sistema), não um XLSX separado."""
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    caminho_saida = PASTA_DOWNLOADS / "relatorio_sla_mobyan.csv"
    criteria = montar_criteria(data_inicial, data_final)

    mobyan.validar_configuracoes()

    print("")
    print("=" * 70)
    print("BAIXANDO RELATÓRIO DE SLA DA MOBYAN")
    print(f"Período: {data_inicial:%d/%m/%Y} a {data_final:%d/%m/%Y}")
    print("=" * 70)

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        # Mesmo viewport usado em todos os outros scripts Mobyan — o clique que
        # abre a bandeja "OPPAY MENU" é por coordenada fixa de pixel (ver
        # abrir_menu_oppay), então precisa bater com a posição real da barra.
        contexto = navegador.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 768},
        )
        page = contexto.new_page()
        page.set_default_timeout(60000)
        page.set_default_navigation_timeout(60000)

        try:
            print("Abrindo sistema da Mobyan...")
            page.goto(mobyan.MOBYAN_URL, wait_until="domcontentloaded", timeout=60000)
            mobyan.fazer_login(page)
            mobyan.finalizar_sessao_anterior_se_aparecer(page)
            page.wait_for_timeout(4000)

            print("Abrindo OPPAY MENU...")
            abrir_menu_oppay(page)

            print("Abrindo Serviços > Ordem de Serviço...")
            try:
                page.get_by_role("menuitem", name="Serviços").click(timeout=4000)
                page.wait_for_timeout(1000)
                page.get_by_role("menuitem", name="Ordem de Serviço", exact=True).click(timeout=5000)
            except Exception:
                salvar_screenshot_erro(page, "erro_menu_servicos")
                raise Exception('Falha ao abrir "Serviços > Ordem de Serviço" (ver screenshot em logs/calculadora_sla).')
            page.wait_for_timeout(3000)

            frame = obter_frame_ativo(page, "loadext-")

            print("Pesquisando (habilita a busca avançada)...")
            frame.get_by_role("img", name="Pesquisar").click()
            page.wait_for_timeout(1500)

            print("Abrindo busca avançada (Advanced Search)...")
            try:
                frame.get_by_role("img", name="Advanced Search").click()
            except Exception:
                salvar_screenshot_erro(page, "erro_botao_advanced_search")
                raise
            page.wait_for_timeout(1000)

            print("Preenchendo critério de busca...")
            campo_criteria = frame.locator("#VALUECRITERIA")
            campo_criteria.click()
            campo_criteria.fill(criteria)

            print("Pesquisando...")
            frame.get_by_role("img", name="Pesquisar").click()

            print("Aguardando resultado...")
            page.wait_for_timeout(8000)

            frame_lista = obter_frame_ativo(page, "listext-")

            print("Exportando relatório (CSV)...")
            with page.expect_download(timeout=60000) as download_info:
                frame_lista.get_by_role("img", name="Exportar CSV").click()
            download_info.value.save_as(caminho_saida)

            print(f"Relatório de SLA salvo em: {caminho_saida}")
            return caminho_saida

        except Exception:
            salvar_screenshot_erro(page, "erro_calculadora_sla")
            raise
        finally:
            contexto.close()
            navegador.close()


def ler_relatorio_baixado(caminho):
    """O botão real desta tela exporta CSV (mesmo padrão do resto do
    sistema — ver mobyan.ler_relatorio_csv), mas aceita .xlsx/.xls também
    pra quem usar --arquivo com algo exportado manualmente."""
    caminho = Path(caminho)
    if caminho.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        return pd.read_excel(caminho, dtype=str)
    return mobyan.ler_relatorio_csv(caminho)


def processar_relatorio_sla(caminho_relatorio, hoje):
    df_bruto = ler_relatorio_baixado(caminho_relatorio)

    cabecalhos_normalizados = {normalizar_cabecalho(c): c for c in df_bruto.columns}
    mapa = {}
    faltando = []
    for alvo, candidatos in CANDIDATOS_COLUNAS.items():
        origem = next((cabecalhos_normalizados[c] for c in candidatos if c in cabecalhos_normalizados), None)
        if origem is None:
            faltando.append(alvo)
        else:
            mapa[alvo] = origem

    if faltando:
        raise ValueError(
            "O relatório baixado não tem todas as colunas esperadas: "
            + ", ".join(faltando)
            + ". Colunas encontradas no arquivo: "
            + ", ".join(str(c) for c in df_bruto.columns)
        )

    df = df_bruto[[mapa[coluna] for coluna in COLUNAS_SAIDA]].copy()
    df.columns = COLUNAS_SAIDA

    for coluna in ("Data Limite", "Data Atendimento"):
        df[coluna] = pd.to_datetime(df[coluna], dayfirst=True, errors="coerce").dt.date

    total_bruto = len(df)

    # 5a: bobina não conta pro SLA (mesmo critério já usado no resto da Central).
    df = mobyan.excluir_bobina(df)

    # 5b: "Retirada" (mesmo fora de bobina) não conta. Desinstalação continua.
    df = df[df["Serviço"].apply(lambda v: mobyan.normalizar_texto(v) != "RETIRADA")]

    # 5c: só o mês atual conta, mesmo que o relatório traga o último dia do
    # mês anterior (isso existe só pra pegar OS com Data Limite no fim do mês
    # passado atendida já em agosto, por exemplo).
    primeiro_dia_mes_atual = hoje.replace(day=1)
    df = df[df["Data Atendimento"].apply(lambda d: d is not None and d >= primeiro_dia_mes_atual)]

    # 5d: só OS Baixada entra no cálculo de SLA.
    df = df[df["Status"].apply(lambda v: mobyan.normalizar_texto(v) == "BAIXADA")]

    df["SLA"] = df.apply(
        lambda linha: (
            "Atraso"
            if linha["Data Atendimento"] and linha["Data Limite"] and linha["Data Atendimento"] > linha["Data Limite"]
            else "No prazo"
        ),
        axis=1,
    )

    print(f"Relatório bruto: {total_bruto} linha(s). Após os filtros: {len(df)} linha(s).")

    return df.reset_index(drop=True)


def formatar_planilha_sla(caminho):
    wb = load_workbook(caminho)
    ws = wb["SLA"]
    ws.freeze_panes = "A2"

    fonte_cabecalho = Font(bold=True, color="FFFFFF")
    preenchimento_cabecalho = PatternFill("solid", fgColor="1F1F1F")
    for celula in ws[1]:
        celula.font = fonte_cabecalho
        celula.fill = preenchimento_cabecalho
        celula.alignment = Alignment(horizontal="center")

    cabecalhos = [celula.value for celula in ws[1]]

    if "SLA" in cabecalhos:
        col_sla = cabecalhos.index("SLA") + 1
        for linha in range(2, ws.max_row + 1):
            celula = ws.cell(row=linha, column=col_sla)
            if celula.value == "Atraso":
                celula.font = Font(color="9C0006", bold=True)
            elif celula.value == "No prazo":
                celula.font = Font(color="006100", bold=True)

    for coluna in ("Data Limite", "Data Atendimento"):
        if coluna in cabecalhos:
            indice = cabecalhos.index(coluna) + 1
            for linha in range(2, ws.max_row + 1):
                ws.cell(row=linha, column=indice).number_format = "DD/MM/YYYY"

    larguras = {
        "Chamado": 14, "Contratante": 18, "Data Limite": 14, "Data Atendimento": 16,
        "Prestador": 30, "Cidade": 18, "Grupo de Serviço": 20, "Serviço": 20,
        "Status": 14, "Justificativa do Abono": 26, "Observação do Abono": 26,
        "Abono Paytech": 16, "Abono Contratante": 18, "SLA": 12,
    }
    for indice, cabecalho in enumerate(cabecalhos, start=1):
        ws.column_dimensions[get_column_letter(indice)].width = larguras.get(cabecalho, 16)

    wb.save(caminho)


def salvar_planilha_sla(df, hoje):
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    PASTA_HISTORICO.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ARQUIVO_ATUAL, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="SLA")

    formatar_planilha_sla(ARQUIVO_ATUAL)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_historico = PASTA_HISTORICO / f"sla_mobyan_{timestamp}.xlsx"
    shutil.copy2(ARQUIVO_ATUAL, caminho_historico)

    return ARQUIVO_ATUAL


def main():
    parser = argparse.ArgumentParser(description="Calculadora de SLA Mobyan")
    parser.add_argument(
        "--arquivo",
        default="",
        help="Usa um .csv ou .xlsx já exportado manualmente da Mobyan em vez de baixar pela automação",
    )
    parser.add_argument("--sem-abrir", action="store_true", help="Não abre a planilha ao terminar")
    args = parser.parse_args()

    hoje = date.today()
    data_inicial, data_final = calcular_periodo(hoje)

    try:
        if args.arquivo:
            caminho_baixado = Path(args.arquivo).expanduser().resolve()
        else:
            caminho_baixado = baixar_relatorio_sla(data_inicial, data_final)

        df_final = processar_relatorio_sla(caminho_baixado, hoje)
        caminho_saida = salvar_planilha_sla(df_final, hoje)

        no_prazo = int((df_final["SLA"] == "No prazo").sum())
        atraso = int((df_final["SLA"] == "Atraso").sum())

        print("")
        print("=" * 70)
        print(f"OSs consideradas no SLA: {len(df_final)}")
        print(f"No prazo: {no_prazo} | Atraso: {atraso}")
        print(f"Planilha: {caminho_saida}")
        print("=" * 70)

        if not args.sem_abrir:
            abrir_caminho(caminho_saida)

        return 0
    except Exception as erro:
        print("")
        print("ERRO AO CALCULAR SLA MOBYAN")
        print(str(erro))
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
