import os
import re
import shutil
import time
import unicodedata
import zipfile
from datetime import datetime, timedelta

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright


from caminho_base import BASE_DIR

PASTA_DOWNLOADS_ROTEIRO = BASE_DIR / "downloads" / "roteirizacao"
PASTA_DOWNLOADS_MOBYAN = PASTA_DOWNLOADS_ROTEIRO / "mobyan"
PASTA_DOWNLOADS_OGEA = PASTA_DOWNLOADS_ROTEIRO / "ogea"
PASTA_LOGS_ROTEIRO = BASE_DIR / "logs" / "roteirizacao"

load_dotenv(BASE_DIR / ".env")

OGEA_URL = os.getenv(
    "OGEA_URL",
    "https://tefti.workfinity.com.br/login/auth",
).strip()
OGEA_USUARIO = os.getenv("OGEA_USUARIO", "").strip()
OGEA_SENHA = os.getenv("OGEA_SENHA", "").strip()
OGEA_ORDEM_SERVICO_URL = os.getenv(
    "OGEA_ORDEM_SERVICO_URL",
    "https://tefti.workfinity.com.br/serviceOrder/index",
).strip()

try:
    OGEA_DIAS_ABERTURA = int(os.getenv("OGEA_DIAS_ABERTURA", "30"))
except ValueError:
    OGEA_DIAS_ABERTURA = 30

try:
    OGEA_TIMEOUT_CARREGAMENTO = int(
        os.getenv("OGEA_TIMEOUT_CARREGAMENTO", "300")
    )
except ValueError:
    OGEA_TIMEOUT_CARREGAMENTO = 300

# Preservados para não quebrar a instalação atual (RS-SMART/SMART TECNOLOGIA):
# contas novas configuram os próprios valores via MOBYAN_PRESTADORES/OGEA_PRESTADOR.
MOBYAN_PRESTADOR_PRINCIPAL = os.getenv("MOBYAN_PRESTADORES", "RS-SMART").split(",")[0].strip() or "RS-SMART"
OGEA_PRESTADOR = os.getenv("OGEA_PRESTADOR", "SMART TECNOLOGIA").strip() or "SMART TECNOLOGIA"


def normalizar_texto(valor):
    texto = "" if valor is None else str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = texto.upper()
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def preparar_pastas():
    for pasta in [
        PASTA_DOWNLOADS_MOBYAN,
        PASTA_DOWNLOADS_OGEA,
        PASTA_LOGS_ROTEIRO,
    ]:
        pasta.mkdir(parents=True, exist_ok=True)


def limpar_downloads_roteirizacao():
    """Remove relatórios antigos antes de iniciar uma nova extração."""
    preparar_pastas()

    for pasta in [PASTA_DOWNLOADS_MOBYAN, PASTA_DOWNLOADS_OGEA]:
        for item in pasta.iterdir():
            try:
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()
            except Exception as erro:
                raise RuntimeError(
                    f"Não consegui remover o arquivo antigo: {item}. "
                    f"Feche qualquer arquivo aberto e tente novamente. Erro: {erro}"
                ) from erro

    print("Downloads antigos da roteirização removidos.")


def salvar_screenshot_erro(page, nome):
    try:
        preparar_pastas()
        caminho = PASTA_LOGS_ROTEIRO / (
            f"{nome}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        )
        page.screenshot(path=str(caminho), full_page=True)
        print(f"Screenshot do erro salvo em: {caminho}")
    except Exception:
        pass


def primeiro_visivel(locators):
    for locator in locators:
        try:
            if locator.count() > 0 and locator.first.is_visible():
                return locator.first
        except Exception:
            continue
    return None


def aguardar_pagina_ogea(
    page,
    descricao,
    timeout_segundos=None,
    exigir_sem_carregamento=True,
):
    """
    Aguarda a página do OGEA ficar realmente utilizável.

    Não espera 5 minutos obrigatoriamente:
    - se a página carregar rápido, avança imediatamente;
    - se estiver lenta, aguarda até o limite configurado;
    - só gera erro após o tempo máximo.
    """
    if timeout_segundos is None:
        timeout_segundos = OGEA_TIMEOUT_CARREGAMENTO

    limite = time.time() + timeout_segundos
    proximo_aviso = time.time()
    ultimo_estado = ""

    seletores_carregamento = [
        ".loading:visible",
        ".loader:visible",
        ".spinner:visible",
        ".spinner-border:visible",
        ".progress:visible",
        ".blockUI:visible",
        ".ui-widget-overlay:visible",
        ".modal-backdrop:visible",
        "[aria-busy='true']:visible",
        "[class*='loading']:visible",
        "[class*='spinner']:visible",
    ]

    while time.time() < limite:
        try:
            estado = page.evaluate("document.readyState")
        except Exception:
            estado = ""

        corpo_ok = False
        try:
            corpo_ok = page.locator("body").count() > 0
        except Exception:
            pass

        carregando = 0

        if exigir_sem_carregamento:
            for seletor in seletores_carregamento:
                try:
                    carregando += page.locator(seletor).count()
                except Exception:
                    continue

        pronto = (
            estado in {"interactive", "complete"}
            and corpo_ok
            and (not exigir_sem_carregamento or carregando == 0)
        )

        if pronto:
            # Pequena folga para scripts e componentes terminarem de montar.
            page.wait_for_timeout(1200)

            try:
                estado_final = page.evaluate("document.readyState")
            except Exception:
                estado_final = estado

            if estado_final in {"interactive", "complete"}:
                print(f"{descricao}: página carregada.")
                return

        agora = time.time()

        if agora >= proximo_aviso:
            restante = max(0, int(limite - agora))
            estado_texto = estado or "indefinido"
            mensagem = (
                f"{descricao}: aguardando carregamento... "
                f"estado={estado_texto}, elementos carregando={carregando}, "
                f"tempo restante={restante}s"
            )

            if mensagem != ultimo_estado:
                print(mensagem)
                ultimo_estado = mensagem

            proximo_aviso = agora + 15

        page.wait_for_timeout(1500)

    salvar_screenshot_erro(page, "timeout_carregamento_ogea")
    raise TimeoutError(
        f"{descricao} não terminou de carregar em "
        f"{timeout_segundos} segundos."
    )


# ---------------------------------------------------------------------------
# MOBYAN
# ---------------------------------------------------------------------------

def selecionar_apenas_rs_smart(frame_relatorio, prestador=None):
    prestador = prestador or MOBYAN_PRESTADOR_PRINCIPAL
    print(f"Selecionando somente o prestador {prestador}...")

    frame_relatorio.get_by_role(
        "cell",
        name="Prestador",
    ).get_by_role("button").click()

    marcado = False

    for tentativa in [
        lambda: frame_relatorio.get_by_role(
            "checkbox",
            name=prestador,
            exact=True,
        ).check(timeout=5000),
        lambda: frame_relatorio.locator("label").filter(
            has_text=re.compile(rf"^{re.escape(prestador)}$")
        ).click(timeout=5000),
    ]:
        try:
            tentativa()
            marcado = True
            break
        except Exception:
            continue

    if not marcado:
        raise RuntimeError(
            f"Não consegui marcar somente o prestador {prestador} na Mobyan."
        )

    try:
        frame_relatorio.locator("html").click()
    except Exception:
        pass

    print(f"Prestador marcado: {prestador}")


def baixar_relatorio_mobyan():
    """
    Reaproveita o fluxo validado do exportador de pendências,
    alterando apenas a seleção dos prestadores para RS-SMART.
    """
    preparar_pastas()

    try:
        import exportador_mobyan as mobyan
    except Exception as erro:
        raise RuntimeError(
            "Não consegui importar app/exportador_mobyan.py. "
            f"Erro: {erro}"
        ) from erro

    mobyan.validar_configuracoes()

    caminho_saida = (
        PASTA_DOWNLOADS_MOBYAN
        / "relatorio_mobyan_atual.csv"
    )

    print("")
    print("=" * 70)
    print("BAIXANDO RELATÓRIO DA MOBYAN")
    print("=" * 70)

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=False)
        contexto = navegador.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 768},
        )
        page = contexto.new_page()
        page.set_default_timeout(OGEA_TIMEOUT_CARREGAMENTO * 1000)
        page.set_default_navigation_timeout(
            OGEA_TIMEOUT_CARREGAMENTO * 1000
        )

        try:
            print("Abrindo sistema da Mobyan...")
            page.goto(
                mobyan.MOBYAN_URL,
                wait_until="domcontentloaded",
                timeout=60000,
            )

            mobyan.fazer_login(page)
            mobyan.finalizar_sessao_anterior_se_aparecer(page)

            print("Aguardando entrada no sistema...")
            page.wait_for_timeout(5000)

            mobyan.abrir_relatorio_ordem_servico_sla(page)
            page.wait_for_timeout(5000)

            frame_relatorio = page.locator(
                'iframe[name^="loadext-"]'
            ).content_frame

            # O relatório pode trazer outros status.
            # O motor filtra depois somente os textos contendo ENCAMINH.
            mobyan.selecionar_status(frame_relatorio)
            selecionar_apenas_rs_smart(frame_relatorio)
            mobyan.selecionar_estado_rs(frame_relatorio)

            print("Pesquisando relatório Mobyan...")
            frame_relatorio.get_by_role(
                "img",
                name="Pesquisar",
            ).click()

            print("Aguardando resultado da Mobyan...")
            page.wait_for_timeout(8000)

            frame_lista = page.locator(
                'iframe[name^="listext-"]'
            ).content_frame

            print("Exportando CSV da Mobyan...")
            with page.expect_download(timeout=60000) as download_info:
                frame_lista.get_by_role(
                    "img",
                    name="Exportar CSV",
                ).click()

            download_info.value.save_as(caminho_saida)
            print(f"Relatório Mobyan salvo em: {caminho_saida}")

        except Exception:
            salvar_screenshot_erro(page, "erro_download_mobyan")
            raise
        finally:
            contexto.close()
            navegador.close()

    return caminho_saida


# ---------------------------------------------------------------------------
# OGEA
# ---------------------------------------------------------------------------

def validar_ogea():
    faltantes = []

    if not OGEA_URL:
        faltantes.append("OGEA_URL")
    if not OGEA_USUARIO:
        faltantes.append("OGEA_USUARIO")
    if not OGEA_SENHA:
        faltantes.append("OGEA_SENHA")

    if faltantes:
        raise ValueError(
            "Configuração do OGEA ausente no arquivo .env: "
            + ", ".join(faltantes)
        )


def esperar_elemento_ogea(locator, descricao, timeout_segundos=None):
    if timeout_segundos is None:
        timeout_segundos = OGEA_TIMEOUT_CARREGAMENTO

    print(f"Aguardando: {descricao}...")

    try:
        locator.wait_for(
            state="visible",
            timeout=timeout_segundos * 1000,
        )
    except Exception as erro:
        raise TimeoutError(
            f"{descricao} não apareceu dentro de "
            f"{timeout_segundos} segundos."
        ) from erro

    return locator


def fazer_login_ogea(page):
    print("Abrindo OGEA...")

    page.goto(
        OGEA_URL,
        wait_until="domcontentloaded",
        timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000,
    )

    campo_login = page.get_by_role("textbox", name="Login")
    campo_senha = page.get_by_role("textbox", name="Senha")

    # Caso já exista uma sessão ativa, a tela de login pode não aparecer.
    if campo_login.count() == 0 or campo_senha.count() == 0:
        print("Tela de login não apareceu. Sessão do OGEA possivelmente ativa.")
        return

    esperar_elemento_ogea(
        campo_login,
        "campo Login do OGEA",
    )
    esperar_elemento_ogea(
        campo_senha,
        "campo Senha do OGEA",
    )

    print("Preenchendo login do OGEA...")
    campo_login.fill(OGEA_USUARIO)
    campo_senha.fill(OGEA_SENHA)

    botao_entrar = page.get_by_role("button", name="Entrar")
    esperar_elemento_ogea(
        botao_entrar,
        "botão Entrar do OGEA",
    )
    botao_entrar.click()

    print("Login enviado. Aguardando autenticação...")

    try:
        page.wait_for_url(
            re.compile(r"(?!.*login/auth).*"),
            timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000,
        )
    except Exception:
        # Alguns ambientes mantêm a mesma URL durante o redirecionamento.
        page.wait_for_timeout(3000)


def abrir_ordem_servico_ogea(page):
    print("Abrindo Ordem de Serviço diretamente pelo link...")

    page.goto(
        OGEA_ORDEM_SERVICO_URL,
        wait_until="domcontentloaded",
        timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000,
    )

    filtro_situacao = page.locator("#situationServiceOrder")
    esperar_elemento_ogea(
        filtro_situacao,
        "filtro Situação da Ordem de Serviço",
    )

    print("Tela Ordem de Serviço carregada.")


def selecionar_prestador_ogea(page, prestador=None):
    prestador = prestador or OGEA_PRESTADOR
    print(f"Selecionando prestador {prestador}...")

    # Seletores gravados diretamente pelo Playwright Codegen.
    lista_prestador = page.get_by_role("list").nth(4)
    esperar_elemento_ogea(
        lista_prestador,
        "lista de Prestador de Serviço",
    )
    lista_prestador.click()

    campo_busca = page.get_by_role("textbox").nth(3)
    esperar_elemento_ogea(
        campo_busca,
        "campo de busca do Prestador",
    )
    campo_busca.fill(prestador)

    opcao_prestador = page.get_by_role(
        "treeitem",
        name=prestador,
    )
    esperar_elemento_ogea(
        opcao_prestador,
        f"opção {prestador}",
    )
    opcao_prestador.click()

    print(f"Prestador selecionado: {prestador}")


def preencher_periodo_ogea(page):
    agora = datetime.now()
    inicio = (
        agora - timedelta(days=OGEA_DIAS_ABERTURA)
    ).replace(
        hour=0,
        minute=0,
        second=0,
        microsecond=0,
    )

    inicio_texto = inicio.strftime("%d/%m/%Y %H:%M")
    fim_texto = agora.strftime("%d/%m/%Y %H:%M")

    print(
        "Data de Abertura OGEA: "
        f"{inicio_texto} até {fim_texto}"
    )

    campo_inicio = page.locator("#openingDateFrom_value")
    campo_fim = page.locator("#openingDateTo_value")

    campo_inicio.wait_for(
        state="attached",
        timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000,
    )
    campo_fim.wait_for(
        state="attached",
        timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000,
    )

    # Os campos usam componente de data legado. Preenche o valor e dispara
    # os mesmos eventos esperados pelo formulário.
    for campo, valor in [
        (campo_inicio, inicio_texto),
        (campo_fim, fim_texto),
    ]:
        campo.evaluate(
            """(el, valor) => {
                el.removeAttribute('readonly');
                el.value = valor;
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
                el.dispatchEvent(new Event('blur', {bubbles: true}));
            }""",
            valor,
        )


def selecionar_situacao_abertas_ogea(page):
    print("Selecionando Situação = Abertas...")

    filtro = page.locator("#situationServiceOrder")
    esperar_elemento_ogea(
        filtro,
        "filtro Situação",
    )
    filtro.select_option("OPEN")


def pesquisar_ogea(page):
    print("Pesquisando OSs no OGEA...")

    botao_pesquisar = page.get_by_role(
        "button",
        name=" Pesquisar",
    )
    esperar_elemento_ogea(
        botao_pesquisar,
        "botão Pesquisar",
    )
    botao_pesquisar.click()

    # Não depende mais do texto Total. A próxima etapa só começa quando o
    # botão Exportar gravado pelo Codegen estiver visível e habilitado.
    botao_exportar = page.get_by_role(
        "button",
        name="Exportar ",
    )

    esperar_elemento_ogea(
        botao_exportar,
        "botão Exportar após a pesquisa",
    )

    limite = time.time() + OGEA_TIMEOUT_CARREGAMENTO

    while time.time() < limite:
        try:
            if botao_exportar.is_visible() and botao_exportar.is_enabled():
                print("Resultado da pesquisa carregado.")
                return
        except Exception:
            pass

        page.wait_for_timeout(1000)

    raise TimeoutError(
        "O botão Exportar apareceu, mas não ficou habilitado dentro de "
        f"{OGEA_TIMEOUT_CARREGAMENTO} segundos."
    )


def solicitar_relatorio_completo_ogea(page):
    print("Solicitando relatório completo pelo Exporter...")

    botao_exportar = page.get_by_role(
        "button",
        name="Exportar ",
    )
    esperar_elemento_ogea(
        botao_exportar,
        "botão Exportar",
    )

    opcao_completo = page.get_by_role(
        "link",
        name="Completo (via Exporter)",
    )

    # O clique no botão Exportar às vezes não abre o menu suspenso na primeira
    # tentativa (o robô segue em frente achando que abriu e trava esperando
    # uma opção que nunca vai aparecer) — clicar de novo até o menu realmente
    # abrir resolve, em vez de confiar cegamente no primeiro clique.
    tentativas = 5
    for tentativa in range(1, tentativas + 1):
        botao_exportar.scroll_into_view_if_needed()
        botao_exportar.click()

        try:
            opcao_completo.wait_for(state="visible", timeout=3000)
            break
        except Exception:
            print(
                f"Menu Exportar não abriu na tentativa {tentativa}/{tentativas}, "
                "tentando de novo..."
            )
    else:
        raise TimeoutError(
            "O menu do botão Exportar não abriu depois de várias tentativas."
        )

    opcao_completo.click()

    print("Solicitação enviada ao Exporter.")

    page.goto(
        "https://tefti.workfinity.com.br/exporterCsvExportRequest/list",
        wait_until="domcontentloaded",
        timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000,
    )

    esperar_elemento_ogea(
        page.locator("table").first,
        "lista Meus Relatórios",
    )


def localizar_linha_relatorio(page, codigo_alvo=None):
    linhas = page.locator("table tbody tr")

    for indice in range(linhas.count()):
        linha = linhas.nth(indice)

        try:
            texto = linha.inner_text()
        except Exception:
            continue

        if "reportRequest-ServiceOrderTransformScript" not in texto:
            continue

        match = re.search(
            r"reportRequest-ServiceOrderTransformScript-[A-Za-z0-9_-]+",
            texto,
        )
        codigo = match.group(0) if match else ""

        if codigo_alvo and codigo != codigo_alvo:
            continue

        return linha, codigo, texto

    return None, "", ""


def aguardar_e_baixar_relatorio_ogea(page):
    print("Aguardando o Exporter concluir o relatório...")

    linha, codigo, texto = localizar_linha_relatorio(page)

    if linha is None or not codigo:
        raise RuntimeError(
            "Não encontrei o relatório recém-criado em Meus Relatórios."
        )

    print(f"Relatório acompanhado: {codigo}")

    limite = time.time() + OGEA_TIMEOUT_CARREGAMENTO
    tentativa = 0

    while time.time() < limite:
        tentativa += 1

        linha, _, texto = localizar_linha_relatorio(
            page,
            codigo,
        )

        texto_norm = normalizar_texto(texto)

        if linha is not None and "CONCLUIDO" in texto_norm:
            print("Relatório OGEA concluído.")

            download_link = linha.get_by_role(
                "link",
                name="Download",
            )

            esperar_elemento_ogea(
                download_link,
                "link Download do relatório concluído",
            )

            with page.expect_download(
                timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000
            ) as info:
                download_link.click()

            download = info.value
            sugerido = download.suggested_filename or ""

            caminho_csv_atual = (
                PASTA_DOWNLOADS_OGEA
                / "relatorio_ogea_atual.csv"
            )

            if sugerido.lower().endswith(".csv"):
                download.save_as(caminho_csv_atual)
                print(f"CSV do OGEA salvo em: {caminho_csv_atual}")
                return caminho_csv_atual

            caminho_zip = (
                PASTA_DOWNLOADS_OGEA
                / "relatorio_ogea_atual.zip"
            )
            download.save_as(caminho_zip)
            print(f"ZIP do OGEA salvo em: {caminho_zip}")

            pasta_extracao = (
                PASTA_DOWNLOADS_OGEA
                / "extraido_atual"
            )
            pasta_extracao.mkdir(parents=True, exist_ok=True)

            with zipfile.ZipFile(caminho_zip, "r") as arquivo_zip:
                arquivo_zip.extractall(pasta_extracao)

            csvs = list(pasta_extracao.rglob("*.csv"))

            if not csvs:
                raise RuntimeError(
                    "O ZIP baixado do OGEA não contém nenhum CSV."
                )

            csv_extraido = max(
                csvs,
                key=lambda arquivo: arquivo.stat().st_size,
            )

            shutil.copy2(csv_extraido, caminho_csv_atual)

            # Mantém somente o CSV atual; remove ZIP e pasta temporária.
            shutil.rmtree(pasta_extracao, ignore_errors=True)
            try:
                caminho_zip.unlink()
            except Exception:
                pass

            print(f"CSV do OGEA extraído em: {caminho_csv_atual}")
            return caminho_csv_atual

        status = texto.strip().replace("\n", " | ")

        print(
            f"Tentativa {tentativa}: relatório ainda não concluído. "
            f"{status}"
        )

        # Conforme o comportamento real do OGEA: espera 20 segundos e
        # atualiza a tela Meus Relatórios.
        time.sleep(20)

        page.reload(
            wait_until="domcontentloaded",
            timeout=OGEA_TIMEOUT_CARREGAMENTO * 1000,
        )

        esperar_elemento_ogea(
            page.locator("table").first,
            "lista Meus Relatórios após atualizar",
        )

    salvar_screenshot_erro(page, "timeout_exporter_ogea")
    raise TimeoutError(
        "O relatório do OGEA não ficou pronto dentro de "
        f"{OGEA_TIMEOUT_CARREGAMENTO} segundos."
    )


def baixar_relatorio_ogea():
    preparar_pastas()
    validar_ogea()

    print("")
    print("=" * 70)
    print("BAIXANDO RELATÓRIO DO OGEA")
    print("=" * 70)

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=False)
        contexto = navegador.new_context(
            accept_downloads=True,
            viewport={"width": 1600, "height": 900},
        )
        page = contexto.new_page()

        page.set_default_timeout(
            OGEA_TIMEOUT_CARREGAMENTO * 1000
        )
        page.set_default_navigation_timeout(
            OGEA_TIMEOUT_CARREGAMENTO * 1000
        )

        try:
            fazer_login_ogea(page)
            abrir_ordem_servico_ogea(page)
            selecionar_prestador_ogea(page)
            preencher_periodo_ogea(page)
            selecionar_situacao_abertas_ogea(page)
            pesquisar_ogea(page)
            solicitar_relatorio_completo_ogea(page)
            return aguardar_e_baixar_relatorio_ogea(page)

        except Exception:
            salvar_screenshot_erro(page, "erro_download_ogea")
            raise
        finally:
            contexto.close()
            navegador.close()

def baixar_relatorios_automaticamente():
    limpar_downloads_roteirizacao()

    caminho_mobyan = baixar_relatorio_mobyan()
    caminho_ogea = baixar_relatorio_ogea()

    print("")
    print("Downloads automáticos concluídos.")
    print(f"Mobyan: {caminho_mobyan}")
    print(f"OGEA: {caminho_ogea}")

    return caminho_mobyan, caminho_ogea
