import argparse
import base64
import os
import re
import shutil
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlsplit

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = None
    PdfWriter = None

from exportador_mobyan import (
    MOBYAN_URL,
    fazer_login,
    finalizar_sessao_anterior_se_aparecer,
    validar_configuracoes,
)
from baixar_relatorios_roteirizacao import (
    OGEA_TIMEOUT_CARREGAMENTO,
    OGEA_URL,
    abrir_ordem_servico_ogea,
    esperar_elemento_ogea,
    fazer_login_ogea,
    validar_ogea,
)


BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

ARQUIVO_ROTEIRIZACAO = (
    BASE_DIR
    / "outputs"
    / "roteirizacao"
    / "roteirizacao_atual.xlsx"
)
PASTA_PDFS = BASE_DIR / "outputs" / "pdfs"
PASTA_PDFS_TEMP = PASTA_PDFS / "_temporarios"
PASTA_LOGS = BASE_DIR / "logs" / "pdfs"

try:
    PDF_TIMEOUT_SEGUNDOS = int(
        os.getenv("PDF_TIMEOUT_SEGUNDOS", "180")
    )
except ValueError:
    PDF_TIMEOUT_SEGUNDOS = 180

try:
    MOBYAN_TIMEOUT_POR_OS_SEGUNDOS = int(
        os.getenv("MOBYAN_TIMEOUT_POR_OS_SEGUNDOS", "60")
    )
except ValueError:
    MOBYAN_TIMEOUT_POR_OS_SEGUNDOS = 60

MOBYAN_TIMEOUT_POR_OS_SEGUNDOS = max(15, MOBYAN_TIMEOUT_POR_OS_SEGUNDOS)

# O valor normalmente é detectado no HTML/URL da sessão. Este fallback é
# utilizado apenas quando o sistema não expõe o parâmetro durante o login.
MOBYAN_RELATORIO_USER = os.getenv(
    "MOBYAN_RELATORIO_USER",
    "72107",
).strip()


class PdfNaoEncontrado(RuntimeError):
    pass


def normalizar_texto(valor):
    texto = "" if valor is None else str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )
    return texto


def limpar_nome_arquivo(texto):
    texto = normalizar_texto(texto)
    texto = re.sub(r'[<>:"/\\|?*]+', " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip(" .")
    return texto or "Sem nome"


def extrair_oss(valor):
    numeros = re.findall(r"\d+", str(valor or ""))
    resultado = []
    vistos = set()

    for numero in numeros:
        if numero not in vistos:
            resultado.append(numero)
            vistos.add(numero)

    return resultado


def preparar_pastas():
    PASTA_PDFS.mkdir(parents=True, exist_ok=True)
    PASTA_PDFS_TEMP.mkdir(parents=True, exist_ok=True)
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)


def limpar_pdfs_antigos(origem=None):
    preparar_pastas()

    for item in PASTA_PDFS.iterdir():
        if origem and not item.name.lower().startswith(
            f"{origem.lower()} - "
        ):
            continue

        try:
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()
        except Exception as erro:
            raise RuntimeError(
                f"Não consegui apagar o PDF antigo: {item}. "
                "Feche o arquivo caso esteja aberto e tente novamente. "
                f"Erro: {erro}"
            ) from erro

    if not origem:
        shutil.rmtree(PASTA_PDFS_TEMP, ignore_errors=True)
        PASTA_PDFS_TEMP.mkdir(parents=True, exist_ok=True)

    if origem:
        print(f"PDFs antigos da origem {origem} removidos.")
    else:
        print("PDFs e roteiros antigos removidos.")


def salvar_screenshot(page, nome):
    try:
        preparar_pastas()
        caminho = PASTA_LOGS / (
            f"{nome}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        )
        page.screenshot(path=str(caminho), full_page=True)
        print(f"Screenshot salvo em: {caminho}")
    except Exception:
        pass


def ler_listas_pdf():
    if not ARQUIVO_ROTEIRIZACAO.exists():
        raise FileNotFoundError(
            "A roteirização ainda não existe. Gere a roteirização antes "
            "de criar os PDFs."
        )

    try:
        df = pd.read_excel(
            ARQUIVO_ROTEIRIZACAO,
            sheet_name="Listas PDF",
            dtype=str,
            keep_default_na=False,
            na_filter=False,
        )
    except Exception as erro:
        raise RuntimeError(
            "Não consegui ler a aba 'Listas PDF' da roteirização."
        ) from erro

    colunas_necessarias = {
        "Técnico",
        "OSs Mobyan",
        "OSs OGEA",
    }
    faltantes = colunas_necessarias - set(df.columns)

    if faltantes:
        raise RuntimeError(
            "A aba 'Listas PDF' não possui as colunas esperadas: "
            + ", ".join(sorted(faltantes))
        )

    listas = []

    for _, linha in df.iterrows():
        tecnico = str(linha.get("Técnico", "")).strip()

        if not tecnico:
            continue

        oss_mobyan = extrair_oss(linha.get("OSs Mobyan", ""))
        oss_ogea = extrair_oss(linha.get("OSs OGEA", ""))

        if not oss_mobyan and not oss_ogea:
            continue

        listas.append({
            "tecnico": tecnico,
            "mobyan": oss_mobyan,
            "ogea": oss_ogea,
        })

    if not listas:
        raise RuntimeError(
            "A aba 'Listas PDF' não contém nenhuma OS para gerar."
        )

    return listas


def resposta_e_pdf(resposta):
    try:
        conteudo = resposta.body()
    except Exception:
        return False, b"", ""

    content_type = str(
        resposta.headers.get("content-type", "")
    ).lower()

    eh_pdf = (
        conteudo.startswith(b"%PDF")
        or "application/pdf" in content_type
    )

    return eh_pdf, conteudo, content_type


def baixar_pdf_por_url(
    contexto,
    url,
    caminho_saida,
    descricao,
    referer=None,
    timeout_segundos=None,
):
    if timeout_segundos is None:
        timeout_segundos = PDF_TIMEOUT_SEGUNDOS

    limite = time.time() + timeout_segundos
    tentativa = 0
    ultimo_detalhe = ""

    while time.time() < limite:
        tentativa += 1
        restante = max(0, int(limite - time.time()))

        try:
            cabecalhos = {}
            if referer:
                cabecalhos["Referer"] = referer

            argumentos_request = {
                "timeout": min(
                    60000,
                    max(10000, restante * 1000),
                ),
                "fail_on_status_code": False,
            }
            if cabecalhos:
                argumentos_request["headers"] = cabecalhos

            resposta = contexto.request.get(
                url,
                **argumentos_request,
            )

            eh_pdf, conteudo, content_type = resposta_e_pdf(resposta)

            if eh_pdf:
                caminho_saida.write_bytes(conteudo)
                print(
                    f"{descricao}: PDF salvo em {caminho_saida}"
                )
                return caminho_saida

            detalhe = (
                f"status={resposta.status}, "
                f"content-type={content_type or 'não informado'}, "
                f"tamanho={len(conteudo)} bytes"
            )

            if detalhe != ultimo_detalhe:
                print(
                    f"{descricao}: relatório ainda não retornou PDF "
                    f"({detalhe}). Nova tentativa em 5s..."
                )
                ultimo_detalhe = detalhe

        except Exception as erro:
            print(
                f"{descricao}: tentativa {tentativa} ainda não concluiu "
                f"({erro}). Nova tentativa em 5s..."
            )

        time.sleep(5)

    raise PdfNaoEncontrado(
        f"{descricao} não ficou pronto dentro de "
        f"{timeout_segundos} segundos."
    )


def aguardar_url_popup(popup):
    limite = time.time() + PDF_TIMEOUT_SEGUNDOS

    while time.time() < limite:
        url = str(popup.url or "").strip()

        if url and url != "about:blank":
            return url

        popup.wait_for_timeout(500)

    raise TimeoutError(
        "A janela do PDF abriu, mas não recebeu uma URL dentro do prazo."
    )


def candidatos_pdf_popup(popup):
    candidatos = []

    url_principal = str(popup.url or "").strip()
    if url_principal.startswith(("http://", "https://")):
        candidatos.append(url_principal)

    for seletor, atributo in [
        ("embed", "src"),
        ("iframe", "src"),
        ("object", "data"),
    ]:
        try:
            elementos = popup.locator(seletor)

            for indice in range(elementos.count()):
                valor = elementos.nth(indice).get_attribute(atributo)
                valor = str(valor or "").strip()

                if valor.startswith(("http://", "https://")):
                    candidatos.append(valor)
        except Exception:
            continue

    resultado = []
    vistos = set()

    for candidato in candidatos:
        if candidato not in vistos:
            resultado.append(candidato)
            vistos.add(candidato)

    return resultado


class CapturadorPdfRede:
    """
    Observa as respostas de rede abertas durante o clique em Imprimir.

    O visualizador interno do Chromium exibe o PDF, mas a página visível é
    apenas a interface com miniaturas e barra de ferramentas. O arquivo real
    precisa ser capturado da resposta de rede que alimenta esse visualizador.
    """

    def __init__(self):
        self.respostas = []
        self.diagnosticos = []
        self.processadas = set()

    def registrar(self, resposta):
        try:
            url = str(resposta.url or "")
            headers = resposta.headers
            content_type = str(
                headers.get("content-type", "")
            ).lower()
            disposicao = str(
                headers.get("content-disposition", "")
            ).lower()
            status = resposta.status
        except Exception:
            return

        url_normalizada = url.lower()

        relevante = (
            "application/pdf" in content_type
            or "application/octet-stream" in content_type
            or ".pdf" in url_normalizada
            or "print" in url_normalizada
            or "report" in url_normalizada
        )

        if not relevante:
            return

        self.respostas.append(resposta)
        self.diagnosticos.append(
            f"status={status}, content-type={content_type or 'n/a'}, "
            f"disposition={disposicao or 'n/a'}, url={url}"
        )

        # Evita crescimento excessivo em páginas com muitos recursos.
        self.respostas = self.respostas[-80:]
        self.diagnosticos = self.diagnosticos[-80:]

    def urls_candidatas(self):
        resultado = []
        vistos = set()

        for resposta in self.respostas:
            try:
                url = str(resposta.url or "").strip()
                headers = resposta.headers
                content_type = str(
                    headers.get("content-type", "")
                ).lower()
                disposicao = str(
                    headers.get("content-disposition", "")
                ).lower()
            except Exception:
                continue

            parece_pdf = (
                "application/pdf" in content_type
                or "application/octet-stream" in content_type
                or ".pdf" in url.lower()
                or ".pdf" in disposicao
            )

            if (
                parece_pdf
                and url.startswith(("http://", "https://"))
                and url not in vistos
            ):
                resultado.append(url)
                vistos.add(url)

        return resultado

    def tentar_salvar_corpo_completo(self, caminho_saida):
        """
        Salva apenas respostas completas. Respostas 206 podem ser fragmentos
        usados pelo visualizador e não devem ser gravadas como PDF final.
        """
        for resposta in list(self.respostas):
            identificador = id(resposta)

            if identificador in self.processadas:
                continue

            self.processadas.add(identificador)

            try:
                status = resposta.status
                headers = resposta.headers
                content_type = str(
                    headers.get("content-type", "")
                ).lower()
                conteudo = resposta.body()
            except Exception:
                continue

            if not conteudo.startswith(b"%PDF"):
                continue

            if status == 200:
                caminho_saida.write_bytes(conteudo)
                return True

            if status == 206:
                content_range = str(
                    headers.get("content-range", "")
                )

                match = re.search(r"/(\d+)$", content_range)

                if match and len(conteudo) == int(match.group(1)):
                    caminho_saida.write_bytes(conteudo)
                    return True

        return False


def localizar_visualizador_pdf_chromium(
    popup,
    descricao,
    timeout_segundos=None,
):
    """
    Localiza o visualizador PDF interno do Chromium e o botão Download.
    """
    if timeout_segundos is None:
        timeout_segundos = PDF_TIMEOUT_SEGUNDOS

    limite = time.time() + timeout_segundos
    proximo_aviso = time.time()

    while time.time() < limite:
        for frame in popup.frames:
            try:
                resultado = frame.evaluate(
                    """() => {
                        const viewer =
                            document.getElementById('viewer') ||
                            document.querySelector('pdf-viewer');

                        if (!viewer || !viewer.shadowRoot) {
                            return null;
                        }

                        const viewerRoot = viewer.shadowRoot;
                        const toolbar =
                            viewerRoot.getElementById('toolbar') ||
                            viewerRoot.querySelector('viewer-toolbar') ||
                            viewerRoot.querySelector('viewer-pdf-toolbar');

                        let downloads = null;
                        let download = null;

                        if (toolbar && toolbar.shadowRoot) {
                            downloads =
                                toolbar.shadowRoot.getElementById('downloads') ||
                                toolbar.shadowRoot.querySelector(
                                    'viewer-download-controls'
                                );

                            download =
                                toolbar.shadowRoot.getElementById('download') ||
                                toolbar.shadowRoot.querySelector(
                                    'cr-icon-button#download'
                                );
                        }

                        if (
                            !download &&
                            downloads &&
                            downloads.shadowRoot
                        ) {
                            download =
                                downloads.shadowRoot.getElementById('download') ||
                                downloads.shadowRoot.querySelector(
                                    'cr-icon-button#download'
                                );
                        }

                        if (!download) {
                            download =
                                viewerRoot.getElementById('download') ||
                                viewerRoot.querySelector(
                                    'cr-icon-button#download'
                                );
                        }

                        const progresso =
                            viewer.loadProgress_ ??
                            viewer.loadProgress ??
                            null;

                        return {
                            encontrou: Boolean(download),
                            desabilitado: Boolean(
                                download && download.disabled
                            ),
                            progresso,
                        };
                    }"""
                )
            except Exception:
                continue

            if not resultado or not resultado.get("encontrou"):
                continue

            progresso = resultado.get("progresso")
            desabilitado = resultado.get("desabilitado", False)

            carregou = (
                progresso is None
                or progresso == 100
                or progresso == 1
            )

            if carregou and not desabilitado:
                print(
                    f"{descricao}: visualizador PDF carregado; "
                    "botão Download localizado."
                )
                return frame

        agora = time.time()

        if agora >= proximo_aviso:
            restante = max(0, int(limite - agora))
            print(
                f"{descricao}: aguardando o visualizador PDF ficar "
                f"pronto... tempo restante={restante}s"
            )
            proximo_aviso = agora + 15

        popup.wait_for_timeout(750)

    raise TimeoutError(
        f"{descricao}: o PDF apareceu na tela, mas o botão Download "
        f"não ficou disponível em {timeout_segundos}s."
    )


def clicar_download_visualizador(frame):
    """
    Aciona exatamente o botão Download do visualizador interno do Chromium.
    """
    return frame.evaluate(
        """() => {
            const viewer =
                document.getElementById('viewer') ||
                document.querySelector('pdf-viewer');

            if (!viewer || !viewer.shadowRoot) {
                return false;
            }

            const viewerRoot = viewer.shadowRoot;
            const toolbar =
                viewerRoot.getElementById('toolbar') ||
                viewerRoot.querySelector('viewer-toolbar') ||
                viewerRoot.querySelector('viewer-pdf-toolbar');

            let downloads = null;
            let download = null;

            if (toolbar && toolbar.shadowRoot) {
                downloads =
                    toolbar.shadowRoot.getElementById('downloads') ||
                    toolbar.shadowRoot.querySelector(
                        'viewer-download-controls'
                    );

                download =
                    toolbar.shadowRoot.getElementById('download') ||
                    toolbar.shadowRoot.querySelector(
                        'cr-icon-button#download'
                    );
            }

            if (
                !download &&
                downloads &&
                downloads.shadowRoot
            ) {
                download =
                    downloads.shadowRoot.getElementById('download') ||
                    downloads.shadowRoot.querySelector(
                        'cr-icon-button#download'
                    );
            }

            if (!download) {
                download =
                    viewerRoot.getElementById('download') ||
                    viewerRoot.querySelector(
                        'cr-icon-button#download'
                    );
            }

            if (!download || download.disabled) {
                return false;
            }

            download.click();
            return true;
        }"""
    )


def salvar_pdf_popup(
    contexto,
    popup,
    caminho_saida,
    descricao,
    referer=None,
    quantidade_esperada=None,
    capturador=None,
):
    """
    Salva o PDF OGEA usando o botão Download do visualizador do Chromium.
    """
    aguardar_url_popup(popup)

    try:
        popup.wait_for_load_state(
            "domcontentloaded",
            timeout=min(PDF_TIMEOUT_SEGUNDOS, 60) * 1000,
        )
    except Exception:
        pass

    frame_visualizador = localizar_visualizador_pdf_chromium(
        popup,
        descricao,
        timeout_segundos=PDF_TIMEOUT_SEGUNDOS,
    )

    popup.wait_for_timeout(1500)

    print(
        f"{descricao}: clicando no botão Download do visualizador..."
    )

    with popup.expect_download(
        timeout=PDF_TIMEOUT_SEGUNDOS * 1000
    ) as download_info:
        clicou = clicar_download_visualizador(
            frame_visualizador
        )

        if not clicou:
            raise RuntimeError(
                f"{descricao}: encontrei o visualizador, mas não consegui "
                "acionar o botão Download."
            )

    download = download_info.value
    download.save_as(caminho_saida)

    if not caminho_saida.exists():
        raise RuntimeError(
            f"{descricao}: o download foi acionado, mas o arquivo não "
            "foi salvo."
        )

    conteudo_inicial = caminho_saida.read_bytes()[:4]

    if conteudo_inicial != b"%PDF":
        try:
            caminho_saida.unlink()
        except Exception:
            pass

        raise RuntimeError(
            f"{descricao}: o arquivo baixado não é um PDF válido."
        )

    tamanho = caminho_saida.stat().st_size

    print(
        f"{descricao}: PDF original baixado pelo visualizador "
        f"({tamanho:,} bytes)."
    )

    return caminho_saida


def capturar_sessao_mobyan(page, contexto):
    padrao_sessao = re.compile(r"(\(S\([^)]+\)\))", re.I)
    padrao_usuario_relatorio = re.compile(
        r"ReportRDLC\.aspx[^\"'<>\s]*[?&]user=(\d+)",
        re.I,
    )

    textos = []

    for pagina in contexto.pages:
        textos.append(str(pagina.url or ""))

        try:
            textos.append(pagina.content())
        except Exception:
            pass

        for frame in pagina.frames:
            textos.append(str(frame.url or ""))

    sessao = ""
    usuario_relatorio = ""

    for texto in textos:
        if not sessao:
            match = padrao_sessao.search(texto)
            if match:
                sessao = match.group(1)

        if not usuario_relatorio:
            match = padrao_usuario_relatorio.search(texto)
            if match:
                usuario_relatorio = match.group(1)

    if not sessao:
        # Em alguns carregamentos a sessão aparece somente depois que a tela
        # inicial termina de montar seus frames.
        limite = time.time() + 30

        while time.time() < limite and not sessao:
            for pagina in contexto.pages:
                for candidato in [pagina.url] + [
                    frame.url for frame in pagina.frames
                ]:
                    match = padrao_sessao.search(str(candidato or ""))
                    if match:
                        sessao = match.group(1)
                        break

                if sessao:
                    break

            if not sessao:
                page.wait_for_timeout(1000)

    if not usuario_relatorio:
        usuario_relatorio = MOBYAN_RELATORIO_USER

    if not sessao:
        raise RuntimeError(
            "Não consegui identificar a sessão atual da Mobyan após o login."
        )

    if not usuario_relatorio:
        raise RuntimeError(
            "Não consegui identificar o usuário usado no relatório da Mobyan."
        )

    return sessao, usuario_relatorio


def montar_url_pdf_mobyan(sessao, usuario_relatorio, oss):
    partes = urlsplit(MOBYAN_URL)
    ids = quote(",".join(oss), safe=",")

    return (
        f"{partes.scheme}://{partes.netloc}/{sessao}/"
        "appdata/ReportRDLC.aspx"
        f"?reportName=OSGetnet&id={ids}"
        f"&param=0&user={quote(usuario_relatorio)}&lang=22&comp=3"
    )


def validar_pdf_baixado(caminho):
    """Confirma que o arquivo baixado é um PDF legível."""
    if not caminho.exists() or caminho.stat().st_size <= 0:
        raise PdfNaoEncontrado(
            "O arquivo não foi criado ou ficou vazio."
        )

    if caminho.read_bytes()[:4] != b"%PDF":
        raise PdfNaoEncontrado(
            "O arquivo retornado não possui assinatura de PDF."
        )

    validar_pypdf()
    leitor = PdfReader(str(caminho))

    if len(leitor.pages) < 1:
        raise PdfNaoEncontrado(
            "O PDF foi criado, mas não possui páginas."
        )

    return len(leitor.pages)


def unir_pdfs_mobyan_do_tecnico(
    tecnico,
    partes,
    caminho_saida,
):
    """Une os PDFs individuais da Mobyan na ordem original das OSs."""
    validar_pypdf()
    escritor = PdfWriter()
    total_paginas = 0

    for parte in partes:
        leitor = PdfReader(str(parte))

        for pagina in leitor.pages:
            escritor.add_page(pagina)
            total_paginas += 1

    escritor.add_metadata({
        "/Title": f"Mobyan - {tecnico}",
        "/Subject": "Ordens de Serviço Mobyan",
        "/Creator": "Central Mobyan",
    })

    with caminho_saida.open("wb") as arquivo_saida:
        escritor.write(arquivo_saida)

    paginas_finais = validar_pdf_baixado(caminho_saida)

    if paginas_finais != total_paginas:
        raise RuntimeError(
            f"Eram esperadas {total_paginas} páginas, "
            f"mas o PDF unido possui {paginas_finais}."
        )

    return total_paginas


def salvar_log_falhas_mobyan(falhas, iniciado_em):
    """Grava um relatório persistente das OSs ignoradas."""
    if not falhas:
        return None

    preparar_pastas()
    timestamp = iniciado_em.strftime("%Y-%m-%d_%H-%M-%S")
    caminho = PASTA_LOGS / f"falhas_mobyan_{timestamp}.txt"

    linhas = [
        "RELATÓRIO DE FALHAS - PDFs MOBYAN",
        "=" * 70,
        f"Execução iniciada em: {iniciado_em.strftime('%d/%m/%Y %H:%M:%S')}",
        f"Total de falhas: {len(falhas)}",
        "",
    ]

    for indice, falha in enumerate(falhas, start=1):
        linhas.extend([
            f"{indice}. Técnico: {falha['tecnico']}",
            f"   OS: {falha['os']}",
            f"   Etapa: {falha['etapa']}",
            f"   Erro: {falha['erro']}",
            "",
        ])

    caminho.write_text("\n".join(linhas), encoding="utf-8")
    return caminho

def gerar_pdfs_mobyan(listas, pasta_destino=PASTA_PDFS):
    tarefas = [
        item
        for item in listas
        if item["mobyan"]
    ]

    if not tarefas:
        print("Nenhum PDF da Mobyan necessário.")
        return []

    validar_configuracoes()
    validar_pypdf()
    preparar_pastas()
    pasta_destino.mkdir(parents=True, exist_ok=True)

    iniciado_em = datetime.now()
    falhas = []
    gerados = []
    total_oss = sum(len(item["mobyan"]) for item in tarefas)
    os_processada = 0

    pasta_individuais = (
        PASTA_PDFS_TEMP
        / "_mobyan_individuais"
        / iniciado_em.strftime("%Y-%m-%d_%H-%M-%S")
    )
    shutil.rmtree(pasta_individuais, ignore_errors=True)
    pasta_individuais.mkdir(parents=True, exist_ok=True)

    print("")
    print("=" * 70)
    print("GERANDO PDFs DA MOBYAN - UMA OS POR VEZ")
    print("=" * 70)
    print(
        f"Total de OSs Mobyan: {total_oss}. "
        "Uma falha individual será ignorada e registrada no log."
    )

    with sync_playwright() as playwright:
        navegador = playwright.chromium.launch(headless=False)
        contexto = navegador.new_context(
            accept_downloads=True,
            viewport={"width": 1600, "height": 900},
        )
        page = contexto.new_page()
        page.set_default_timeout(PDF_TIMEOUT_SEGUNDOS * 1000)
        page.set_default_navigation_timeout(
            PDF_TIMEOUT_SEGUNDOS * 1000
        )

        try:
            print("Abrindo sistema da Mobyan...")
            page.goto(
                MOBYAN_URL,
                wait_until="domcontentloaded",
                timeout=PDF_TIMEOUT_SEGUNDOS * 1000,
            )
            fazer_login(page)
            finalizar_sessao_anterior_se_aparecer(page)
            page.wait_for_timeout(5000)

            sessao, usuario_relatorio = capturar_sessao_mobyan(
                page,
                contexto,
            )
            print("Sessão atual da Mobyan identificada.")

            for indice_tecnico, item in enumerate(tarefas, start=1):
                tecnico = item["tecnico"]
                oss = item["mobyan"]
                nome = limpar_nome_arquivo(tecnico)
                pasta_tecnico = pasta_individuais / nome
                pasta_tecnico.mkdir(parents=True, exist_ok=True)
                partes_validas = []

                print("")
                print(
                    f"[{indice_tecnico}/{len(tarefas)}] "
                    f"Mobyan - {tecnico}: {len(oss)} OS(s)."
                )

                for indice_os, numero_os in enumerate(oss, start=1):
                    os_processada += 1
                    caminho_individual = (
                        pasta_tecnico
                        / f"{indice_os:03d} - OS {numero_os}.pdf"
                    )
                    url = montar_url_pdf_mobyan(
                        sessao,
                        usuario_relatorio,
                        [numero_os],
                    )
                    descricao = (
                        f"Mobyan - {tecnico} - OS {numero_os}"
                    )

                    print(
                        f"  [{os_processada}/{total_oss}] "
                        f"OS {numero_os} ({indice_os}/{len(oss)})..."
                    )

                    try:
                        baixar_pdf_por_url(
                            contexto,
                            url,
                            caminho_individual,
                            descricao,
                            referer=page.url,
                            timeout_segundos=(
                                MOBYAN_TIMEOUT_POR_OS_SEGUNDOS
                            ),
                        )
                        paginas = validar_pdf_baixado(
                            caminho_individual
                        )
                        partes_validas.append(caminho_individual)
                        print(
                            f"  OK | Técnico: {tecnico} | "
                            f"OS: {numero_os} | "
                            f"{paginas} página(s)."
                        )

                    except Exception as erro:
                        try:
                            caminho_individual.unlink(missing_ok=True)
                        except Exception:
                            pass

                        mensagem = str(erro).strip() or type(erro).__name__
                        falhas.append({
                            "tecnico": tecnico,
                            "os": numero_os,
                            "etapa": "geração individual",
                            "erro": mensagem,
                        })
                        print(
                            "  ERRO IGNORADO | "
                            f"Técnico: {tecnico} | OS: {numero_os} | "
                            f"{mensagem}"
                        )
                        print(
                            "  A automação continuará para a próxima OS."
                        )

                caminho_final = (
                    pasta_destino / f"Mobyan - {nome}.pdf"
                )

                if not partes_validas:
                    print(
                        f"Mobyan - {tecnico}: nenhuma OS válida foi "
                        "gerada. O técnico continuará com o PDF do OGEA, "
                        "caso exista."
                    )
                    continue

                try:
                    total_paginas = unir_pdfs_mobyan_do_tecnico(
                        tecnico,
                        partes_validas,
                        caminho_final,
                    )
                    gerados.append(caminho_final)
                    print(
                        f"Mobyan - {tecnico}: PDF unido com "
                        f"{len(partes_validas)}/{len(oss)} OS(s) e "
                        f"{total_paginas} página(s)."
                    )

                except Exception as erro:
                    try:
                        caminho_final.unlink(missing_ok=True)
                    except Exception:
                        pass

                    mensagem = str(erro).strip() or type(erro).__name__
                    falhas.append({
                        "tecnico": tecnico,
                        "os": "PDF do técnico",
                        "etapa": "união das OSs individuais",
                        "erro": mensagem,
                    })
                    print(
                        "ERRO IGNORADO AO UNIR | "
                        f"Técnico: {tecnico} | {mensagem}"
                    )
                    print(
                        "A automação continuará para o próximo técnico."
                    )

        except Exception:
            # Falhas estruturais, como login ou sessão inválida, ainda são
            # fatais porque impedem todas as OSs de serem processadas.
            salvar_screenshot(page, "erro_pdfs_mobyan")
            raise
        finally:
            contexto.close()
            navegador.close()

    caminho_log = salvar_log_falhas_mobyan(
        falhas,
        iniciado_em,
    )
    falhas_individuais = sum(
        1
        for falha in falhas
        if falha["etapa"] == "geração individual"
    )

    print("")
    print("-" * 70)
    print("RESUMO DOS PDFs MOBYAN")
    print(f"OSs solicitadas: {total_oss}")
    print(f"OSs geradas individualmente: {total_oss - falhas_individuais}")
    print(f"OSs com erro ignorado: {falhas_individuais}")

    if caminho_log:
        print(f"Log das falhas: {caminho_log}")
        print(
            "As OSs listadas nesse arquivo deverão ser feitas "
            "manualmente."
        )
    else:
        print("Nenhuma OS individual da Mobyan apresentou erro.")

    print("-" * 70)

    shutil.rmtree(pasta_individuais, ignore_errors=True)
    return gerados


def primeiro_elemento_visivel_ogea(
    page,
    candidatos,
    descricao,
    timeout_segundos=None,
):
    """
    Localiza um elemento usando somente seletores Playwright.
    Não usa coordenadas, JavaScript nem cliques no escuro.
    """
    if timeout_segundos is None:
        timeout_segundos = PDF_TIMEOUT_SEGUNDOS

    limite = time.time() + timeout_segundos

    while time.time() < limite:
        for candidato in candidatos:
            try:
                quantidade = candidato.count()
            except Exception:
                continue

            for indice in range(quantidade):
                item = candidato.nth(indice)

                try:
                    if item.is_visible():
                        return item
                except Exception:
                    continue

        page.wait_for_timeout(750)

    raise TimeoutError(
        f"{descricao} não apareceu dentro de "
        f"{timeout_segundos} segundos."
    )


def localizar_botao_pesquisar_ogea(page):
    return primeiro_elemento_visivel_ogea(
        page,
        [
            # Seletor exato gravado pelo Playwright Codegen.
            page.get_by_role(
                "button",
                name=" Pesquisar",
                exact=True,
            ),
            # Fallbacks ainda baseados no conteúdo real do botão.
            page.locator("button").filter(
                has_text=re.compile(r"Pesquisar", re.I)
            ),
            page.locator(
                "input[type='submit'][value*='Pesquisar']"
            ),
            page.get_by_text(
                "Pesquisar",
                exact=True,
            ).locator("xpath=ancestor::button[1]"),
        ],
        "botão Pesquisar",
    )


def localizar_botao_imprimir_ogea(page):
    return primeiro_elemento_visivel_ogea(
        page,
        [
            # Seletor exato gravado pelo Playwright Codegen.
            page.get_by_role(
                "button",
                name=" Imprimir",
                exact=True,
            ),
            page.locator("button").filter(
                has_text=re.compile(r"Imprimir", re.I)
            ),
            page.get_by_text(
                "Imprimir",
                exact=True,
            ).locator("xpath=ancestor::button[1]"),
        ],
        "botão Imprimir",
    )


def localizar_checkbox_todos_ogea(
    page,
    timeout_segundos=None,
):
    if timeout_segundos is None:
        timeout_segundos = PDF_TIMEOUT_SEGUNDOS

    return primeiro_elemento_visivel_ogea(
        page,
        [
            page.get_by_role(
                "checkbox",
                name="Marcar/Desmarcar Todos",
                exact=True,
            ),
            page.locator(
                "input[type='checkbox'][title*='Marcar']"
            ),
        ],
        "caixa Marcar/Desmarcar Todos",
        timeout_segundos=timeout_segundos,
    )


def aguardar_resultado_ogea(page):
    return localizar_checkbox_todos_ogea(
        page,
        timeout_segundos=PDF_TIMEOUT_SEGUNDOS,
    )


def salvar_pdf_ogea_pela_requisicao(
    contexto,
    page,
    botao_imprimir,
    caminho_saida,
    descricao,
):
    """
    Captura a requisição real usada pelo OGEA para gerar o PDF.

    Diferente da Mobyan, o OGEA envia as OSs em uma requisição POST para
    /serviceOrderPrint/print. A barra de endereço mostra apenas a URL,
    mas o conteúdo do PDF depende do corpo dessa requisição.

    A função:
    1. observa a requisição antes do clique;
    2. captura método, corpo e cabeçalhos reais;
    3. tenta salvar a resposta original;
    4. se necessário, reexecuta a mesma Request dentro da sessão atual.
    """
    requisicoes = []

    def registrar_requisicao(request):
        try:
            if "/serviceOrderPrint/print" in request.url:
                requisicoes.append(request)
        except Exception:
            pass

    contexto.on("request", registrar_requisicao)

    popup = None

    try:
        with page.expect_popup(
            timeout=PDF_TIMEOUT_SEGUNDOS * 1000
        ) as popup_info:
            botao_imprimir.click()

        popup = popup_info.value

        limite = time.time() + 20

        while time.time() < limite and not requisicoes:
            popup.wait_for_timeout(250)

        if not requisicoes:
            raise RuntimeError(
                f"{descricao}: a janela abriu, mas não capturei a "
                "requisição /serviceOrderPrint/print."
            )

        # Dá prioridade à requisição POST, pois ela contém as OSs selecionadas.
        requisicao = next(
            (
                item
                for item in requisicoes
                if item.method.upper() == "POST"
            ),
            requisicoes[-1],
        )

        print(
            f"{descricao}: requisição capturada "
            f"({requisicao.method} {requisicao.url})."
        )

        # Percorre eventual cadeia de redirecionamentos.
        cadeia = [requisicao]
        atual = requisicao

        while True:
            try:
                proxima = atual.redirected_to
            except Exception:
                proxima = None

            if proxima is None:
                break

            cadeia.append(proxima)
            atual = proxima

        # Primeiro tenta usar a resposta que já gerou o PDF visível.
        for item in reversed(cadeia):
            try:
                resposta = item.response()

                if resposta is None:
                    continue

                conteudo = resposta.body()
                content_type = str(
                    resposta.headers.get("content-type", "")
                ).lower()

                if conteudo.startswith(b"%PDF"):
                    caminho_saida.write_bytes(conteudo)

                    print(
                        f"{descricao}: PDF salvo da resposta original "
                        f"({len(conteudo):,} bytes, "
                        f"content-type={content_type or 'n/a'})."
                    )
                    return caminho_saida

            except Exception:
                continue

        # Se o Chromium não disponibilizar o corpo original, reexecuta
        # exatamente a mesma Request. O APIRequestContext do BrowserContext
        # compartilha os cookies da sessão autenticada.
        print(
            f"{descricao}: resposta original não ficou acessível; "
            "reexecutando a mesma requisição POST..."
        )

        resposta_api = contexto.request.fetch(
            requisicao,
            timeout=PDF_TIMEOUT_SEGUNDOS * 1000,
            fail_on_status_code=False,
        )

        conteudo = resposta_api.body()
        content_type = str(
            resposta_api.headers.get("content-type", "")
        ).lower()

        if not conteudo.startswith(b"%PDF"):
            tamanho_post = 0

            try:
                tamanho_post = len(
                    requisicao.post_data_buffer or b""
                )
            except Exception:
                pass

            raise RuntimeError(
                f"{descricao}: a requisição foi repetida, mas não retornou "
                f"PDF. status={resposta_api.status}, "
                f"content-type={content_type or 'n/a'}, "
                f"tamanho={len(conteudo)} bytes, "
                f"método={requisicao.method}, "
                f"post_data={tamanho_post} bytes."
            )

        caminho_saida.write_bytes(conteudo)

        print(
            f"{descricao}: PDF original salvo pela requisição capturada "
            f"({len(conteudo):,} bytes)."
        )

        return caminho_saida

    finally:
        try:
            contexto.remove_listener(
                "request",
                registrar_requisicao,
            )
        except Exception:
            pass

        if popup is not None:
            try:
                popup.close()
            except Exception:
                pass

def gerar_pdfs_ogea(listas, pasta_destino=PASTA_PDFS):
    tarefas = [
        item
        for item in listas
        if item["ogea"]
    ]

    if not tarefas:
        print("Nenhum PDF do OGEA necessário.")
        return []

    validar_ogea()

    print("")
    print("=" * 70)
    print("GERANDO PDFs DO OGEA")
    print("=" * 70)

    gerados = []

    with sync_playwright() as playwright:
        navegador = playwright.chromium.launch(headless=False)
        contexto = navegador.new_context(
            accept_downloads=True,
            viewport={"width": 1600, "height": 900},
        )
        page = contexto.new_page()
        page.set_default_timeout(PDF_TIMEOUT_SEGUNDOS * 1000)
        page.set_default_navigation_timeout(
            PDF_TIMEOUT_SEGUNDOS * 1000
        )

        try:
            fazer_login_ogea(page)

            for indice, item in enumerate(tarefas, start=1):
                tecnico = item["tecnico"]
                oss = item["ogea"]
                nome = limpar_nome_arquivo(tecnico)
                caminho = pasta_destino / f"OGEA - {nome}.pdf"

                print("")
                print(
                    f"[{indice}/{len(tarefas)}] OGEA - {tecnico}: "
                    f"{len(oss)} OS(s)."
                )

                abrir_ordem_servico_ogea(page)

                campo_oss = page.locator("#serviceOrderCodes")
                esperar_elemento_ogea(
                    campo_oss,
                    "campo de números das OSs",
                    timeout_segundos=PDF_TIMEOUT_SEGUNDOS,
                )

                campo_oss.fill(",".join(oss))
                campo_oss.press("Enter")
                page.wait_for_timeout(750)

                # Em algumas respostas lentas do OGEA, o Enter pode iniciar
                # a pesquisa. Se os resultados já apareceram, não tenta clicar
                # novamente no botão Pesquisar.
                checkbox_todos = None

                try:
                    checkbox_todos = localizar_checkbox_todos_ogea(
                        page,
                        timeout_segundos=3,
                    )
                    print(
                        "A pesquisa foi iniciada ao confirmar as OSs."
                    )
                except Exception:
                    botao_pesquisar = localizar_botao_pesquisar_ogea(
                        page
                    )
                    botao_pesquisar.scroll_into_view_if_needed()
                    botao_pesquisar.click()
                    checkbox_todos = aguardar_resultado_ogea(page)

                try:
                    if checkbox_todos.is_checked():
                        checkbox_todos.uncheck()
                except Exception:
                    pass

                checkbox_todos.check()

                botao_imprimir = localizar_botao_imprimir_ogea(page)
                botao_imprimir.scroll_into_view_if_needed()

                salvar_pdf_ogea_pela_requisicao(
                    contexto,
                    page,
                    botao_imprimir,
                    caminho,
                    f"OGEA - {tecnico}",
                )

                gerados.append(caminho)

        except Exception:
            salvar_screenshot(page, "erro_pdfs_ogea")
            raise
        finally:
            contexto.close()
            navegador.close()

    return gerados


def validar_pypdf():
    if PdfReader is None or PdfWriter is None:
        raise RuntimeError(
            "A dependência pypdf ainda não está instalada. "
            "Execute INSTALAR_MAC.command para instalar a dependência pypdf."
        )


def quantidade_paginas_pdf(caminho):
    leitor = PdfReader(str(caminho))
    return len(leitor.pages)


def unir_pdfs_por_tecnico(listas):
    """
    Cria um único arquivo final por técnico.

    Ordem do arquivo:
    1. todas as OSs da Mobyan;
    2. todas as OSs do OGEA.

    Técnicos que possuírem somente uma das operações também recebem
    um arquivo final com o padrão Roteiro - Técnico.pdf.
    """
    validar_pypdf()

    print("")
    print("=" * 70)
    print("UNINDO PDFs POR TÉCNICO")
    print("=" * 70)

    roteiros = []

    for indice, item in enumerate(listas, start=1):
        tecnico = item["tecnico"]
        nome = limpar_nome_arquivo(tecnico)

        partes = [
            PASTA_PDFS_TEMP / f"Mobyan - {nome}.pdf",
            PASTA_PDFS_TEMP / f"OGEA - {nome}.pdf",
        ]
        partes = [
            caminho
            for caminho in partes
            if caminho.exists() and caminho.stat().st_size > 0
        ]

        if not partes:
            print(
                f"[{indice}/{len(listas)}] {tecnico}: "
                "nenhum PDF encontrado para unir."
            )
            continue

        destino = PASTA_PDFS / f"Roteiro - {nome}.pdf"
        escritor = PdfWriter()
        total_paginas = 0

        for parte in partes:
            leitor = PdfReader(str(parte))

            for pagina in leitor.pages:
                escritor.add_page(pagina)
                total_paginas += 1

        escritor.add_metadata({
            "/Title": f"Roteiro - {tecnico}",
            "/Subject": "Ordens de Serviço Mobyan e OGEA",
            "/Creator": "Central Mobyan",
        })

        with destino.open("wb") as arquivo_saida:
            escritor.write(arquivo_saida)

        # Validação final: arquivo PDF legível e com a quantidade esperada.
        paginas_validas = quantidade_paginas_pdf(destino)

        if paginas_validas != total_paginas:
            raise RuntimeError(
                f"Roteiro - {tecnico}: eram esperadas "
                f"{total_paginas} páginas, mas o arquivo final possui "
                f"{paginas_validas}."
            )

        origens = " + ".join(
            parte.name.split(" - ", 1)[0]
            for parte in partes
        )

        print(
            f"[{indice}/{len(listas)}] Roteiro - {tecnico}: "
            f"{total_paginas} página(s) ({origens})."
        )

        roteiros.append(destino)

    if not roteiros:
        raise RuntimeError(
            "Nenhum roteiro final foi criado."
        )

    # Os arquivos separados servem apenas como intermediários.
    shutil.rmtree(PASTA_PDFS_TEMP, ignore_errors=True)

    return roteiros

def abrir_pasta(caminho):
    try:
        if os.name == "nt":
            os.startfile(str(caminho))
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", str(caminho)])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", str(caminho)])
    except Exception:
        pass


def main():
    parser = argparse.ArgumentParser(
        description="Gera um roteiro PDF unificado por técnico."
    )
    parser.add_argument(
        "--sem-abrir",
        action="store_true",
        help="Não abre a pasta de PDFs ao terminar.",
    )

    grupo = parser.add_mutually_exclusive_group()
    grupo.add_argument(
        "--somente-ogea",
        action="store_true",
        help="Gera somente os PDFs do OGEA.",
    )
    grupo.add_argument(
        "--somente-mobyan",
        action="store_true",
        help="Gera somente os PDFs da Mobyan.",
    )
    parser.add_argument(
        "--tecnico",
        default="",
        help="Gera somente o técnico informado.",
    )

    args = parser.parse_args()

    if args.somente_ogea:
        limpar_pdfs_antigos("OGEA")
    elif args.somente_mobyan:
        limpar_pdfs_antigos("Mobyan")
    else:
        limpar_pdfs_antigos()

    listas = ler_listas_pdf()

    if args.tecnico:
        tecnico_procurado = normalizar_texto(
            args.tecnico
        ).casefold()

        listas = [
            item
            for item in listas
            if normalizar_texto(
                item["tecnico"]
            ).casefold() == tecnico_procurado
        ]

        if not listas:
            raise RuntimeError(
                f"Não encontrei o técnico: {args.tecnico}"
            )

    total_mobyan = sum(len(item["mobyan"]) for item in listas)
    total_ogea = sum(len(item["ogea"]) for item in listas)

    print(f"Técnicos encontrados: {len(listas)}")
    print(f"OSs Mobyan: {total_mobyan}")
    print(f"OSs OGEA: {total_ogea}")

    gerados = []

    # Modos de teste continuam gerando o arquivo separado da origem.
    if args.somente_ogea:
        gerados.extend(
            gerar_pdfs_ogea(
                listas,
                pasta_destino=PASTA_PDFS,
            )
        )
    elif args.somente_mobyan:
        gerados.extend(
            gerar_pdfs_mobyan(
                listas,
                pasta_destino=PASTA_PDFS,
            )
        )
    else:
        # Execução normal: os PDFs separados ficam temporariamente em
        # _temporarios e são unidos em um único roteiro por técnico.
        gerar_pdfs_mobyan(
            listas,
            pasta_destino=PASTA_PDFS_TEMP,
        )
        gerar_pdfs_ogea(
            listas,
            pasta_destino=PASTA_PDFS_TEMP,
        )
        gerados = unir_pdfs_por_tecnico(listas)

    print("")
    print("=" * 70)

    if args.somente_ogea or args.somente_mobyan:
        print(f"PDFs de teste gerados: {len(gerados)}")
    else:
        print(f"Roteiros finais gerados: {len(gerados)}")

    print(f"Pasta final: {PASTA_PDFS}")
    print("=" * 70)

    for caminho in gerados:
        print(caminho.name)

    if not args.sem_abrir:
        abrir_pasta(PASTA_PDFS)

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as erro:
        print("")
        print("ERRO AO GERAR PDFs")
        print(str(erro))
        raise
