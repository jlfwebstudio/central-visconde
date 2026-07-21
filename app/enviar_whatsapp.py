import os
import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd
from PIL import Image
from playwright.sync_api import sync_playwright

from caminho_base import BASE_DIR
ARQUIVO_PLANILHA = BASE_DIR / "outputs" / "pendencias_do_dia" / "pendencias_do_dia_atual.xlsx"
PASTA_PERFIL_WHATSAPP = BASE_DIR / "whatsapp_profile"
PASTA_TEMP_WHATSAPP = BASE_DIR / "outputs" / "whatsapp_temp"

# Modelo recebido pela Central Mobyan:
# "manha" ou "acompanhamento"
MODELO_MENSAGEM = os.getenv("MODELO_MENSAGEM", "manha").strip().lower()

# Destino recebido pela Central Mobyan (tela "Para quem é esse envio?"):
# "bases" (prestadores, aba "Envios") ou "tecnicos" (RS-SMART, aba "Envios Técnicos")
DESTINO_ENVIO = os.getenv("DESTINO_ENVIO", "bases").strip().lower()

# True = envia automaticamente
# False = prepara o envio, mas você envia manualmente
ENVIAR_DE_VERDADE = True

# None = envia para todos os prestadores com Status Envio = Pronto para envio e Enviar = Sim
# Exemplo: 2 = envia somente para os 2 primeiros, caso queira testar novamente
LIMITE_ENVIOS = None

# Tempos de segurança para o WhatsApp concluir o envio
TEMPO_APOS_ENVIO = 12
TEMPO_ENTRE_CONTATOS = 10

# Tentativas ao abrir conversa, para evitar falhas por oscilação do WhatsApp Web
TENTATIVAS_ABRIR_CONVERSA = 3
TEMPO_ENTRE_TENTATIVAS = 10


def obter_coluna_mensagem():
    modelo = MODELO_MENSAGEM.strip().lower()

    if modelo in ["acompanhamento", "operacao", "operação", "tarde", "followup"]:
        return "Mensagem Acompanhamento"

    return "Mensagem Manhã"


def ler_fila_envios():
    if not ARQUIVO_PLANILHA.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {ARQUIVO_PLANILHA}")

    df = pd.read_excel(
        ARQUIVO_PLANILHA,
        sheet_name="Envios",
        dtype=str,
        keep_default_na=False,
        na_filter=False,
    )

    # Evita erro quando o Excel salva cabeçalhos com espaços acidentais.
    df.columns = [str(coluna).strip() for coluna in df.columns]

    coluna_mensagem = obter_coluna_mensagem()

    if coluna_mensagem not in df.columns:
        if coluna_mensagem == "Mensagem Manhã" and "Mensagem" in df.columns:
            coluna_mensagem = "Mensagem"
        else:
            raise Exception(
                f"Coluna de mensagem não encontrada na aba Envios: {coluna_mensagem}. "
                "Gere as pendências novamente pela Central Mobyan."
            )

    colunas_obrigatorias = [
        "Prestador",
        "Responsável",
        "WhatsApp",
        "Enviar",
        "Imagem",
        "Status Envio",
        coluna_mensagem,
    ]

    for coluna in colunas_obrigatorias:
        if coluna not in df.columns:
            raise Exception(f"Coluna obrigatória não encontrada na aba Envios: {coluna}")

    for coluna in df.columns:
        df[coluna] = df[coluna].astype(str).str.strip()

    df["Enviar_Normalizado"] = (
        df["Enviar"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    df["_Mensagem_Selecionada"] = (
        df[coluna_mensagem]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df = df[df["Enviar_Normalizado"].isin(["sim", "s", "yes", "y"])].copy()
    df = df[df["Status Envio"] == "Pronto para envio"].copy()
    df = df[df["WhatsApp"] != ""].copy()
    df = df[df["Imagem"] != ""].copy()
    df = df[df["_Mensagem_Selecionada"] != ""].copy()

    df = df.drop(columns=["Enviar_Normalizado"])

    print(f"Modelo de mensagem selecionado: {MODELO_MENSAGEM}")
    print(f"Coluna usada para envio: {coluna_mensagem}")

    return df


def ler_fila_envios_tecnicos():
    """Mesmo formato de ler_fila_envios(), mas pra fila de técnicos (RS-SMART
    matriz — aba "Envios Técnicos", ver exportador_mobyan.py). Tolerante à
    ausência da aba: uma planilha gerada antes dessa atualização, ou uma conta
    sem nenhum técnico com WhatsApp cadastrado, simplesmente não tem envio de
    técnico nenhum, sem quebrar o envio das bases."""
    if not ARQUIVO_PLANILHA.exists():
        return pd.DataFrame(columns=["Prestador", "Responsável", "WhatsApp", "Imagem", "_Mensagem_Selecionada"])

    try:
        df = pd.read_excel(
            ARQUIVO_PLANILHA,
            sheet_name="Envios Técnicos",
            dtype=str,
            keep_default_na=False,
            na_filter=False,
        )
    except Exception:
        return pd.DataFrame(columns=["Prestador", "Responsável", "WhatsApp", "Imagem", "_Mensagem_Selecionada"])

    df.columns = [str(coluna).strip() for coluna in df.columns]

    coluna_mensagem = obter_coluna_mensagem()

    if coluna_mensagem not in df.columns:
        if coluna_mensagem == "Mensagem Manhã" and "Mensagem" in df.columns:
            coluna_mensagem = "Mensagem"
        else:
            return pd.DataFrame(columns=["Prestador", "Responsável", "WhatsApp", "Imagem", "_Mensagem_Selecionada"])

    colunas_obrigatorias = ["Técnico", "WhatsApp", "Imagem", "Status Envio", coluna_mensagem]

    for coluna in colunas_obrigatorias:
        if coluna not in df.columns:
            return pd.DataFrame(columns=["Prestador", "Responsável", "WhatsApp", "Imagem", "_Mensagem_Selecionada"])

    for coluna in df.columns:
        df[coluna] = df[coluna].astype(str).str.strip()

    df["_Mensagem_Selecionada"] = df[coluna_mensagem].fillna("").astype(str).str.strip()

    df = df[df["Status Envio"] == "Pronto para envio"].copy()
    df = df[df["WhatsApp"] != ""].copy()
    df = df[df["Imagem"] != ""].copy()
    df = df[df["_Mensagem_Selecionada"] != ""].copy()

    # Normaliza pro mesmo formato que processar_envio já espera, sem precisar
    # tocar em processar_envio/abrir_conversa/anexar_imagem/clicar_enviar.
    df["Prestador"] = "Técnico: " + df["Técnico"]
    df["Responsável"] = df["Técnico"]

    print(f"Envios de técnico prontos: {len(df)}")

    return df[["Prestador", "Responsável", "WhatsApp", "Imagem", "_Mensagem_Selecionada"]]


def aguardar_login_whatsapp(page):
    print("Abrindo WhatsApp Web...")

    page.goto("https://web.whatsapp.com/", wait_until="domcontentloaded", timeout=90000)

    print("")
    print("Se aparecer QR Code, escaneie com o WhatsApp do celular.")
    print("Aguardando WhatsApp Web carregar...")

    seletores_carregamento = [
        'div[aria-label="Lista de conversas"]',
        'div[aria-label="Chat list"]',
        'div[aria-label="Lista de chats"]',
        'div[aria-label="Pesquisar ou começar uma nova conversa"]',
        'div[title="Pesquisar ou começar uma nova conversa"]',
        'button[aria-label="Nova conversa"]',
        'span[data-icon="new-chat-outline"]',
        'span[data-icon="menu"]',
        'div[role="grid"]',
        'div[role="textbox"]',
    ]

    tempo_limite_segundos = 120
    inicio = time.time()

    while time.time() - inicio < tempo_limite_segundos:
        for seletor in seletores_carregamento:
            try:
                if page.locator(seletor).first.is_visible(timeout=1000):
                    print("WhatsApp Web carregado.")
                    return
            except Exception:
                pass

        page.wait_for_timeout(1000)

    raise Exception(
        "WhatsApp Web não carregou a tempo. "
        "Confira se o QR Code foi escaneado e se a tela principal do WhatsApp apareceu."
    )


def abrir_conversa(page, whatsapp, mensagem):
    mensagem_url = quote(mensagem)
    url = f"https://web.whatsapp.com/send?phone={whatsapp}&text={mensagem_url}"

    for tentativa in range(1, TENTATIVAS_ABRIR_CONVERSA + 1):
        try:
            print(
                f"Abrindo conversa do número {whatsapp}... "
                f"tentativa {tentativa}/{TENTATIVAS_ABRIR_CONVERSA}"
            )

            page.goto(url, wait_until="domcontentloaded", timeout=90000)

            print("Aguardando conversa carregar...")

            seletores_conversa = [
                'div[role="textbox"]',
                'div[contenteditable="true"]',
                'span[data-icon="send"]',
                'button[aria-label="Enviar"]',
                'button[aria-label="Send"]',
            ]

            tempo_limite_segundos = 60
            inicio = time.time()

            while time.time() - inicio < tempo_limite_segundos:
                for seletor in seletores_conversa:
                    try:
                        if page.locator(seletor).last.is_visible(timeout=1000):
                            page.wait_for_timeout(3000)
                            print("Conversa carregada.")
                            return
                    except Exception:
                        pass

                page.wait_for_timeout(1000)

            raise Exception("Conversa não carregou dentro do tempo esperado.")

        except Exception as erro:
            print("")
            print(f"Aviso: falha ao abrir conversa na tentativa {tentativa}/{TENTATIVAS_ABRIR_CONVERSA}.")
            print(f"Erro: {erro}")

            if tentativa < TENTATIVAS_ABRIR_CONVERSA:
                print(f"Aguardando {TEMPO_ENTRE_TENTATIVAS} segundos e tentando novamente...")

                page.wait_for_timeout(TEMPO_ENTRE_TENTATIVAS * 1000)

                try:
                    page.goto(
                        "https://web.whatsapp.com/",
                        wait_until="domcontentloaded",
                        timeout=90000
                    )
                    page.wait_for_timeout(5000)
                except Exception:
                    pass

            else:
                raise Exception(
                    f"Não consegui abrir a conversa do WhatsApp: {whatsapp}. "
                    "Confira se o número está correto ou tente rodar novamente."
                )


def clicar_botao_anexo(page):
    seletores = [
        'span[data-icon="plus"]',
        'span[data-icon="clip"]',
        'button[aria-label="Anexar"]',
        'button[aria-label="Attach"]',
        'div[aria-label="Anexar"]',
        'div[aria-label="Attach"]',
    ]

    for seletor in seletores:
        try:
            elemento = page.locator(seletor).first
            elemento.click(timeout=3000)
            page.wait_for_timeout(1000)
            return True
        except Exception:
            pass

    return False


def converter_imagem_para_jpg(caminho_imagem):
    caminho = Path(caminho_imagem).resolve()

    if not caminho.exists():
        raise FileNotFoundError(f"Imagem não encontrada: {caminho}")

    PASTA_TEMP_WHATSAPP.mkdir(parents=True, exist_ok=True)

    caminho_jpg = PASTA_TEMP_WHATSAPP / f"{caminho.stem}.jpg"

    imagem = Image.open(caminho)

    if imagem.mode in ("RGBA", "LA"):
        fundo = Image.new("RGB", imagem.size, "white")
        fundo.paste(imagem, mask=imagem.split()[-1])
        imagem = fundo
    else:
        imagem = imagem.convert("RGB")

    imagem.save(caminho_jpg, "JPEG", quality=95)

    return caminho_jpg


def anexar_imagem(page, caminho_imagem):
    caminho_jpg = converter_imagem_para_jpg(caminho_imagem)

    print(f"Anexando imagem pela opção Fotos e vídeos: {caminho_jpg}")

    clicou = clicar_botao_anexo(page)

    if not clicou:
        raise Exception("Não encontrei o botão de anexo do WhatsApp.")

    page.wait_for_timeout(1000)

    opcoes_fotos = [
        "Fotos e vídeos",
        "Fotos e videos",
        "Photos & videos",
        "Photos and videos",
    ]

    arquivo_selecionado = False
    ultimo_erro = None

    for texto_opcao in opcoes_fotos:
        try:
            print(f"Tentando opção: {texto_opcao}")

            with page.expect_file_chooser(timeout=5000) as file_chooser_info:
                page.get_by_text(texto_opcao, exact=True).click(timeout=3000)

            file_chooser = file_chooser_info.value
            file_chooser.set_files(str(caminho_jpg))

            arquivo_selecionado = True
            break

        except Exception as erro:
            ultimo_erro = erro

    if not arquivo_selecionado:
        raise Exception(
            "Não consegui clicar na opção 'Fotos e vídeos'. "
            f"Último erro: {ultimo_erro}"
        )

    print("Imagem anexada pela opção Fotos e vídeos. Aguardando prévia carregar...")
    page.wait_for_timeout(5000)


def clicar_enviar(page):
    seletores = [
        'span[data-icon="send"]',
        'span[data-icon="send-light"]',
        'button[aria-label="Enviar"]',
        'button[aria-label="Send"]',
        'div[aria-label="Enviar"]',
        'div[aria-label="Send"]',
        'button:has(span[data-icon="send"])',
        'button:has(span[data-icon="send-light"])',
    ]

    for seletor in seletores:
        try:
            elementos = page.locator(seletor)
            quantidade = elementos.count()

            if quantidade > 0:
                elementos.last.click(timeout=5000)
                page.wait_for_timeout(3000)
                return True
        except Exception:
            pass

    try:
        largura = page.viewport_size["width"]
        altura = page.viewport_size["height"]

        page.mouse.click(largura - 55, altura - 55)
        page.wait_for_timeout(3000)
        return True
    except Exception:
        pass

    return False


def enviar_mensagem_com_imagem(page, imagem):
    anexar_imagem(page, imagem)

    print("Enviando imagem com legenda...")
    if not clicar_enviar(page):
        raise Exception("Não consegui clicar no botão de enviar imagem com legenda.")

    print(f"Aguardando {TEMPO_APOS_ENVIO} segundos para confirmar o envio...")
    page.wait_for_timeout(TEMPO_APOS_ENVIO * 1000)

    print("Imagem e mensagem enviadas.")


def processar_envio(page, linha):
    prestador = linha["Prestador"]
    responsavel = linha["Responsável"]
    whatsapp = linha["WhatsApp"]
    imagem = linha["Imagem"]
    mensagem = linha["_Mensagem_Selecionada"]

    print("")
    print("=" * 80)
    print(f"Prestador: {prestador}")
    print(f"Responsável: {responsavel}")
    print(f"WhatsApp: {whatsapp}")
    print(f"Modelo: {MODELO_MENSAGEM}")
    print("=" * 80)

    abrir_conversa(page, whatsapp, mensagem)

    if ENVIAR_DE_VERDADE:
        enviar_mensagem_com_imagem(page, imagem)
        print("Envio concluído.")
    else:
        print("MODO TESTE: mensagem foi preparada na conversa, mas não será enviada automaticamente.")
        anexar_imagem(page, imagem)

        print("")
        print("Confira no WhatsApp Web:")
        print("- contato correto")
        print("- mensagem correta")
        print("- imagem correta")
        print("")
        print("Como ENVIAR_DE_VERDADE está False, clique manualmente em enviar se estiver tudo certo.")


def enviar_whatsapp():
    if DESTINO_ENVIO == "tecnicos":
        df_envios = ler_fila_envios_tecnicos()
    else:
        df_envios = ler_fila_envios()

    if df_envios.empty:
        print(f"Nenhum envio encontrado para o destino '{DESTINO_ENVIO}' com Status Envio = Pronto para envio.")
        return

    if LIMITE_ENVIOS is not None:
        df_envios = df_envios.head(LIMITE_ENVIOS)

    print(f"Total de envios a processar: {len(df_envios)}")

    with sync_playwright() as p:
        contexto = p.chromium.launch_persistent_context(
            user_data_dir=str(PASTA_PERFIL_WHATSAPP),
            headless=False,
            viewport={"width": 1366, "height": 768},
            accept_downloads=True,
        )

        page = contexto.new_page()

        aguardar_login_whatsapp(page)

        for _, linha in df_envios.iterrows():
            processar_envio(page, linha)

            print(f"Aguardando {TEMPO_ENTRE_CONTATOS} segundos antes do próximo envio...")
            time.sleep(TEMPO_ENTRE_CONTATOS)

            if not ENVIAR_DE_VERDADE:
                print("")
                print("Teste finalizado no primeiro contato.")
                print("O navegador ficará aberto para você conferir.")
                input("Depois de conferir, pressione ENTER aqui no Terminal para encerrar o teste...")
                break

        print(f"Aguardando {TEMPO_APOS_ENVIO} segundos antes de fechar o navegador...")
        page.wait_for_timeout(TEMPO_APOS_ENVIO * 1000)

        contexto.close()

    print("Processo de WhatsApp finalizado.")


if __name__ == "__main__":
    enviar_whatsapp()