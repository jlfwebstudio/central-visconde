import os
import re
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

BASE_DIR = Path(__file__).resolve().parent.parent
ARQUIVO_PLANILHA = BASE_DIR / "outputs" / "pendencias_do_dia" / "pendencias_do_dia_atual.xlsx"
PASTA_IMAGENS = BASE_DIR / "outputs" / "por_prestador_imagens"

COLUNAS_IMAGEM = [
    "SITUAÇÃO",
    "Chamado",
    "Numero Referencia",
    "Contratante",
    "Serviço",
    "Status",
    "Data Limite",
    "Cliente",
    "CNPJ / CPF",
    "Cidade",
]

LARGURAS = {
    "SITUAÇÃO": 110,
    "Chamado": 110,
    "Numero Referencia": 150,
    "Contratante": 120,
    "Serviço": 180,
    "Status": 150,
    "Data Limite": 170,
    "Cliente": 300,
    "CNPJ / CPF": 170,
    "Cidade": 180,
}

ALTURA_CABECALHO = 34
ALTURA_LINHA = 32
MARGEM = 16

COR_CABECALHO = "#1F4E78"
COR_TEXTO_CABECALHO = "#FFFFFF"
COR_LINHA_VENCIDA = "#F4CCCC"
COR_LINHA_HOJE = "#FFF2CC"
COR_LINHA_FUTURA = "#D9EAD3"
COR_SITUACAO_VENCIDA = "#CC0000"
COR_SITUACAO_HOJE = "#F1C232"
COR_SITUACAO_FUTURA = "#70AD47"
COR_TEXTO = "#000000"
COR_TEXTO_BRANCO = "#FFFFFF"
COR_BORDA = "#D9E2F3"
COR_FUNDO = "#FFFFFF"


def carregar_fonte(tamanho=14, negrito=False):
    caminhos = []

    pasta_windows = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"

    if negrito:
        caminhos.extend([
            str(pasta_windows / "arialbd.ttf"),
            str(pasta_windows / "calibrib.ttf"),
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold Unicode.ttf",
            "/Library/Fonts/Arial Bold.ttf",
        ])
    else:
        caminhos.extend([
            str(pasta_windows / "arial.ttf"),
            str(pasta_windows / "calibri.ttf"),
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
            "/Library/Fonts/Arial.ttf",
        ])

    for caminho in caminhos:
        try:
            return ImageFont.truetype(caminho, tamanho)
        except Exception:
            pass

    return ImageFont.load_default()


FONTE_CABECALHO = carregar_fonte(14, negrito=True)
FONTE_TEXTO = carregar_fonte(13, negrito=False)
FONTE_TEXTO_NEGRITO = carregar_fonte(13, negrito=True)


def limpar_nome_arquivo(texto):
    texto = "" if texto is None else str(texto).strip()

    substituicoes = {
        "/": "-",
        "\\": "-",
        ":": "-",
        "*": "",
        "?": "",
        '"': "",
        "<": "",
        ">": "",
        "|": "",
    }

    for antigo, novo in substituicoes.items():
        texto = texto.replace(antigo, novo)

    texto = re.sub(r"\s+", " ", texto).strip()

    if texto == "":
        texto = "SEM_PRESTADOR"

    return texto


def limpar_valor(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() in ["nan", "none", "null", "nat"]:
        return ""

    return texto


def normalizar_nome_coluna(nome):
    texto = str(nome).strip().lower()
    texto = texto.replace("\n", " ")
    texto = texto.replace("\r", " ")
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.replace(" / ", "/")
    texto = texto.replace(" /", "/")
    texto = texto.replace("/ ", "/")
    return texto


def encontrar_coluna_real(df, nomes_possiveis):
    colunas_normalizadas = {
        normalizar_nome_coluna(coluna): coluna
        for coluna in df.columns
    }

    for nome in nomes_possiveis:
        nome_normalizado = normalizar_nome_coluna(nome)

        if nome_normalizado in colunas_normalizadas:
            return colunas_normalizadas[nome_normalizado]

    return None


def texto_curto(texto, fonte, largura_maxima):
    texto = limpar_valor(texto)

    if texto == "":
        return ""

    draw_temp = ImageDraw.Draw(Image.new("RGB", (1, 1)))

    if draw_temp.textlength(texto, font=fonte) <= largura_maxima:
        return texto

    sufixo = "..."
    texto_base = texto

    while len(texto_base) > 0:
        tentativa = texto_base + sufixo

        if draw_temp.textlength(tentativa, font=fonte) <= largura_maxima:
            return tentativa

        texto_base = texto_base[:-1]

    return sufixo


def desenhar_texto_centralizado(draw, caixa, texto, fonte, cor):
    x1, y1, x2, y2 = caixa
    texto = limpar_valor(texto)

    bbox = draw.textbbox((0, 0), texto, font=fonte)
    largura_texto = bbox[2] - bbox[0]
    altura_texto = bbox[3] - bbox[1]

    x = x1 + ((x2 - x1) - largura_texto) / 2
    y = y1 + ((y2 - y1) - altura_texto) / 2 - 1

    draw.text((x, y), texto, font=fonte, fill=cor)


def desenhar_texto_esquerda(draw, caixa, texto, fonte, cor):
    x1, y1, x2, y2 = caixa
    texto = texto_curto(texto, fonte, (x2 - x1) - 10)

    bbox = draw.textbbox((0, 0), texto, font=fonte)
    altura_texto = bbox[3] - bbox[1]

    x = x1 + 5
    y = y1 + ((y2 - y1) - altura_texto) / 2 - 1

    draw.text((x, y), texto, font=fonte, fill=cor)


def cor_linha_por_situacao(situacao):
    situacao = limpar_valor(situacao).upper()

    if situacao == "VENCIDA":
        return COR_LINHA_VENCIDA

    if situacao == "VENCE HOJE":
        return COR_LINHA_HOJE

    return COR_LINHA_FUTURA


def cor_situacao(situacao):
    situacao = limpar_valor(situacao).upper()

    if situacao == "VENCIDA":
        return COR_SITUACAO_VENCIDA, COR_TEXTO_BRANCO

    if situacao == "VENCE HOJE":
        return COR_SITUACAO_HOJE, COR_TEXTO

    return COR_SITUACAO_FUTURA, COR_TEXTO_BRANCO


def preparar_dataframe():
    if not ARQUIVO_PLANILHA.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {ARQUIVO_PLANILHA}")

    df = pd.read_excel(
        ARQUIVO_PLANILHA,
        sheet_name="Pendências",
        dtype=str,
        keep_default_na=False,
        na_filter=False,
    )

    df.columns = [str(coluna).strip() for coluna in df.columns]

    mapa_colunas = {
        "SITUAÇÃO": [
            "SITUAÇÃO",
            "SITUACAO",
            "Situação",
            "Situacao",
        ],
        "Chamado": [
            "Chamado",
            "OS",
            "Ordem de Serviço",
            "Ordem de Servico",
        ],
        "Numero Referencia": [
            "Numero Referencia",
            "Número Referência",
            "Referencia",
            "Referência",
            "Nº Referência",
            "N Referencia",
        ],
        "Contratante": [
            "Contratante",
        ],
        "Serviço": [
            "Serviço",
            "Servico",
            "Tipo Serviço",
            "Tipo Servico",
        ],
        "Status": [
            "Status",
            "Situação Status",
            "Situacao Status",
        ],
        "Data Limite": [
            "Data Limite",
            "Data Limite SLA",
            "Limite",
            "Prazo",
        ],
        "Cliente": [
            "Cliente",
            "Nome Cliente",
            "Nome do Cliente",
            "Cliente Nome",
            "Razão Social",
            "Razao Social",
            "Nome Fantasia",
            "Estabelecimento",
            "EC",
            "Nome EC",
            "Nome do EC",
            "Ponto de Atendimento",
            "PA",
            "Nome PA",
        ],
        "CNPJ / CPF": [
            "CNPJ / CPF",
            "CNPJ/CPF",
            "CNPJ CPF",
            "CPF / CNPJ",
            "CPF/CNPJ",
            "Documento",
            "Documento Cliente",
            "CPF",
            "CNPJ",
        ],
        "Cidade": [
            "Cidade",
            "Município",
            "Municipio",
        ],
        "Prestador": [
            "Prestador",
            "Base",
            "Fornecedor",
        ],
    }

    for coluna_padrao, aliases in mapa_colunas.items():
        coluna_real = encontrar_coluna_real(df, aliases)

        if coluna_real:
            df[coluna_padrao] = df[coluna_real]
        else:
            df[coluna_padrao] = ""

    for coluna in df.columns:
        df[coluna] = (
            df[coluna]
            .astype(str)
            .str.strip()
            .replace(["nan", "NaN", "None", "NULL", "null", "NaT"], "")
        )

    if "Prestador" not in df.columns or df["Prestador"].fillna("").astype(str).str.strip().eq("").all():
        raise Exception("Coluna Prestador não encontrada na aba Pendências.")

    return df


def gerar_imagem_prestador(prestador, df_prestador):
    df_img = df_prestador[COLUNAS_IMAGEM].copy()

    total_largura_tabela = sum(LARGURAS[coluna] for coluna in COLUNAS_IMAGEM)
    total_altura_tabela = ALTURA_CABECALHO + (len(df_img) * ALTURA_LINHA)

    largura_imagem = total_largura_tabela + (MARGEM * 2)
    altura_imagem = total_altura_tabela + (MARGEM * 2)

    imagem = Image.new("RGB", (largura_imagem, altura_imagem), COR_FUNDO)
    draw = ImageDraw.Draw(imagem)

    x = MARGEM
    y = MARGEM

    for coluna in COLUNAS_IMAGEM:
        largura = LARGURAS[coluna]
        caixa = (x, y, x + largura, y + ALTURA_CABECALHO)

        draw.rectangle(caixa, fill=COR_CABECALHO, outline=COR_BORDA)
        desenhar_texto_centralizado(
            draw,
            caixa,
            coluna,
            FONTE_CABECALHO,
            COR_TEXTO_CABECALHO,
        )

        x += largura

    y += ALTURA_CABECALHO

    for _, linha in df_img.iterrows():
        situacao = linha.get("SITUAÇÃO", "")
        cor_linha = cor_linha_por_situacao(situacao)

        x = MARGEM

        for coluna in COLUNAS_IMAGEM:
            largura = LARGURAS[coluna]
            caixa = (x, y, x + largura, y + ALTURA_LINHA)

            if coluna == "SITUAÇÃO":
                fill, cor_texto = cor_situacao(situacao)
                draw.rectangle(caixa, fill=fill, outline=COR_BORDA)
                desenhar_texto_centralizado(
                    draw,
                    caixa,
                    linha.get(coluna, ""),
                    FONTE_TEXTO_NEGRITO,
                    cor_texto,
                )
            else:
                draw.rectangle(caixa, fill=cor_linha, outline=COR_BORDA)

                if coluna in [
                    "Chamado",
                    "Numero Referencia",
                    "Contratante",
                    "Status",
                    "Data Limite",
                    "CNPJ / CPF",
                    "Cidade",
                ]:
                    desenhar_texto_centralizado(
                        draw,
                        caixa,
                        linha.get(coluna, ""),
                        FONTE_TEXTO,
                        COR_TEXTO,
                    )
                else:
                    desenhar_texto_esquerda(
                        draw,
                        caixa,
                        linha.get(coluna, ""),
                        FONTE_TEXTO,
                        COR_TEXTO,
                    )

            x += largura

        y += ALTURA_LINHA

    nome_arquivo = limpar_nome_arquivo(prestador) + ".png"
    caminho_saida = PASTA_IMAGENS / nome_arquivo

    imagem.save(caminho_saida, "PNG")

    return caminho_saida


def limpar_imagens_antigas():
    PASTA_IMAGENS.mkdir(parents=True, exist_ok=True)

    for arquivo in PASTA_IMAGENS.glob("*.png"):
        try:
            arquivo.unlink()
        except Exception as erro:
            print(f"Aviso: não consegui apagar imagem antiga {arquivo}. Erro: {erro}")


def gerar_imagens():
    print("Gerando imagens por prestador...")

    PASTA_IMAGENS.mkdir(parents=True, exist_ok=True)
    limpar_imagens_antigas()

    df = preparar_dataframe()

    df["Prestador"] = df["Prestador"].fillna("").astype(str).str.strip()

    prestadores = sorted(df["Prestador"].replace("", "SEM_PRESTADOR").unique())

    total = 0

    for prestador in prestadores:
        df_prestador = df[
            df["Prestador"].replace("", "SEM_PRESTADOR") == prestador
        ].copy()

        if df_prestador.empty:
            continue

        caminho = gerar_imagem_prestador(prestador, df_prestador)

        print(f"Imagem gerada: {caminho}")
        total += 1

    print("")
    print(f"Total de imagens geradas: {total}")
    print(f"Pasta: {PASTA_IMAGENS}")


if __name__ == "__main__":
    gerar_imagens()