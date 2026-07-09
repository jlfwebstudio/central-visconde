"""Paleta de cores e tokens de tamanho compartilhados pelas janelas do ViscondeApp.

Antes cada janela (central_mobyan.py, gestao_rotas.py) redefinia sua própria paleta
COR_*, e elas já tinham divergido sutilmente (tons de cinza levemente diferentes,
hovers ausentes numa das duas). Mesmo padrão usado em caminho_base.py pra parar de
duplicar BASE_DIR em cada script: um lugar só de verdade, as duas janelas importam
daqui em vez de redeclarar.
"""

COR_FUNDO = "#080808"
COR_FUNDO_2 = "#101010"
COR_CARD = "#171717"
COR_CARD_HOVER = "#222222"
COR_BORDA = "#584A18"

COR_DOURADO = "#F4C430"
COR_DOURADO_HOVER = "#D8AA18"
COR_DOURADO_ESCURO = "#9B7910"

COR_BRANCO = "#F5F5F5"
COR_TEXTO_SECUNDARIO = "#D5D5D5"
COR_TEXTO_FRACO = "#8E8E8E"

COR_VERDE = "#2EAD68"
COR_VERDE_HOVER = "#248A54"
COR_VERMELHO = "#D14949"
COR_VERMELHO_HOVER = "#AA3939"
COR_AZUL = "#3278C8"
COR_AZUL_HOVER = "#275FA0"
COR_LARANJA = "#D88A24"
COR_LARANJA_HOVER = "#B36F19"
COR_CINZA = "#5B5B5B"

# Tamanho padrão dos cards de ação nas páginas do painel principal — antes cada
# página (e até fileiras dentro da mesma página) usava larguras/alturas diferentes
# sem critério, inclusive uma fileira assimétrica (520px + 260px lado a lado).
LARGURA_CARD_PAR = 390
ALTURA_CARD_PAR = 118
LARGURA_CARD_CHEIO = 790
ALTURA_CARD_CHEIO = 118
