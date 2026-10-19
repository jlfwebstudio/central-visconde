"""Progresso amigável das automações, mostrado na barra da Central Visconde.

Cada script chama emitir(percentual, mensagem) nos pontos-chave da execução.
central_mobyan.py reconhece o prefixo e atualiza a barra + status em vez de
jogar a linha no log cru — ver PREFIXO abaixo, usado também lá.
"""

PREFIXO = "@PROGRESSO:"


def emitir(percentual, mensagem):
    print(f"{PREFIXO}{percentual}:{mensagem}", flush=True)
