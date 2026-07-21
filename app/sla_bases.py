"""SLA por base (prestador): acumula um histórico diário das pendências (Mobyan +
OGEA) por Origem+Prestador e calcula tendência/nível de cobrança a partir dele.

Não calcula o SLA% contratual real (isso exigiria um relatório de OS concluída com
data de finalização, que não existe no scraping atual — ver plano de implementação).
Ataca o que dá pra medir com o dado que já passa pelo pipeline: risco de atraso e
recorrência por base, pra embasar a cobrança do dia a dia.
"""

from datetime import timedelta

import pandas as pd

from caminho_base import BASE_DIR

PASTA_PENDENCIAS = BASE_DIR / "outputs" / "pendencias_do_dia"
ARQUIVO_HISTORICO_SLA = PASTA_PENDENCIAS / "historico_sla.xlsx"

COLUNAS_HISTORICO = [
    "Data",
    "Origem",
    "Prestador",
    "Total Pendentes",
    "Vencidas",
    "Vence Hoje",
    "Futuras",
    "Falta Abonar",
    "Atraso Médio (dias)",
    "Maior Atraso (dias)",
]

COLUNAS_RANKING = COLUNAS_HISTORICO + ["Média Vencidas (7d)", "Tendência", "Nível de Cobrança"]

COLUNAS_DETALHE_TECNICO = [
    "Origem",
    "Prestador",
    "Técnico",
    "Cidade",
    "Total Pendentes",
    "Vencidas",
    "Vence Hoje",
    "Maior Atraso (dias)",
]

# Janela e limiares de tendência/nível de cobrança — ajustáveis depois de ver dado
# real (não há nenhuma base histórica ainda pra calibrar com precisão).
JANELA_TENDENCIA_DIAS = 7
LIMIAR_SUBINDO = 1.2
LIMIAR_CAINDO = 0.8

LIMIAR_CRITICO_VENCIDAS = 5
LIMIAR_CRITICO_VENCIDAS_SUBINDO = 3
LIMIAR_ATENCAO_VENCIDAS = 1


def _dias_atraso(data_limite_str, hoje):
    data = pd.to_datetime(data_limite_str, format="%d/%m/%Y", errors="coerce")
    if pd.isna(data):
        return 0
    dias = (hoje - data.date()).days
    return dias if dias > 0 else 0


def calcular_snapshot_diario(df_chamados_unificado, hoje):
    """Agrega o dataframe unificado de Chamados (Mobyan+OGEA, coluna Origem já
    presente) por Origem+Prestador — um retrato do dia por base."""
    if df_chamados_unificado is None or df_chamados_unificado.empty:
        return pd.DataFrame(columns=COLUNAS_HISTORICO)

    df = df_chamados_unificado.copy()
    df["Prestador"] = df["Prestador"].fillna("").astype(str).str.strip()
    df = df[df["Prestador"] != ""]

    if df.empty:
        return pd.DataFrame(columns=COLUNAS_HISTORICO)

    df["_dias_atraso"] = df.apply(
        lambda linha: _dias_atraso(linha.get("Data Limite", ""), hoje)
        if linha.get("SITUAÇÃO") == "VENCIDA"
        else 0,
        axis=1,
    )

    linhas = []

    for (origem, prestador), grupo in df.groupby(["Origem", "Prestador"]):
        vencidas_mask = grupo["SITUAÇÃO"] == "VENCIDA"
        atrasos = grupo.loc[vencidas_mask, "_dias_atraso"]

        linhas.append({
            "Data": hoje.strftime("%d/%m/%Y"),
            "Origem": origem,
            "Prestador": prestador,
            "Total Pendentes": len(grupo),
            "Vencidas": int(vencidas_mask.sum()),
            "Vence Hoje": int((grupo["SITUAÇÃO"] == "VENCE HOJE").sum()),
            "Futuras": int((grupo["SITUAÇÃO"] == "FUTURA").sum()),
            "Falta Abonar": int((grupo["Justificativa do Abono"] == "FALTA ABONAR").sum()),
            "Atraso Médio (dias)": round(atrasos.mean(), 1) if len(atrasos) else 0,
            "Maior Atraso (dias)": int(atrasos.max()) if len(atrasos) else 0,
        })

    return pd.DataFrame(linhas, columns=COLUNAS_HISTORICO)


def carregar_historico():
    if not ARQUIVO_HISTORICO_SLA.exists():
        return pd.DataFrame(columns=COLUNAS_HISTORICO)

    try:
        df = pd.read_excel(ARQUIVO_HISTORICO_SLA, sheet_name="Histórico")
    except Exception:
        return pd.DataFrame(columns=COLUNAS_HISTORICO)

    for coluna in COLUNAS_HISTORICO:
        if coluna not in df.columns:
            df[coluna] = ""

    return df[COLUNAS_HISTORICO]


def registrar_historico(snapshot_df, hoje):
    """Acrescenta o snapshot de hoje ao histórico acumulado (arquivo único,
    append-only), substituindo qualquer linha já existente do mesmo dia+base —
    idempotente a reprocessamentos no mesmo dia."""
    PASTA_PENDENCIAS.mkdir(parents=True, exist_ok=True)

    historico = carregar_historico()

    if snapshot_df is None or snapshot_df.empty:
        return historico

    data_hoje_str = hoje.strftime("%d/%m/%Y")
    chaves_hoje = set(zip(
        [data_hoje_str] * len(snapshot_df),
        snapshot_df["Origem"],
        snapshot_df["Prestador"],
    ))

    if historico.empty:
        historico_atualizado = snapshot_df.copy()
    else:
        chaves_historico = list(zip(historico["Data"], historico["Origem"], historico["Prestador"]))
        historico = historico[[chave not in chaves_hoje for chave in chaves_historico]]
        historico_atualizado = pd.concat([historico, snapshot_df], ignore_index=True)
    historico_atualizado.to_excel(ARQUIVO_HISTORICO_SLA, index=False, sheet_name="Histórico")

    return historico_atualizado


def calcular_ranking_sla(historico_df, snapshot_hoje_df, hoje):
    """Cruza o snapshot de hoje com a média dos últimos JANELA_TENDENCIA_DIAS dias
    (excluindo hoje) por Origem+Prestador, pra apontar tendência e nível de cobrança
    de cada base."""
    if snapshot_hoje_df is None or snapshot_hoje_df.empty:
        return pd.DataFrame(columns=COLUNAS_RANKING)

    limite_janela = hoje - timedelta(days=JANELA_TENDENCIA_DIAS)
    data_hoje_str = hoje.strftime("%d/%m/%Y")

    historico_recente = historico_df.copy() if historico_df is not None else pd.DataFrame(columns=COLUNAS_HISTORICO)

    if not historico_recente.empty:
        datas = pd.to_datetime(historico_recente["Data"], format="%d/%m/%Y", errors="coerce")
        historico_recente = historico_recente[
            (datas.dt.date >= limite_janela) & (historico_recente["Data"] != data_hoje_str)
        ]

    linhas = []

    for _, linha in snapshot_hoje_df.iterrows():
        origem = linha["Origem"]
        prestador = linha["Prestador"]
        vencidas_hoje = int(linha["Vencidas"])

        if not historico_recente.empty:
            historico_base = historico_recente[
                (historico_recente["Origem"] == origem) & (historico_recente["Prestador"] == prestador)
            ]
        else:
            historico_base = historico_recente

        if len(historico_base):
            media_vencidas_7d = round(historico_base["Vencidas"].astype(float).mean(), 1)
        else:
            media_vencidas_7d = 0.0

        if media_vencidas_7d == 0:
            tendencia = "SUBINDO" if vencidas_hoje > 0 else "ESTÁVEL"
        elif vencidas_hoje > media_vencidas_7d * LIMIAR_SUBINDO:
            tendencia = "SUBINDO"
        elif vencidas_hoje < media_vencidas_7d * LIMIAR_CAINDO:
            tendencia = "CAINDO"
        else:
            tendencia = "ESTÁVEL"

        if vencidas_hoje >= LIMIAR_CRITICO_VENCIDAS or (
            vencidas_hoje >= LIMIAR_CRITICO_VENCIDAS_SUBINDO and tendencia == "SUBINDO"
        ):
            nivel = "CRÍTICO"
        elif vencidas_hoje >= LIMIAR_ATENCAO_VENCIDAS:
            nivel = "ATENÇÃO"
        else:
            nivel = "NORMAL"

        nova_linha = linha.to_dict()
        nova_linha["Média Vencidas (7d)"] = media_vencidas_7d
        nova_linha["Tendência"] = tendencia
        nova_linha["Nível de Cobrança"] = nivel
        linhas.append(nova_linha)

    df_ranking = pd.DataFrame(linhas, columns=COLUNAS_RANKING)

    ordem_nivel = {"CRÍTICO": 0, "ATENÇÃO": 1, "NORMAL": 2}
    df_ranking["_ordem"] = df_ranking["Nível de Cobrança"].map(ordem_nivel)
    df_ranking = (
        df_ranking.sort_values(by=["_ordem", "Vencidas"], ascending=[True, False])
        .drop(columns="_ordem")
        .reset_index(drop=True)
    )

    return df_ranking


def calcular_detalhe_por_tecnico(df_chamados_unificado, hoje):
    """Agrega por Prestador+Técnico — só informativo, pra apontar qual técnico/
    cidade está puxando o atraso de uma base pra baixo (ex.: rota visitada só 1x
    por semana). Não depende de nenhum cadastro de vínculo Técnico->Prestador:
    as duas colunas já convivem na mesma linha do Chamado."""
    if df_chamados_unificado is None or df_chamados_unificado.empty:
        return pd.DataFrame(columns=COLUNAS_DETALHE_TECNICO)

    df = df_chamados_unificado.copy()
    df["Prestador"] = df["Prestador"].fillna("").astype(str).str.strip()
    df["Técnico"] = df["Técnico"].fillna("").astype(str).str.strip()
    df = df[(df["Prestador"] != "") & (df["Técnico"] != "")]

    if df.empty:
        return pd.DataFrame(columns=COLUNAS_DETALHE_TECNICO)

    df["_dias_atraso"] = df.apply(
        lambda linha: _dias_atraso(linha.get("Data Limite", ""), hoje)
        if linha.get("SITUAÇÃO") == "VENCIDA"
        else 0,
        axis=1,
    )

    linhas = []

    for (origem, prestador, tecnico), grupo in df.groupby(["Origem", "Prestador", "Técnico"]):
        vencidas_mask = grupo["SITUAÇÃO"] == "VENCIDA"
        atrasos = grupo.loc[vencidas_mask, "_dias_atraso"]
        cidades = sorted({c for c in grupo["Cidade"].fillna("").astype(str).str.strip() if c})

        linhas.append({
            "Origem": origem,
            "Prestador": prestador,
            "Técnico": tecnico,
            "Cidade": ", ".join(cidades),
            "Total Pendentes": len(grupo),
            "Vencidas": int(vencidas_mask.sum()),
            "Vence Hoje": int((grupo["SITUAÇÃO"] == "VENCE HOJE").sum()),
            "Maior Atraso (dias)": int(atrasos.max()) if len(atrasos) else 0,
        })

    df_resultado = pd.DataFrame(linhas, columns=COLUNAS_DETALHE_TECNICO)

    return df_resultado.sort_values(by="Vencidas", ascending=False).reset_index(drop=True)
