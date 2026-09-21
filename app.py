from datetime import date, datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Painel Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_FUSOS = "lancamentos_fusos_v2.xlsx"

colunas_obrigatorias = [
    "Data_Lancamento",
    "Mes",
    "Ano",
    "Maquina_TAG",
    "Quantidade_Quebras",
    "Observacoes",
]

# Garante o arquivo e a estrutura de colunas corretos
if not os.path.exists(ARQUIVO_FUSOS):
    pd.DataFrame(columns=colunas_obrigatorias).to_excel(
        ARQUIVO_FUSOS, index=False
    )

df_fusos = pd.read_excel(ARQUIVO_FUSOS)

if not all(col in df_fusos.columns for col in colunas_obrigatorias):
    df_fusos = pd.DataFrame(columns=colunas_obrigatorias)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Controle de navegação
if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Fusos"


def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina


# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    st.title("⚙️ Painel Manutenção")
    st.markdown("---")

    st.subheader("📊 Painéis")
    tipo_kpi = (
        "primary"
        if st.session_state.pagina_atual == "Visão Geral (KPIs)"
        else "secondary"
    )
    st.button(
        "📈 Visão Geral (KPIs)",
        use_container_width=True,
        type=tipo_kpi,
        on_click=navegar,
        args=("Visão Geral (KPIs)",),
    )

    tipo_backlog = (
        "primary"
        if st.session_state.pagina_atual == "Controle de Backlog"
        else "secondary"
    )
    st.button(
        "⏳ Controle de Backlog",
        use_container_width=True,
        type=tipo_backlog,
        on_click=navegar,
        args=("Controle de Backlog",),
    )

    tipo_hist = (
        "primary"
        if st.session_state.pagina_atual == "Histórico por TAG"
        else "secondary"
    )
    st.button(
        "🔍 Histórico por TAG",
        use_container_width=True,
        type=tipo_hist,
        on_click=navegar,
        args=("Histórico por TAG",),
    )

    st.markdown("---")

    st.subheader("📝 Lançamentos")
    tipo_fusos = (
        "primary" if st.session_state.pagina_atual == "Fusos" else "secondary"
    )
    st.button(
        "🔩 Fusos",
        use_container_width=True,
        type=tipo_fusos,
        on_click=navegar,
        args=("Fusos",),
    )

    tipo_correias = (
        "primary"
        if st.session_state.pagina_atual == "Correias"
        else "secondary"
    )
    st.button(
        "🔄 Correias",
        use_container_width=True,
        type=tipo_correias,
        on_click=navegar,
        args=("Correias",),
    )

    tipo_prev = (
        "primary"
        if st.session_state.pagina_atual == "Preventiva"
        else "secondary"
    )
    st.button(
        "🛠️ Preventiva",
        use_container_width=True,
        type=tipo_prev,
        on_click=navegar,
        args=("Preventiva",),
    )

    tipo_maq = (
        "primary"
        if st.session_state.pagina_atual == "Máquinas"
        else "secondary"
    )
    st.button(
        "🏭 Máquinas",
        use_container_width=True,
        type=tipo_maq,
        on_click=navegar,
        args=("Máquinas",),
    )

    st.markdown("---")
    st.caption("Perfil: Analista de PCM")

# ==========================================
# ÁREA PRINCIPAL (TELA DA DIREITA)
# ==========================================
tela = st.session_state.pagina_atual

if tela == "Fusos":
    st.header("🔩 Controle de Quebras de Fusos")
    st.write(
        "Filtre por período ou visualize o consolidado geral de quebras de fusos."
    )

    # 1. SELEÇÃO COM OPÇÃO 'TODOS'
    c_ano, c_mes = st.columns([1, 2])

    lista_anos = ["Todos", "2024", "2025", "2026", "2027", "2028"]
    lista_meses_puros = [
        "Janeiro",
        "Fevereiro",
        "Março",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto",
        "Setembro",
        "Outubro",
        "Novembro",
        "Dezembro",
    ]
    lista_meses = ["Todos"] + lista_meses_puros

    with c_ano:
        ano_selecionado = st.selectbox(
            "📅 Selecione o Ano:", lista_anos, index=3
        )  # Padrão: 2026

    with c_mes:
        mes_selecionado = st.selectbox(
            "🗓️ Selecione o Mês:", lista_meses, index=0
        )  # Padrão: Todos

    st.markdown("---")

    # Lê a base gravada
    df_fusos = pd.read_excel(ARQUIVO_FUSOS)

    # Rótulo de exibição amigável
    periodo_texto = f"{mes_selecionado} / {ano_selecionado}"
    if mes_selecionado == "Todos" and ano_selecionado == "Todos":
        periodo_texto = "Histórico Completo (Todos os Períodos)"
    elif mes_selecionado == "Todos":
        periodo_texto = f"Todos os Meses de {ano_selecionado}"
    elif ano_selecionado == "Todos":
        periodo_texto = f"Mês de {mes_selecionado} (Todos os Anos)"

    tab_dados, tab_novo = st.tabs(
        [
            f"📋 Consulta & Indicadores ({periodo_texto})",
            "➕ Registrar Nova Quebra",
        ]
    )

    with tab_dados:
        st.subheader(f"Ocorrências: {periodo_texto}")

        # Aplica os filtros dinâmicos
        df_filtrado = df_fusos.copy()

        if ano_selecionado != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Ano"] == int(ano_selecionado)
            ]

        if mes_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Mes"] == mes_selecionado]

        if not df_filtrado.empty:
            total_quebras = df_filtrado["Quantidade_Quebras"].sum()
            maquina_campea = (
                df_filtrado.groupby("Maquina_TAG")["Quantidade_Quebras"]
                .sum()
                .idxmax()
            )
            qtd_campea = df_filtrado.groupby("Maquina_TAG")[
                "Quantidade_Quebras"
            ].sum().max()

            k1, k2 = st.columns(2)
            k1.metric("Total de Quebras no Período", f"{total_quebras} fusos")
            k2.metric(
                "Máquina Crítica (Mais Quebras)",
                f"{maquina_campea} ({qtd_campea} quebras)",
            )

            st.write("### Registros Detalhados:")
            st.dataframe(
                df_filtrado[
                    [
                        "Data_Lancamento",
                        "Mes",
                        "Ano",
                        "Maquina_TAG",
                        "Quantidade_Quebras",
                        "Observacoes",
                    ]
                ],
                use_container_width=True,
            )
        else:
            st.info(f"Nenhum registro encontrado para {periodo_texto}.")

    with tab_novo:
        st.subheader("Apontamento de Nova Quebra de Fuso")

        with st.form("form_quebras_fusos", clear_on_submit=True):
            col_data, col_ano_form, col_mes_form = st.columns(3)

            with col_data:
                data_lancto = st.date_input("Data do Ocorrido", value=date.today())

            with col_ano_form:
                ano_padrao = (
                    int(ano_selecionado) if ano_selecionado != "Todos" else 2026
                )
                ano_salvar = st.selectbox(
                    "Ano de Referência",
                    [2024, 2025, 2026, 2027, 2028],
                    index=[2024, 2025, 2026, 2027, 2028].index(ano_padrao),
                )

            with col_mes_form:
                mes_padrao = (
                    mes_selecionado
                    if mes_selecionado != "Todos"
                    else "Setembro"
                )
                mes_salvar = st.selectbox(
                    "Mês de Referência",
                    lista_meses_puros,
                    index=lista_meses_puros.index(mes_padrao),
                )

            col_maq, col_qtd = st.columns([2, 1])

            with col_maq:
                maquina_tag = st.text_input(
                    "Máquina / TAG",
                    placeholder="Ex: TORNO-CNC-01, CENTRO-02",
                )

            with col_qtd:
                qtd_quebras = st.number_input(
                    "Qtd de Quebras", min_value=1, max_value=50, step=1
                )

            obs = st.text_input(
                "Observações / Causa (Opcional)",
                placeholder="Ex: Colisão mecânica / rolamento travado",
            )

            btn_gravar = st.form_submit_button("Salvar Lançamento")

            if btn_gravar:
                if maquina_tag.strip() != "":
                    novo_dado = {
                        "Data_Lancamento": data_lancto.strftime("%d/%m/%Y"),
                        "Mes": mes_salvar,
                        "Ano": int(ano_salvar),
                        "Maquina_TAG": maquina_tag.strip().upper(),
                        "Quantidade_Quebras": int(qtd_quebras),
                        "Observacoes": obs.strip(),
                    }

                    df_fusos = pd.concat(
                        [df_fusos, pd.DataFrame([novo_dado])], ignore_index=True
                    )
                    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
                    st.success(
                        f"✅ Gravado: {qtd_quebras} quebra(s) para {maquina_tag.strip().upper()} ({mes_salvar}/{ano_salvar})!"
                    )
                    st.rerun()
                else:
                    st.error("Por favor, preencha o campo Máquina / TAG.")

# Outras telas mantidas
elif tela == "Correias":
    st.header("🔄 Lançamentos: Correias")
elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Máquinas")
elif tela == "Visão Geral (KPIs)":
    st.header("📊 Painel Geral de Indicadores")
elif tela == "Controle de Backlog":
    st.header("⏳ Controle de Backlog de Ordens")
elif tela == "Histórico por TAG":
    st.header("🔍 Histórico e Falhas por TAG")
