from datetime import date
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
    st.session_state.pagina_atual = "Painel Fusos"


def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina


# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    st.title("⚙️ Painel Manutenção")
    st.markdown("---")

    # SEÇÃO 1: PAINÉIS (APENAS FUSOS)
    st.subheader("📊 Painéis")
    tipo_painel_fusos = (
        "primary"
        if st.session_state.pagina_atual == "Painel Fusos"
        else "secondary"
    )
    st.button(
        "🔩 Fusos",
        use_container_width=True,
        type=tipo_painel_fusos,
        on_click=navegar,
        args=("Painel Fusos",),
    )

    st.markdown("---")

    # SEÇÃO 2: LANÇAMENTOS
    st.subheader("📝 Lançamentos")

    tipo_fusos = (
        "primary"
        if st.session_state.pagina_atual == "Lançamento Fusos"
        else "secondary"
    )
    st.button(
        "🔩 Fusos",
        use_container_width=True,
        type=tipo_fusos,
        on_click=navegar,
        args=("Lançamento Fusos",),
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

# ------------------------------------------
# 1. PAINEL GERENCIAL DE FUSOS
# ------------------------------------------
if tela == "Painel Fusos":
    st.header("📊 Painel Gerencial: Quebras de Fusos")
    st.write("Análise visual e indicadores de falhas em fusos da fábrica.")

    # Filtros de Período
    col_ano, col_mes = st.columns([1, 2])
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

    with col_ano:
        ano_painel = st.selectbox(
            "📅 Filtrar por Ano:", lista_anos, index=3, key="p_ano"
        )
    with col_mes:
        mes_painel = st.selectbox(
            "🗓️ Filtrar por Mês:", lista_meses, index=0, key="p_mes"
        )

    st.markdown("---")

    df_dados = pd.read_excel(ARQUIVO_FUSOS)
    df_filtrado = df_dados.copy()

    if ano_painel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Ano"] == int(ano_painel)]
    if mes_painel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Mes"] == mes_painel]

    if not df_filtrado.empty:
        total_quebras = df_filtrado["Quantidade_Quebras"].sum()
        total_maquinas_afetadas = df_filtrado["Maquina_TAG"].nunique()

        agrupado_maq = (
            df_filtrado.groupby("Maquina_TAG")["Quantidade_Quebras"]
            .sum()
            .reset_index()
        )
        agrupado_maq = agrupado_maq.sort_values(
            by="Quantidade_Quebras", ascending=False
        )

        maquina_top = agrupado_maq.iloc[0]["Maquina_TAG"]
        qtd_top = agrupado_maq.iloc[0]["Quantidade_Quebras"]

        # Cartões de Métricas
        m1, m2, m3 = st.columns(3)
        m1.metric("Total de Fusos Quebrados", f"{total_quebras} unid.")
        m2.metric("Máquinas com Falhas", f"{total_maquinas_afetadas} ativas")
        m3.metric("Maior Ofensor", f"{maquina_top} ({qtd_top} quebras)")

        st.markdown("---")

        # Gráfico e Tabela lado a lado
        graf_col, tab_col = st.columns([2, 1])

        with graf_col:
            st.subheader("Ranking de Quebras por Máquina")
            # Gráfico nativo do Streamlit
            st.bar_chart(
                data=agrupado_maq.set_index("Maquina_TAG")[
                    "Quantidade_Quebras"
                ]
            )

        with tab_col:
            st.subheader("Consolidado por Ativo")
            st.dataframe(
                agrupado_maq.rename(
                    columns={
                        "Maquina_TAG": "Máquina",
                        "Quantidade_Quebras": "Total Quebras",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    else:
        st.info("Nenhum dado registrado para o período selecionado.")

# ------------------------------------------
# 2. ENTRADA / LANÇAMENTO DE FUSOS
# ------------------------------------------
elif tela == "Lançamento Fusos":
    st.header("🔩 Lançamento: Apontamento de Quebras de Fusos")
    st.write(
        "Insira as novas quebras identificadas nas máquinas para alimentar os painéis."
    )

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

    with st.form("form_novo_fuso", clear_on_submit=True):
        col_data, col_ano_l, col_mes_l = st.columns(3)

        with col_data:
            dt_reg = st.date_input("Data do Ocorrido", value=date.today())
        with col_ano_l:
            ano_reg = st.selectbox(
                "Ano de Referência", [2024, 2025, 2026, 2027, 2028], index=2
            )
        with col_mes_l:
            mes_reg = st.selectbox(
                "Mês de Referência", lista_meses_puros, index=8
            )

        c_tag, c_qtd = st.columns([2, 1])
        with c_tag:
            tag_input = st.text_input(
                "Máquina / TAG", placeholder="Ex: TORNO-01, CENTRO-USINAGEM-02"
            )
        with c_qtd:
            qtd_input = st.number_input(
                "Qtd de Quebras", min_value=1, max_value=50, step=1
            )

        obs_input = st.text_input(
            "Observações / Causa (Opcional)",
            placeholder="Ex: Fadiga prematura, colisão",
        )

        salvar = st.form_submit_button("Salvar Registro")

        if salvar:
            if tag_input.strip() != "":
                novo = {
                    "Data_Lancamento": dt_reg.strftime("%d/%m/%Y"),
                    "Mes": mes_reg,
                    "Ano": int(ano_reg),
                    "Maquina_TAG": tag_input.strip().upper(),
                    "Quantidade_Quebras": int(qtd_input),
                    "Observacoes": obs_input.strip(),
                }
                df_atual = pd.read_excel(ARQUIVO_FUSOS)
                df_atual = pd.concat(
                    [df_atual, pd.DataFrame([novo])], ignore_index=True
                )
                df_atual.to_excel(ARQUIVO_FUSOS, index=False)
                st.success(
                    f"✅ Registrado com sucesso: {qtd_input} quebra(s) em {tag_input.strip().upper()} ({mes_reg}/{ano_reg})!"
                )
            else:
                st.error("Preencha o campo Máquina / TAG.")

# Outras telas de lançamentos
elif tela == "Correias":
    st.header("🔄 Lançamentos: Correias")
elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Máquinas")
