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

# Arquivo Excel onde os dados de fusos ficam gravados
ARQUIVO_FUSOS = "lancamentos_fusos_v2.xlsx"

colunas_obrigatorias = [
    "Data_Lancamento",
    "Mes",
    "Ano",
    "Maquina_TAG",
    "Quantidade_Quebras",
    "Observacoes",
]

# Cria ou garante as colunas corretas na planilha
if not os.path.exists(ARQUIVO_FUSOS):
    pd.DataFrame(columns=colunas_obrigatorias).to_excel(
        ARQUIVO_FUSOS, index=False
    )

df_fusos = pd.read_excel(ARQUIVO_FUSOS)

# Se faltar qualquer coluna no arquivo existente, recria limpo
if not all(col in df_fusos.columns for col in colunas_obrigatorias):
    df_fusos = pd.DataFrame(columns=colunas_obrigatorias)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Controle de navegação da página
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
    st.write("Selecione o período (Mês/Ano) para lançamentos e acompanhamento.")

    c_ano, c_mes = st.columns([1, 2])

    lista_anos = [2024, 2025, 2026, 2027, 2028]
    lista_meses = [
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

    with c_ano:
        ano_selecionado = st.selectbox(
            "📅 Selecione o Ano:", lista_anos, index=2
        )  # 2026

    with c_mes:
        mes_selecionado = st.selectbox(
            "🗓️ Selecione o Mês:", lista_meses, index=8  # Setembro
        )

    st.markdown("---")

    # Lê os dados da planilha atualizada
    df_fusos = pd.read_excel(ARQUIVO_FUSOS)

    tab_novo, tab_dados = st.tabs(
        [
            f"➕ Registrar Quebras ({mes_selecionado}/{ano_selecionado})",
            f"📋 Consulta & Indicadores ({mes_selecionado}/{ano_selecionado})",
        ]
    )

    with tab_novo:
        st.subheader(
            f"Apontamento de Quebras - {mes_selecionado}/{ano_selecionado}"
        )

        with st.form("form_quebras_fusos", clear_on_submit=True):
            col_maq, col_qtd = st.columns([2, 1])

            with col_maq:
                maquina_tag = st.text_input(
                    "Máquina / TAG",
                    placeholder="Ex: TORNO-CNC-01, CENTRO-USINAGEM-03",
                )

            with col_qtd:
                qtd_quebras = st.number_input(
                    "Qtd de Quebras de Fuso", min_value=1, max_value=50, step=1
                )

            obs = st.text_input(
                "Observações / Causa (Opcional)",
                placeholder="Ex: Falha por colisão ou desgaste prematuro",
            )

            btn_gravar = st.form_submit_button("Salvar Lançamento")

            if btn_gravar:
                if maquina_tag.strip() != "":
                    novo_dado = {
                        "Data_Lancamento": date.today().strftime("%d/%m/%Y"),
                        "Mes": mes_selecionado,
                        "Ano": int(ano_selecionado),
                        "Maquina_TAG": maquina_tag.strip().upper(),
                        "Quantidade_Quebras": int(qtd_quebras),
                        "Observacoes": obs.strip(),
                    }

                    df_fusos = pd.concat(
                        [df_fusos, pd.DataFrame([novo_dado])], ignore_index=True
                    )
                    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
                    st.success(
                        f"✅ Gravado: {qtd_quebras} quebra(s) na máquina {maquina_tag.strip().upper()}!"
                    )
                    st.rerun()
                else:
                    st.error("Por favor, preencha o campo Máquina / TAG.")

    with tab_dados:
        st.subheader(f"Ocorrências em {mes_selecionado}/{ano_selecionado}")

        # Filtra pelo mês e ano escolhidos
        df_periodo = df_fusos[
            (df_fusos["Mes"] == mes_selecionado)
            & (df_fusos["Ano"] == ano_selecionado)
        ]

        if not df_periodo.empty:
            total_quebras = df_periodo["Quantidade_Quebras"].sum()
            maquina_campea = (
                df_periodo.groupby("Maquina_TAG")["Quantidade_Quebras"]
                .sum()
                .idxmax()
            )

            kpi1, kpi2 = st.columns(2)
            kpi1.metric("Total de Quebras no Período", f"{total_quebras} fusos")
            kpi2.metric("Máquina com Mais Quebras", maquina_campea)

            st.write("### Registros Detalhados:")
            st.dataframe(
                df_periodo[
                    [
                        "Data_Lancamento",
                        "Maquina_TAG",
                        "Quantidade_Quebras",
                        "Observacoes",
                    ]
                ],
                use_container_width=True,
            )
        else:
            st.info(
                f"Nenhuma quebra registrada para {mes_selecionado}/{ano_selecionado}."
            )

# Outras seções
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
