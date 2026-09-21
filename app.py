from datetime import date
import os
import pandas as pd
import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="Painel Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Arquivo Excel onde os fusos ficam salvos
ARQUIVO_FUSOS = "lancamentos_fusos.xlsx"

# Se o arquivo não existir, cria a planilha com as colunas base
if not os.path.exists(ARQUIVO_FUSOS):
    colunas = [
        "Data",
        "Mês",
        "Máquina/TAG",
        "Posição/Eixo",
        "Tipo Ação",
        "Técnico",
        "Observações",
    ]
    pd.DataFrame(columns=colunas).to_excel(ARQUIVO_FUSOS, index=False)

# 1. Guarda na memória qual tela está aberta
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

    # SEÇÃO 1: PAINÉIS
    st.subheader("📊 Painéis")

    tipo_btn1 = (
        "primary"
        if st.session_state.pagina_atual == "Visão Geral (KPIs)"
        else "secondary"
    )
    st.button(
        "📈 Visão Geral (KPIs)",
        use_container_width=True,
        type=tipo_btn1,
        on_click=navegar,
        args=("Visão Geral (KPIs)",),
    )

    tipo_btn2 = (
        "primary"
        if st.session_state.pagina_atual == "Controle de Backlog"
        else "secondary"
    )
    st.button(
        "⏳ Controle de Backlog",
        use_container_width=True,
        type=tipo_btn2,
        on_click=navegar,
        args=("Controle de Backlog",),
    )

    tipo_btn3 = (
        "primary"
        if st.session_state.pagina_atual == "Histórico por TAG"
        else "secondary"
    )
    st.button(
        "🔍 Histórico por TAG",
        use_container_width=True,
        type=tipo_btn3,
        on_click=navegar,
        args=("Histórico por TAG",),
    )

    st.markdown("---")

    # SEÇÃO 2: LANÇAMENTOS
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

    tipo_preventiva = (
        "primary"
        if st.session_state.pagina_atual == "Preventiva"
        else "secondary"
    )
    st.button(
        "🛠️ Preventiva",
        use_container_width=True,
        type=tipo_preventiva,
        on_click=navegar,
        args=("Preventiva",),
    )

    tipo_maquinas = (
        "primary"
        if st.session_state.pagina_atual == "Máquinas"
        else "secondary"
    )
    st.button(
        "🏭 Máquinas",
        use_container_width=True,
        type=tipo_maquinas,
        on_click=navegar,
        args=("Máquinas",),
    )

    st.markdown("---")
    st.caption("Perfil: Analista de PCM")

# ==========================================
# ÁREA DA DIREITA (CONTEÚDO DINÂMICO)
# ==========================================
tela = st.session_state.pagina_atual

if tela == "Fusos":
    st.header("🔩 Controle e Lançamento de Fusos")
    st.write(
        "Selecione o mês para visualizar o histórico ou registrar uma nova intervenção."
    )

    # Lista dos meses
    meses_ano = [
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

    # Seletor de mês em destaque no topo
    mes_selecionado = st.selectbox(
        "📅 Filtrar dados por Mês:",
        meses_ano,
        index=8,  # Começa selecionado em Setembro
    )

    st.markdown("---")

    # Lê os dados atuais da planilha
    df_fusos = pd.read_excel(ARQUIVO_FUSOS)

    # Criação de duas abas visuais na tela: uma para Lançar e outra para Ver Histórico
    tab_lancamento, tab_historico = st.tabs(
        [
            f"➕ Novo Registro ({mes_selecionado})",
            f"📋 Histórico ({mes_selecionado})",
        ]
    )

    with tab_lancamento:
        st.subheader(f"Registrar Troca / Ação em {mes_selecionado}")

        with st.form(key="form_fuso", clear_on_submit=True):
            col1, col2 = st.columns(2)

            with col1:
                data_evento = st.date_input("Data da Ação", value=date.today())
                tag = st.text_input(
                    "Máquina / TAG", placeholder="Ex: TORNO-CNC-01"
                )
                posicao = st.text_input(
                    "Posição / Eixo do Fuso", placeholder="Ex: Eixo X, Eixo Z"
                )

            with col2:
                tipo_acao = st.selectbox(
                    "Tipo de Ação",
                    [
                        "Troca Completa",
                        "Substituição de Castanha",
                        "Ajuste / Alinhamento",
                        "Lubrificação / Limpeza",
                        "Inspeção de Folga",
                    ],
                )
                tecnico = st.text_input("Técnico / Responsável")
                obs = st.text_area(
                    "Observações / Motivo da Troca",
                    placeholder="Ex: Fuso com folga excessiva acima de 0.05mm",
                )

            btn_salvar = st.form_submit_button("Salvar Registro")

            if btn_salvar:
                if tag.strip() != "":
                    novo_registro = {
                        "Data": data_evento.strftime("%d/%m/%Y"),
                        "Mês": mes_selecionado,
                        "Máquina/TAG": tag.strip().upper(),
                        "Posição/Eixo": posicao.strip(),
                        "Tipo Ação": tipo_acao,
                        "Técnico": tecnico.strip(),
                        "Observações": obs.strip(),
                    }

                    df_fusos = pd.concat(
                        [df_fusos, pd.DataFrame([novo_registro])],
                        ignore_index=True,
                    )
                    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
                    st.success(
                        f"✅ Registro salvo com sucesso no mês de {mes_selecionado}!"
                    )
                    st.rerun()
                else:
                    st.error("Por favor, preencha pelo menos o campo TAG.")

    with tab_historico:
        st.subheader(f"Registros de Fusos em {mes_selecionado}")

        # Filtra o DataFrame apenas para o mês selecionado
        df_mes = df_fusos[df_fusos["Mês"] == mes_selecionado]

        if not df_mes.empty:
            st.dataframe(df_mes, use_container_width=True)
            st.metric("Total de Ações no Mês", len(df_mes))
        else:
            st.info(
                f"Nenhum registro encontrado para o mês de {mes_selecionado}."
            )

# --- DEMAIS TELAS (MANTIDAS) ---
elif tela == "Correias":
    st.header("🔄 Lançamentos: Controle de Correias")
    st.write("Registro de tensionamento, trocas e inspeção de correias.")

elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Planos de Manutenção Preventiva")
    st.write("Checklists e fechamentos de rotinas preventivas periódicas.")

elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Cadastro e Dados de Máquinas")
    st.write("Cadastro de TAGs, capacidades, setores e criticidades.")

elif tela == "Visão Geral (KPIs)":
    st.header("📊 Painel Geral de Indicadores")

elif tela == "Controle de Backlog":
    st.header("⏳ Controle de Backlog de Ordens")

elif tela == "Histórico por TAG":
    st.header("🔍 Histórico e Falhas por TAG")
