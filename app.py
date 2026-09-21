import os
from datetime import date
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Painel Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_FUSOS = "lancamentos_fusos_v4.xlsx"

colunas_obrigatorias = [
    "Ano",
    "Mes",
    "Setor",
    "Maquina_TAG",
    "Quantidade_Quebras",
    "Observacoes",
]

# Inicializa o ficheiro se ainda não existir
if not os.path.exists(ARQUIVO_FUSOS):
    pd.DataFrame(columns=colunas_obrigatorias).to_excel(
        ARQUIVO_FUSOS, index=False
    )

df_fusos = pd.read_excel(ARQUIVO_FUSOS)
if not all(col in df_fusos.columns for col in colunas_obrigatorias):
    df_fusos = pd.DataFrame(columns=colunas_obrigatorias)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Mapeamento oficial dos setores e máquinas
maquinas_setor_a = [f"L-{i:02d}" for i in range(1, 29)]

maquinas_setor_b = [
    "L-29",
    "L-30",
    "L-31",
    "L-32",
    "L-33",
    "L-34",
    "L-35",
    "L-36",
    "L-37",
    "L-38",
    "L-39",
    "L-40",
    "L-41",
    "L-42",
    "L-43",
    "L-44",
    "L-45",
    "L-46",
    "L-50",
    "L-51",
    "L-52",
    "L-53",
]

maquinas_setor_latex = [
    "B-71",
    "B-72",
    "B-73",
    "B-74",
    "B-75",
    "B-76",
    "B-77",
    "B-78",
    "B-79",
    "B-80",
    "B-83",
    "B-84",
    "B-85",
    "B-86",
    "B-87",
    "B-88",
    "B-89",
    "B-102",
    "B-103",
    "B-104",
]

maquinas_setor_menegatto = [
    "B-47",
    "B-48",
    "B-49",
    "B-81",
    "B-82",
    "B-90",
    "B-91",
    "B-92",
    "B-93",
    "B-94",
    "B-95",
    "B-96",
    "B-97",
    "B-98",
    "B-99",
    "B-100",
    "B-101",
]

DICIONARIO_SETORES = {
    "Setor A": maquinas_setor_a,
    "Setor B": maquinas_setor_b,
    "Setor Látex": maquinas_setor_latex,
    "Setor Menegatto": maquinas_setor_menegatto,
}

# Gestão de navegação da aplicação
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

    # PAINÉIS
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

    # LANÇAMENTOS
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
# ÁREA PRINCIPAL
# ==========================================
tela = st.session_state.pagina_atual

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

# ------------------------------------------
# 1. PAINEL GERENCIAL DE FUSOS
# ------------------------------------------
if tela == "Painel Fusos":
    st.header("📊 Painel Gerencial: Tendência e Quebras de Fusos")
    st.write(
        "Acompanhe a curva de evolução de Janeiro a Dezembro para o ano selecionado."
    )

    # Filtros de Ano, Setor e Máquina (Ano com foco anual exato)
    col_ano, col_setor, col_maq_f = st.columns([1, 1.5, 1.5])
    lista_anos_painel = [2024, 2025, 2026, 2027, 2028]
    lista_setores_painel = ["Todos"] + list(DICIONARIO_SETORES.keys())

    with col_ano:
        ano_painel = st.selectbox(
            "📅 Ano:", lista_anos_painel, index=2, key="p_ano"
        )
    with col_setor:
        setor_painel = st.selectbox(
            "🏭 Setor:",
            lista_setores_painel,
            index=0,
            key="p_setor",
        )

    # Lista de máquinas correspondente ao setor selecionado
    if setor_painel != "Todos":
        lista_maquinas_painel = ["Todas"] + DICIONARIO_SETORES[setor_painel]
    else:
        todas_maquinas = []
        for m_list in DICIONARIO_SETORES.values():
            todas_maquinas.extend(m_list)
        lista_maquinas_painel = ["Todas"] + sorted(todas_maquinas)

    with col_maq_f:
        maquina_painel = st.selectbox(
            "⚙️ Máquina:",
            lista_maquinas_painel,
            index=0,
            key="p_maquina",
        )

    st.markdown("---")

    df_dados = pd.read_excel(ARQUIVO_FUSOS)
    df_filtrado = df_dados.copy()

    # Aplica os filtros: o Ano filtra exatamente o exercício anual escolhido
    df_filtrado = df_filtrado[df_filtrado["Ano"] == int(ano_painel)]

    if setor_painel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Setor"] == setor_painel]
    if maquina_painel != "Todas":
        df_filtrado = df_filtrado[df_filtrado["Maquina_TAG"] == maquina_painel]

    df_quebras_reais = df_filtrado[df_filtrado["Quantidade_Quebras"] > 0]

    # Subtítulo explicativo
    escopo_texto = []
    if setor_painel != "Todos":
        escopo_texto.append(setor_painel)
    if maquina_painel != "Todas":
        escopo_texto.append(f"Máquina {maquina_painel}")
    texto_cabecalho = " - ".join(escopo_texto) if escopo_texto else "Total Fábrica"

    # --- SEÇÃO 1: GRÁFICO DE LINHA (JANEIRO A DEZEMBRO DO ANO) ---
    st.subheader(f"📈 Tendência Mensal de Quebras em {ano_painel} ({texto_cabecalho})")

    # Garante os 12 meses fixos e sequenciais
    df_meses_base = pd.DataFrame({"Mes": lista_meses_puros})
    agrupado_mes = (
        df_filtrado.groupby("Mes")["Quantidade_Quebras"].sum().reset_index()
    )
    df_evolucao = pd.merge(df_meses_base, agrupado_mes, on="Mes", how="left").fillna(0)
    df_evolucao["Quantidade_Quebras"] = df_evolucao["Quantidade_Quebras"].astype(int)

    st.line_chart(
        data=df_evolucao.set_index("Mes")["Quantidade_Quebras"],
        use_container_width=True,
    )

    # --- SEÇÃO 2: MÉTRICAS RESUMIDAS ---
    if not df_quebras_reais.empty:
        total_quebras = df_filtrado["Quantidade_Quebras"].sum()
        total_maquinas_falharam = df_quebras_reais["Maquina_TAG"].nunique()

        agrupado_maq = (
            df_quebras_reais.groupby("Maquina_TAG")["Quantidade_Quebras"]
            .sum()
            .reset_index()
        )
        agrupado_maq = agrupado_maq.sort_values(
            by="Quantidade_Quebras", ascending=False
        )

        maquina_top = agrupado_maq.iloc[0]["Maquina_TAG"]
        qtd_top = agrupado_maq.iloc[0]["Quantidade_Quebras"]

        m1, m2, m3 = st.columns(3)
        m1.metric(f"Total Quebras em {ano_painel}", f"{total_quebras} unid.")
        
        if maquina_painel == "Todas":
            m2.metric(
                "Máquinas com Ocorrência", f"{total_maquinas_falharam} ativas"
            )
            m3.metric("Maior Ofensor", f"{maquina_top} ({qtd_top} quebras)")
        else:
            m2.metric("Máquina Analisada", maquina_painel)
            m3.metric("Total no Período", f"{total_quebras} quebras")

        st.markdown("---")

        # --- SEÇÃO 3: RANKING OU HISTÓRICO ---
        if maquina_painel == "Todas":
            graf_col, tab_col = st.columns([2, 1])
            with graf_col:
                st.subheader("Ranking por Equipamento")
                st.bar_chart(
                    data=agrupado_maq.set_index("Maquina_TAG")["Quantidade_Quebras"]
                )

            with tab_col:
                st.subheader("Consolidado")
                st.dataframe(
                    agrupado_maq.rename(
                        columns={
                            "Maquina_TAG": "Máquina",
                            "Quantidade_Quebras": "Total Falhas",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
        else:
            st.subheader(f"Histórico Detalhado: Máquina {maquina_painel} em {ano_painel}")
            df_detalhe_maq = df_quebras_reais[
                ["Ano", "Mes", "Setor", "Quantidade_Quebras", "Observacoes"]
            ]
            st.dataframe(df_detalhe_maq, use_container_width=True, hide_index=True)

    else:
        st.info(f"Nenhuma quebra registrada em {ano_painel} para os filtros selecionados.")

# ------------------------------------------
# 2. LANÇAMENTOS: FUSOS
# ------------------------------------------
elif tela == "Lançamento Fusos":
    st.header("🔩 Lançamentos: Grade de Quebras de Fusos")
    st.write(
        "Selecione o Ano e o Setor abaixo e aponte as quebras de cada máquina na aba do mês correspondente."
    )

    col_ano, col_setor, _ = st.columns([1.5, 2, 3])
    with col_ano:
        ano_selecionado = st.selectbox(
            "📅 Ano:", [2024, 2025, 2026, 2027, 2028], index=2
        )
    with col_setor:
        setor_selecionado = st.selectbox(
            "🏭 Setor:", list(DICIONARIO_SETORES.keys())
        )

    st.markdown("---")

    abas_meses = st.tabs(lista_meses_puros)
    maquinas_do_setor = DICIONARIO_SETORES[setor_selecionado]

    for idx, nome_mes in enumerate(lista_meses_puros):
        with abas_meses[idx]:
            st.subheader(
                f"Apontamento: {setor_selecionado} — {nome_mes}/{ano_selecionado}"
            )

            df_atual = pd.read_excel(ARQUIVO_FUSOS)

            df_filtrado = df_atual[
                (df_atual["Ano"] == ano_selecionado)
                & (df_atual["Mes"] == nome_mes)
                & (df_atual["Setor"] == setor_selecionado)
            ]

            dados_grade = []
            for maq in maquinas_do_setor:
                registro_existente = df_filtrado[
                    df_filtrado["Maquina_TAG"] == maq
                ]
                if not registro_existente.empty:
                    qtd = int(registro_existente.iloc[0]["Quantidade_Quebras"])
                    obs = str(registro_existente.iloc[0]["Observacoes"])
                    if obs == "nan":
                        obs = ""
                else:
                    qtd = 0
                    obs = ""

                dados_grade.append(
                    {
                        "Máquina": maq,
                        "Quantidade de Quebras": qtd,
                        "Observações": obs,
                    }
                )

            df_grade = pd.DataFrame(dados_grade)

            st.write(
                "👉 Dê dois cliques na célula de quantidade para editar os valores:"
            )

            tabela_editada = st.data_editor(
                df_grade,
                disabled=["Máquina"],
                hide_index=True,
                use_container_width=True,
                key=f"editor_{ano_selecionado}_{setor_selecionado}_{nome_mes}",
            )

            col_btn, _ = st.columns([2, 4])
            with col_btn:
                salvar_mes = st.button(
                    f"💾 Salvar {setor_selecionado} ({nome_mes}/{ano_selecionado})",
                    key=f"btn_{ano_selecionado}_{setor_selecionado}_{nome_mes}",
                    type="primary",
                )

            if salvar_mes:
                df_limpo = df_atual[
                    ~(
                        (df_atual["Ano"] == ano_selecionado)
                        & (df_atual["Mes"] == nome_mes)
                        & (df_atual["Setor"] == setor_selecionado)
                    )
                ]

                novos_registros = []
                for _, linha in tabela_editada.iterrows():
                    novos_registros.append(
                        {
                            "Ano": int(ano_selecionado),
                            "Mes": nome_mes,
                            "Setor": setor_selecionado,
                            "Maquina_TAG": linha["Máquina"],
                            "Quantidade_Quebras": int(
                                linha["Quantidade de Quebras"]
                            ),
                            "Observacoes": str(linha["Observações"]),
                        }
                    )

                df_final = pd.concat(
                    [df_limpo, pd.DataFrame(novos_registros)], ignore_index=True
                )
                df_final.to_excel(ARQUIVO_FUSOS, index=False)
                st.success(
                    f"✅ Fechamento do {setor_selecionado} para {nome_mes}/{ano_selecionado} salvo com sucesso!"
                )
                st.rerun()

            total_mes = tabela_editada["Quantidade de Quebras"].sum()
            st.caption(
                f"Total de fusos quebrados no {setor_selecionado} em {nome_mes}: **{total_mes} unid.**"
            )

# DEMAIS TELAS MANTIDAS
elif tela == "Correias":
    st.header("🔄 Lançamentos: Correias")
elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Máquinas")
