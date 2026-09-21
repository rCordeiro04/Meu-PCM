import os
from datetime import date
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Portal PCM - Gestão de Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .metric-card {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 16px 20px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
            border-left: 5px solid #1E88E5;
            margin-bottom: 12px;
        }
        .metric-card.warning {
            border-left-color: #E53935;
        }
        .metric-card.success {
            border-left-color: #43A047;
        }
        .metric-label {
            font-size: 0.85rem;
            color: #616161;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }
        .metric-value {
            font-size: 1.8rem;
            color: #212121;
            font-weight: 700;
            line-height: 1.2;
        }
        .metric-sub {
            font-size: 0.8rem;
            color: #757575;
            margin-top: 4px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# Arquivos de dados blindados e separados
ARQUIVO_FUSOS = "lancamentos_fusos_v5.xlsx"
ARQUIVO_CORREIAS = "lancamentos_correias.xlsx"

colunas_fusos = [
    "Ano",
    "Mes",
    "Setor",
    "Maquina_TAG",
    "Quantidade_Quebras",
    "Tipo_Fuso",
]

colunas_correias = [
    "Ano",
    "Mes",
    "Setor",
    "Maquina_TAG",
    "Quantidade_Correias",
    "Observacoes",
]

OPCOES_TIPO_FUSO = ["FAG", "TEP", "M4BA", "MENEGATTO", "M4ZD", "USL"]

# Base de Fusos
if not os.path.exists(ARQUIVO_FUSOS):
    pd.DataFrame(columns=colunas_fusos).to_excel(ARQUIVO_FUSOS, index=False)

df_fusos = pd.read_excel(ARQUIVO_FUSOS)
if not all(col in df_fusos.columns for col in colunas_fusos):
    df_fusos = pd.DataFrame(columns=colunas_fusos)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Base de Correias
if not os.path.exists(ARQUIVO_CORREIAS):
    pd.DataFrame(columns=colunas_correias).to_excel(ARQUIVO_CORREIAS, index=False)

df_correias = pd.read_excel(ARQUIVO_CORREIAS)
if not all(col in df_correias.columns for col in colunas_correias):
    df_correias = pd.DataFrame(columns=colunas_correias)
    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

# Mapeamento oficial de ativos por setor
maquinas_setor_a = [f"L-{i:02d}" for i in range(1, 29)]

maquinas_setor_b = [
    "L-29", "L-30", "L-31", "L-32", "L-33", "L-34", "L-35", "L-36", "L-37",
    "L-38", "L-39", "L-40", "L-41", "L-42", "L-43", "L-44", "L-45", "L-46",
    "L-50", "L-51", "L-52", "L-53"
]

maquinas_setor_latex = [
    "B-71", "B-72", "B-73", "B-74", "B-75", "B-76", "B-77", "B-78", "B-79",
    "B-80", "B-83", "B-84", "B-85", "B-86", "B-87", "B-88", "B-89", "B-102",
    "B-103", "B-104"
]

maquinas_setor_menegatto = [
    "B-47", "B-48", "B-49", "B-81", "B-82", "B-90", "B-91", "B-92", "B-93",
    "B-94", "B-95", "B-96", "B-97", "B-98", "B-99", "B-100", "B-101"
]

DICIONARIO_SETORES = {
    "Setor A": maquinas_setor_a,
    "Setor B": maquinas_setor_b,
    "Setor Látex": maquinas_setor_latex,
    "Setor Menegatto": maquinas_setor_menegatto,
}

if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Correias"


def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina


# ==========================================
# BARRA LATERAL (SIDEBAR COM KEYS EXCLUSIVAS)
# ==========================================
with st.sidebar:
    st.title("⚙️ Portal PCM")
    st.caption("Planejamento e Controle de Manutenção")
    st.markdown("---")

    st.subheader("📊 Painéis")
    tipo_painel_fusos = (
        "primary"
        if st.session_state.pagina_atual == "Painel Fusos"
        else "secondary"
    )
    st.button(
        "🔩 Fusos",
        key="btn_nav_painel_fusos",
        use_container_width=True,
        type=tipo_painel_fusos,
        on_click=navegar,
        args=("Painel Fusos",),
    )

    st.markdown("---")

    st.subheader("📝 Lançamentos")
    tipo_fusos = (
        "primary"
        if st.session_state.pagina_atual == "Lançamento Fusos"
        else "secondary"
    )
    st.button(
        "🔩 Fusos",
        key="btn_nav_lancto_fusos",
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
        key="btn_nav_lancto_correias",
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
        key="btn_nav_lancto_preventiva",
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
        key="btn_nav_lancto_maquinas",
        use_container_width=True,
        type=tipo_maq,
        on_click=navegar,
        args=("Máquinas",),
    )

    st.markdown("---")
    st.caption("PCM • Versão Gerencial")

# ==========================================
# ÁREA PRINCIPAL
# ==========================================
tela = st.session_state.pagina_atual

lista_meses_puros = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# ------------------------------------------
# 1. LANÇAMENTOS: CORREIAS (ISOLADO)
# ------------------------------------------
if tela == "Correias":
    st.title("🔄 Lançamento: Fechamento Mensal de Correias")
    st.caption("Registo e acompanhamento de trocas de correias por setor e ativo")

    with st.container(border=True):
        col_ano, col_setor, _ = st.columns([1.5, 2, 3])
        with col_ano:
            ano_selecionado = st.selectbox(
                "📅 Ano de Fechamento:", [2024, 2025, 2026, 2027, 2028], index=2, key="sel_ano_correias"
            )
        with col_setor:
            setor_selecionado = st.selectbox(
                "🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_setor_correias"
            )

    st.markdown("<br>", unsafe_allow_html=True)

    abas_meses = st.tabs(lista_meses_puros)
    maquinas_do_setor = DICIONARIO_SETORES[setor_selecionado]

    for idx, nome_mes in enumerate(lista_meses_puros):
        with abas_meses[idx]:
            st.subheader(f"Apontamento de Correias — {setor_selecionado} ({nome_mes}/{ano_selecionado})")

            df_atual_cor = pd.read_excel(ARQUIVO_CORREIAS)

            df_filtrado_cor = df_atual_cor[
                (df_atual_cor["Ano"] == ano_selecionado)
                & (df_atual_cor["Mes"] == nome_mes)
                & (df_atual_cor["Setor"] == setor_selecionado)
            ]

            dados_grade_cor = []
            for maq in maquinas_do_setor:
                registro_existente = df_filtrado_cor[
                    df_filtrado_cor["Maquina_TAG"] == maq
                ]
                if not registro_existente.empty:
                    qtd = int(registro_existente.iloc[0]["Quantidade_Correias"])
                    obs = str(registro_existente.iloc[0]["Observacoes"])
                    if obs == "nan":
                        obs = ""
                else:
                    qtd = 0
                    obs = ""

                dados_grade_cor.append(
                    {
                        "Máquina": maq,
                        "Quantidade de Correias": qtd,
                        "Observações": obs,
                    }
                )

            df_grade_cor = pd.DataFrame(dados_grade_cor)

            configuracao_colunas_cor = {
                "Máquina": st.column_config.TextColumn(
                    "Máquina",
                    disabled=True,
                ),
                "Quantidade de Correias": st.column_config.NumberColumn(
                    "Qtd. Correias Substituídas",
                    min_value=0,
                    step=1,
                    format="%d",
                ),
                "Observações": st.column_config.TextColumn(
                    "Observações Técnicas / Causa",
                    max_chars=200,
                ),
            }

            tabela_editada_cor = st.data_editor(
                df_grade_cor,
                column_config=configuracao_colunas_cor,
                hide_index=True,
                use_container_width=True,
                key=f"editor_cor_{ano_selecionado}_{setor_selecionado}_{nome_mes}",
            )

            col_btn, _ = st.columns([2, 4])
            with col_btn:
                salvar_mes_cor = st.button(
                    f"💾 Salvar Correias: {setor_selecionado} ({nome_mes}/{ano_selecionado})",
                    key=f"btn_salvar_cor_{ano_selecionado}_{setor_selecionado}_{nome_mes}",
                    type="primary",
                )

            if salvar_mes_cor:
                df_limpo_cor = df_atual_cor[
                    ~(
                        (df_atual_cor["Ano"] == ano_selecionado)
                        & (df_atual_cor["Mes"] == nome_mes)
                        & (df_atual_cor["Setor"] == setor_selecionado)
                    )
                ]

                novos_registros_cor = []
                for _, linha in tabela_editada_cor.iterrows():
                    novos_registros_cor.append(
                        {
                            "Ano": int(ano_selecionado),
                            "Mes": nome_mes,
                            "Setor": setor_selecionado,
                            "Maquina_TAG": linha["Máquina"],
                            "Quantidade_Correias": int(linha["Quantidade de Correias"]),
                            "Observacoes": str(linha["Observações"]),
                        }
                    )

                df_final_cor = pd.concat(
                    [df_limpo_cor, pd.DataFrame(novos_registros_cor)], ignore_index=True
                )
                df_final_cor.to_excel(ARQUIVO_CORREIAS, index=False)
                st.success(
                    f"✅ Apontamento de Correias do {setor_selecionado} ({nome_mes}/{ano_selecionado}) salvo com sucesso!"
                )
                st.rerun()

            total_mes_cor = tabela_editada_cor["Quantidade de Correias"].sum()
            st.caption(
                f"Total de correias substituídas no {setor_selecionado} em {nome_mes}: **{total_mes_cor} unid.**"
            )

# ------------------------------------------
# 2. PAINEL GERENCIAL DE FUSOS (INTACTO)
# ------------------------------------------
elif tela == "Painel Fusos":
    st.title("🔩 Dashboard Gerencial — Quebras de Fusos")
    st.caption("Visão estratégica de falhas, confiabilidade de eixos e criticidade de ativos")

    with st.container(border=True):
        st.markdown("### 🔍 Filtros de Análise")
        c_ano, c_setor, c_maq, c_tipo = st.columns(4)

        lista_anos_painel = [2024, 2025, 2026, 2027, 2028]
        lista_setores_painel = ["Todos"] + list(DICIONARIO_SETORES.keys())
        lista_tipos_painel = ["Todos"] + OPCOES_TIPO_FUSO

        with c_ano:
            ano_painel = st.selectbox(
                "📅 Ano:", lista_anos_painel, index=2, key="filtro_ano_v5"
            )

        with c_setor:
            setor_painel = st.selectbox(
                "🏭 Setor:", lista_setores_painel, index=0, key="filtro_setor_v5"
            )

        if setor_painel != "Todos":
            lista_maquinas_painel = ["Todas"] + DICIONARIO_SETORES[setor_painel]
        else:
            todas_maquinas = []
            for m_list in DICIONARIO_SETORES.values():
                todas_maquinas.extend(m_list)
            lista_maquinas_painel = ["Todas"] + sorted(todas_maquinas)

        with c_maq:
            maquina_painel = st.selectbox(
                "⚙️ Máquina:", lista_maquinas_painel, index=0, key="filtro_maq_v5"
            )

        with c_tipo:
            tipo_fuso_painel = st.selectbox(
                "🔩 Tipo de Fuso:", lista_tipos_painel, index=0, key="filtro_tipo_v5"
            )

    df_dados = pd.read_excel(ARQUIVO_FUSOS)
    df_filtrado = df_dados.copy()

    df_filtrado = df_filtrado[df_filtrado["Ano"] == int(ano_painel)]

    if setor_painel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Setor"] == setor_painel]
    if maquina_painel != "Todas":
        df_filtrado = df_filtrado[df_filtrado["Maquina_TAG"] == maquina_painel]
    if tipo_fuso_painel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Tipo_Fuso"] == tipo_fuso_painel]

    df_quebras_reais = df_filtrado[df_filtrado["Quantidade_Quebras"] > 0]

    escopo_texto = []
    if setor_painel != "Todos":
        escopo_texto.append(setor_painel)
    if maquina_painel != "Todas":
        escopo_texto.append(f"Máq. {maquina_painel}")
    if tipo_fuso_painel != "Todos":
        escopo_texto.append(f"Fuso {tipo_fuso_painel}")
    texto_cabecalho = " • ".join(escopo_texto) if escopo_texto else "Fábrica Completa"

    st.markdown("<br>", unsafe_allow_html=True)

    total_quebras = int(df_filtrado["Quantidade_Quebras"].sum()) if not df_filtrado.empty else 0
    total_maquinas_falharam = df_quebras_reais["Maquina_TAG"].nunique() if not df_quebras_reais.empty else 0

    if not df_quebras_reais.empty:
        agrupado_maq = (
            df_quebras_reais.groupby("Maquina_TAG")["Quantidade_Quebras"]
            .sum()
            .reset_index()
            .sort_values(by="Quantidade_Quebras", ascending=False)
        )
        maquina_top = agrupado_maq.iloc[0]["Maquina_TAG"]
        qtd_top = int(agrupado_maq.iloc[0]["Quantidade_Quebras"])
    else:
        agrupado_maq = pd.DataFrame(columns=["Maquina_TAG", "Quantidade_Quebras"])
        maquina_top = "Nenhuma"
        qtd_top = 0

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Total de Quebras ({ano_painel})</div>
                <div class="metric-value">{total_quebras} <span style="font-size:1rem; font-weight:400; color:#757575;">fusos</span></div>
                <div class="metric-sub">Escopo: {texto_cabecalho}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with m2:
        if maquina_painel == "Todas":
            st.markdown(
                f"""
                <div class="metric-card success">
                    <div class="metric-label">Ativos com Ocorrência</div>
                    <div class="metric-value">{total_maquinas_falharam} <span style="font-size:1rem; font-weight:400; color:#757575;">máquinas</span></div>
                    <div class="metric-sub">Apresentaram ao menos 1 quebra</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f"""
                <div class="metric-card success">
                    <div class="metric-label">Ativo Monitorado</div>
                    <div class="metric-value">{maquina_painel}</div>
                    <div class="metric-sub">Setor: {setor_painel}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with m3:
        if maquina_painel == "Todas":
            st.markdown(
                f"""
                <div class="metric-card warning">
                    <div class="metric-label">Maior Ofensor (Gargalo)</div>
                    <div class="metric-value">{maquina_top}</div>
                    <div class="metric-sub">{qtd_top} quebra(s) apontadas</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            media_mensal = round(total_quebras / 12, 1)
            st.markdown(
                f"""
                <div class="metric-card warning">
                    <div class="metric-label">Taxa Média Mensal</div>
                    <div class="metric-value">{media_mensal} <span style="font-size:1rem; font-weight:400; color:#757575;">/mês</span></div>
                    <div class="metric-sub">Frequência média no ano</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    with st.container(border=True):
        st.subheader(f"📈 Curva de Tendência Mensal — {ano_painel}")
        st.caption(f"Evolução cronológica de quebras apontadas para: **{texto_cabecalho}**")

        df_meses_base = pd.DataFrame({"Mes": lista_meses_puros})
        agrupado_mes = (
            df_filtrado.groupby("Mes")["Quantidade_Quebras"].sum().reset_index()
        )
        df_evolucao = pd.merge(df_meses_base, agrupado_mes, on="Mes", how="left").fillna(0)
        df_evolucao["Quantidade_Quebras"] = df_evolucao["Quantidade_Quebras"].astype(int)

        df_evolucao["Mes"] = pd.Categorical(
            df_evolucao["Mes"], categories=lista_meses_puros, ordered=True
        )
        df_evolucao = df_evolucao.sort_values("Mes")

        st.line_chart(
            data=df_evolucao,
            x="Mes",
            y="Quantidade_Quebras",
            use_container_width=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    if not df_quebras_reais.empty:
        if maquina_painel == "Todas":
            c_graf, c_tab = st.columns([1.6, 1.2])

            with c_graf:
                with st.container(border=True):
                    st.subheader("📊 Ranking de Máquinas com Mais Falhas")
                    st.caption("Ativos ordenados pelo volume de fusos trocados")
                    st.bar_chart(
                        data=agrupado_maq.set_index("Maquina_TAG")["Quantidade_Quebras"],
                        use_container_width=True,
                    )

            with c_tab:
                with st.container(border=True):
                    st.subheader("📋 Tabela Consolidada")
                    st.caption("Visão detalhada por equipamento")
                    df_tabela = agrupado_maq.rename(
                        columns={
                            "Maquina_TAG": "Equipamento",
                            "Quantidade_Quebras": "Total Quebras",
                        }
                    )
                    st.dataframe(
                        df_tabela,
                        use_container_width=True,
                        hide_index=True,
                    )
        else:
            with st.container(border=True):
                st.subheader(f"📋 Histórico Operacional de Ocorrências — {maquina_painel}")
                df_detalhe_maq = df_quebras_reais[
                    ["Mes", "Setor", "Quantidade_Quebras", "Tipo_Fuso"]
                ].rename(
                    columns={
                        "Mes": "Mês de Referência",
                        "Setor": "Setor",
                        "Quantidade_Quebras": "Quebras Apontadas",
                        "Tipo_Fuso": "Tipo de Fuso",
                    }
                )
                st.dataframe(
                    df_detalhe_maq,
                    use_container_width=True,
                    hide_index=True,
                )
    else:
        st.info(f"Nenhum registro de quebra localizado em {ano_painel} com os filtros atuais.")

# ------------------------------------------
# 3. LANÇAMENTOS: FUSOS (INTACTO)
# ------------------------------------------
elif tela == "Lançamento Fusos":
    st.title("🔩 Lançamento: Fechamento Mensal de Fusos")
    st.caption("Preenchimento rápido de quebras por máquina, seleção de tipo de fuso e fechamento por setor")

    with st.container(border=True):
        col_ano, col_setor, _ = st.columns([1.5, 2, 3])
        with col_ano:
            ano_selecionado = st.selectbox(
                "📅 Ano de Fechamento:", [2024, 2025, 2026, 2027, 2028], index=2, key="sel_ano_fusos"
            )
        with col_setor:
            setor_selecionado = st.selectbox(
                "🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_setor_fusos"
            )

    st.markdown("<br>", unsafe_allow_html=True)

    abas_meses = st.tabs(lista_meses_puros)
    maquinas_do_setor = DICIONARIO_SETORES[setor_selecionado]

    for idx, nome_mes in enumerate(lista_meses_puros):
        with abas_meses[idx]:
            st.subheader(f"Apontamento — {setor_selecionado} ({nome_mes}/{ano_selecionado})")

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
                    tipo_salvo = str(registro_existente.iloc[0]["Tipo_Fuso"])
                    if tipo_salvo not in OPCOES_TIPO_FUSO:
                        tipo_salvo = OPCOES_TIPO_FUSO[0]
                else:
                    qtd = 0
                    tipo_salvo = OPCOES_TIPO_FUSO[0]

                dados_grade.append(
                    {
                        "Máquina": maq,
                        "Quantidade de Quebras": qtd,
                        "Tipo de Fuso": tipo_salvo,
                    }
                )

            df_grade = pd.DataFrame(dados_grade)

            configuracao_colunas = {
                "Máquina": st.column_config.TextColumn(
                    "Máquina",
                    disabled=True,
                ),
                "Quantidade de Quebras": st.column_config.NumberColumn(
                    "Quantidade de Quebras",
                    min_value=0,
                    step=1,
                    format="%d",
                ),
                "Tipo de Fuso": st.column_config.SelectboxColumn(
                    "Tipo de Fuso",
                    options=OPCOES_TIPO_FUSO,
                    required=True,
                ),
            }

            tabela_editada = st.data_editor(
                df_grade,
                column_config=configuracao_colunas,
                hide_index=True,
                use_container_width=True,
                key=f"editor_fusos_{ano_selecionado}_{setor_selecionado}_{nome_mes}",
            )

            col_btn, _ = st.columns([2, 4])
            with col_btn:
                salvar_mes = st.button(
                    f"💾 Salvar {setor_selecionado} ({nome_mes}/{ano_selecionado})",
                    key=f"btn_salvar_fusos_{ano_selecionado}_{setor_selecionado}_{nome_mes}",
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
                            "Tipo_Fuso": str(linha["Tipo de Fuso"]),
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
                f"Total de fusos apontados no {setor_selecionado} em {nome_mes}: **{total_mes} unid.**"
            )

# DEMAIS TELAS MANTIDAS
elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Máquinas")
