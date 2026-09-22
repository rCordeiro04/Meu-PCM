import os
from datetime import date, datetime
import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(
    page_title="Portal PCM - Gestão de Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Arquivos de dados blindados e separados
ARQUIVO_FUSOS = "lancamentos_fusos_v5.xlsx"
ARQUIVO_CORREIAS = "lancamentos_correias_v3.xlsx"

colunas_fusos = [
    "Ano",
    "Mes",
    "Setor",
    "Maquina_TAG",
    "Quantidade_Quebras",
    "Tipo_Fuso",
]

colunas_correias = [
    "Setor",
    "Maquina_TAG",
    "Tipo_Correia",
    "Data_Instalacao",
]

OPCOES_TIPO_FUSO = ["FAG", "TEP", "M4BA", "MENEGATTO", "M4ZD", "USL"]

# Base de Fusos (congelada)
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
    st.session_state.pagina_atual = "Painel Correias"

if "maq_clicada_cor" not in st.session_state:
    st.session_state.maq_clicada_cor = None


def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina


# ==========================================
# CÁLCULO DAS CORES E DADOS DAS MÁQUINAS
# ==========================================
todas_as_maquinas = []
mapa_setor_maquina = {}
for setor_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m:
        todas_as_maquinas.append(m)
        mapa_setor_maquina[m] = setor_nome

data_hoje = date.today()
dados_maquinas = {}
css_botoes = []

for maq_tag in todas_as_maquinas:
    setor_m = mapa_setor_maquina[maq_tag]
    reg_maq = df_correias[
        (df_correias["Setor"] == setor_m) & (df_correias["Maquina_TAG"] == maq_tag)
    ]

    tipo_txt = "Não informada"
    data_txt = "Sem registro"
    tempo_txt = "Sem histórico"
    classe_card = "status-cinza"
    
    # Cores sólidas padrão: cinza
    cor_fundo = "#94a3b8"
    cor_texto = "#ffffff"
    cor_borda = "#64748b"

    if not reg_maq.empty:
        ultimo = reg_maq.iloc[-1]
        tipo_val = str(ultimo["Tipo_Correia"]).strip()
        if tipo_val and tipo_val != "nan":
            tipo_txt = tipo_val

        dt_val = str(ultimo["Data_Instalacao"]).strip()
        if dt_val and dt_val != "nan":
            try:
                dt_inst = pd.to_datetime(dt_val).date()
                dt_fmt = dt_inst.strftime("%d/%m/%Y")
                data_txt = f"{dt_fmt}"

                dias = (data_hoje - dt_inst).days
                meses = round(dias / 30.4, 1)

                if dias <= 365:
                    classe_card = "status-verde"
                    cor_fundo = "#16a34a"   # Verde
                    cor_texto = "#ffffff"
                    cor_borda = "#15803d"
                    tempo_txt = f"{meses} meses ({dias} dias)"
                elif 365 < dias <= 547:
                    classe_card = "status-amarelo"
                    cor_fundo = "#facc15"   # Amarelo
                    cor_texto = "#713f12"
                    cor_borda = "#ca8a04"
                    tempo_txt = f"{meses} meses ({dias} dias)"
                else:
                    classe_card = "status-vermelho"
                    cor_fundo = "#dc2626"   # Vermelho
                    cor_texto = "#ffffff"
                    cor_borda = "#b91c1c"
                    tempo_txt = f"{meses} meses ({dias} dias)"
            except Exception:
                data_txt = f"{dt_val}"

    dados_maquinas[maq_tag] = {
        "setor": setor_m,
        "tipo": tipo_txt,
        "data": data_txt,
        "tempo": tempo_txt,
        "classe_card": classe_card,
    }

    # Injeção de regra CSS direta pela chave única do botão
    chave_btn = f"btn_q_{maq_tag.replace('-', '_')}"
    css_botoes.append(
        f"""
        button[key="{chave_btn}"],
        div.st-key-{chave_btn} button {{
            background-color: {cor_fundo} !important;
            color: {cor_texto} !important;
            border: 1px solid {cor_borda} !important;
        }}
        button[key="{chave_btn}"]:hover,
        div.st-key-{chave_btn} button:hover {{
            background-color: {cor_fundo} !important;
            filter: brightness(0.9) !important;
            color: {cor_texto} !important;
        }}
        button[key="{chave_btn}"] p,
        div.st-key-{chave_btn} button p {{
            color: {cor_texto} !important;
            font-weight: 800 !important;
        }}
        """
    )

regras_css_botoes = "\n".join(css_botoes)

st.markdown(
    f"""
    <style>
        /* Compactação geral da página */
        .block-container {{
            padding-top: 1rem !important;
            padding-bottom: 0.8rem !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
        }}
        
        /* Reduz o espaçamento horizontal e vertical entre os botões */
        div[data-testid="column"] {{
            padding: 0px 1px !important;
            margin: 0px !important;
        }}
        div[data-testid="stHorizontalBlock"] {{
            gap: 2px !important;
            margin-bottom: 2px !important;
        }}

        /* Trava contra scroll e zoom acidental em gráficos e tabelas */
        div[data-testid="stVegaLiteChart"] summary,
        div[data-testid="stVegaLiteChart"] .vega-actions {{
            display: none !important;
        }}
        div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {{
            overscroll-behavior: contain;
        }}
        div[data-testid="stDataFrame"] > div, div[data-testid="stDataEditor"] > div {{
            resize: none !important;
        }}

        /* Formato dos quadradinhos de máquinas */
        div[data-testid="stButton"] button {{
            padding: 0px !important;
            font-size: 0.78rem !important;
            font-weight: 800 !important;
            height: 28px !important;
            min-height: 28px !important;
            line-height: 26px !important;
            border-radius: 4px !important;
            transition: transform 0.08s ease, filter 0.08s ease !important;
        }}
        div[data-testid="stButton"] button:hover {{
            transform: scale(1.06) !important;
        }}

        /* REGRAS DINÂMICAS DE CORES INJETADAS */
        {regras_css_botoes}

        /* Card KPI Executivo */
        .metric-card {{
            background-color: #ffffff;
            border-radius: 8px;
            padding: 10px 14px;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.05);
            border-left: 4px solid #1E88E5;
            margin-bottom: 8px;
        }}
        .metric-card.warning {{ border-left-color: #E53935; }}
        .metric-card.success {{ border-left-color: #43A047; }}
        .metric-label {{
            font-size: 0.75rem;
            color: #616161;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.4px;
        }}
        .metric-value {{
            font-size: 1.3rem;
            color: #212121;
            font-weight: 700;
            line-height: 1.1;
        }}
        .metric-sub {{
            font-size: 0.72rem;
            color: #757575;
        }}

        /* Balão de Detalhes da Máquina Clicada */
        .card-balao-compacto {{
            border-radius: 6px;
            padding: 8px 14px;
            margin: 4px 0 8px 0;
            font-size: 0.88rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-left: 6px solid #94a3b8;
            box-shadow: 0 1px 4px rgba(0,0,0,0.06);
        }}
        .card-balao-compacto.status-verde {{
            background-color: #f0fdf4;
            border-left-color: #16a34a;
            color: #14532d;
        }}
        .card-balao-compacto.status-amarelo {{
            background-color: #fefce8;
            border-left-color: #ca8a04;
            color: #713f12;
        }}
        .card-balao-compacto.status-vermelho {{
            background-color: #fef2f2;
            border-left-color: #dc2626;
            color: #7f1d1d;
        }}
        .card-balao-compacto.status-cinza {{
            background-color: #f8fafc;
            border-left-color: #94a3b8;
            color: #334155;
        }}
        
        .amostra-cor {{
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 3px;
            vertical-align: middle;
            margin-right: 4px;
            border: 1px solid rgba(0,0,0,0.15);
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# BARRA LATERAL (SIDEBAR)
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

    tipo_painel_correias = (
        "primary"
        if st.session_state.pagina_atual == "Painel Correias"
        else "secondary"
    )
    st.button(
        "🔄 Correias",
        key="btn_nav_painel_correias",
        use_container_width=True,
        type=tipo_painel_correias,
        on_click=navegar,
        args=("Painel Correias",),
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
# 1. PAINEL GERENCIAL DE CORREIAS (QUADRADOS PINTADOS REALMENTE)
# ------------------------------------------
if tela == "Painel Correias":
    # Cabeçalho limpo com legenda gráfica compacta
    c_title, c_legenda = st.columns([1.6, 2.4])
    with c_title:
        st.markdown("<h4 style='margin:0; padding:0;'>🔄 Mapa Geral de Correias (Ativos da Fábrica)</h4>", unsafe_allow_html=True)
    with c_legenda:
        st.markdown(
            """
            <div style='text-align:right; font-size:0.82rem; font-weight:600; padding-top:4px;'>
                <span class='amostra-cor' style='background:#16a34a;'></span> Nova (&le;1 ano) &nbsp;|&nbsp;
                <span class='amostra-cor' style='background:#facc15;'></span> Meia-Vida (1 a 1,5 anos) &nbsp;|&nbsp;
                <span class='amostra-cor' style='background:#dc2626;'></span> Fim de Vida (&gt;1,5 anos) &nbsp;|&nbsp;
                <span class='amostra-cor' style='background:#94a3b8;'></span> Sem Apontamento
            </div>
            """,
            unsafe_allow_html=True,
        )

    # SE HOUVER MÁQUINA CLICADA: BALÃO EM LINHA NO TOPO
    if st.session_state.maq_clicada_cor is not None:
        maq_sel = st.session_state.maq_clicada_cor
        c_box, c_close = st.columns([6.2, 0.8])
        with c_box:
            st.markdown(
                f"""
                <div class="card-balao-compacto {maq_sel['classe_card']}">
                    <span>⚙️ <b>Ativo: {maq_sel['tag']}</b> ({maq_sel['setor']}) &nbsp;|&nbsp; 🏷️ <b>Modelo:</b> {maq_sel['tipo']} &nbsp;|&nbsp; 📅 <b>Instalação:</b> {maq_sel['data']}</span>
                    <span>⏱️ <b>Uso:</b> {maq_sel['tempo']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with c_close:
            st.write("")
            if st.button("✖ Fechar", key="btn_fechar_balao_topo"):
                st.session_state.maq_clicada_cor = None
                st.rerun()

    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)

    # GRELHA ULTRA-COMPACTA: 12 QUADRADINHOS POR LINHA COM CORES DIRETAS
    COLS_GRELHA = 12
    linhas_grid = [todas_as_maquinas[i:i + COLS_GRELHA] for i in range(0, len(todas_as_maquinas), COLS_GRELHA)]

    for linha in linhas_grid:
        cols = st.columns(COLS_GRELHA)
        for idx_col, maq_tag in enumerate(linha):
            info = dados_maquinas[maq_tag]
            chave_btn = f"btn_q_{maq_tag.replace('-', '_')}"
            with cols[idx_col]:
                if st.button(
                    maq_tag,
                    key=chave_btn,
                    use_container_width=True,
                ):
                    st.session_state.maq_clicada_cor = {
                        "tag": maq_tag,
                        **info,
                    }
                    st.rerun()

# ------------------------------------------
# 2. LANÇAMENTOS: CORREIAS (INTACTO)
# ------------------------------------------
elif tela == "Correias":
    st.title("🔄 Lançamento: Gestão de Correias")
    st.caption("Cadastro contínuo de tipos de correias e datas de instalação por máquina")

    with st.container(border=True):
        col_setor, _ = st.columns([2, 3])
        with col_setor:
            setor_selecionado = st.selectbox(
                "🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_setor_correias_direto"
            )

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(f"Apontamento de Correias — {setor_selecionado}")

    df_atual_cor = pd.read_excel(ARQUIVO_CORREIAS)
    df_filtrado_cor = df_atual_cor[df_atual_cor["Setor"] == setor_selecionado]
    maquinas_do_setor = DICIONARIO_SETORES[setor_selecionado]

    dados_grade_cor = []
    for maq in maquinas_do_setor:
        reg_existente = df_filtrado_cor[df_filtrado_cor["Maquina_TAG"] == maq]
        if not reg_existente.empty:
            tipo_c = str(reg_existente.iloc[-1]["Tipo_Correia"])
            if tipo_c == "nan":
                tipo_c = ""
            dt_val = reg_existente.iloc[-1]["Data_Instalacao"]
            try:
                if pd.notna(dt_val) and str(dt_val).strip() != "":
                    dt_inst = pd.to_datetime(dt_val).date()
                else:
                    dt_inst = None
            except Exception:
                dt_inst = None
        else:
            tipo_c = ""
            dt_inst = None

        dados_grade_cor.append(
            {
                "Máquina": maq,
                "Tipo / Modelo da Correia": tipo_c,
                "Data de Instalação": dt_inst,
            }
        )

    df_grade_cor = pd.DataFrame(dados_grade_cor)

    configuracao_colunas_cor = {
        "Máquina": st.column_config.TextColumn(
            "Máquina",
            disabled=True,
        ),
        "Tipo / Modelo da Correia": st.column_config.TextColumn(
            "Tipo / Modelo da Correia",
        ),
        "Data de Instalação": st.column_config.DateColumn(
            "Data de Instalação",
            format="DD/MM/YYYY",
        ),
    }

    tabela_editada_cor = st.data_editor(
        df_grade_cor,
        column_config=configuracao_colunas_cor,
        hide_index=True,
        use_container_width=True,
        height=440,
        key=f"editor_cor_direto_{setor_selecionado}",
    )

    col_btn, _ = st.columns([2, 4])
    with col_btn:
        salvar_cor = st.button(
            f"💾 Salvar Correias: {setor_selecionado}",
            key=f"btn_salvar_cor_direto_{setor_selecionado}",
            type="primary",
        )

    if salvar_cor:
        df_limpo_cor = df_atual_cor[df_atual_cor["Setor"] != setor_selecionado]

        novos_registros_cor = []
        for _, linha in tabela_editada_cor.iterrows():
            d_inst = linha["Data de Instalação"]
            
            dt_str = ""
            if pd.notna(d_inst) and d_inst is not None and str(d_inst).strip() != "":
                try:
                    dt_str = pd.to_datetime(d_inst).strftime("%Y-%m-%d")
                except Exception:
                    dt_str = ""

            novos_registros_cor.append(
                {
                    "Setor": setor_selecionado,
                    "Maquina_TAG": linha["Máquina"],
                    "Tipo_Correia": str(linha["Tipo / Modelo da Correia"]).strip() if pd.notna(linha["Tipo / Modelo da Correia"]) and str(linha["Tipo / Modelo da Correia"]).strip() != "None" else "",
                    "Data_Instalacao": dt_str,
                }
            )

        df_final_cor = pd.concat(
            [df_limpo_cor, pd.DataFrame(novos_registros_cor)], ignore_index=True
        )
        df_final_cor.to_excel(ARQUIVO_CORREIAS, index=False)
        st.success(f"✅ Dados de correias do {setor_selecionado} salvos com sucesso!")
        st.rerun()

# ------------------------------------------
# 3. PAINEL GERENCIAL DE FUSOS (INTACTO)
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

        chart_linha = (
            alt.Chart(df_evolucao)
            .mark_line(point=alt.OverlayMarkDef(color="#1E88E5", size=60), color="#1E88E5", strokeWidth=3)
            .encode(
                x=alt.X("Mes:N", sort=lista_meses_puros, title="Mês", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Quantidade_Quebras:Q", title="Qtd. Quebras"),
                tooltip=["Mes", "Quantidade_Quebras"],
            )
            .properties(height=320)
        )
        st.altair_chart(chart_linha, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if not df_quebras_reais.empty:
        if maquina_painel == "Todas":
            c_graf, c_tab = st.columns([1.6, 1.2])

            with c_graf:
                with st.container(border=True):
                    st.subheader("📊 Ranking de Máquinas com Mais Falhas")
                    st.caption("Ativos ordenados pelo volume de fusos trocados")
                    
                    chart_barras = (
                        alt.Chart(agrupado_maq)
                        .mark_bar(color="#1E88E5", cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                        .encode(
                            x=alt.X("Maquina_TAG:N", sort="-y", title="Equipamento"),
                            y=alt.Y("Quantidade_Quebras:Q", title="Total de Quebras"),
                            tooltip=["Maquina_TAG", "Quantidade_Quebras"],
                        )
                        .properties(height=320)
                    )
                    st.altair_chart(chart_barras, use_container_width=True)

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
                        height=320,
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
                    height=320,
                )
    else:
        st.info(f"Nenhum registo de quebra localizado em {ano_painel} com os filtros atuais.")

# ------------------------------------------
# 4. LANÇAMENTOS: FUSOS (INTACTO)
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
                height=440,
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
