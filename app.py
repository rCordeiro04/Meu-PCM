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

# Estilização CSS global
st.markdown(
    """
    <style>
        .card-kpi-bonito {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 16px 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 10px;
        }
        .kpi-lbl {
            font-size: 0.75rem;
            text-transform: uppercase;
            font-weight: 700;
            color: #64748b;
            letter-spacing: 0.05em;
        }
        .kpi-val {
            font-size: 1.6rem;
            font-weight: 800;
            color: #0f172a;
            margin-top: 2px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

ARQUIVO_FUSOS = "lancamentos_fusos_v5.xlsx"
ARQUIVO_CORREIAS = "lancamentos_correias_v4.xlsx"

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
    "Tipo_Correia_1",
    "Data_Instalacao_1",
    "Tipo_Correia_2",
    "Data_Instalacao_2",
]

OPCOES_TIPO_FUSO = ["FAG", "TEP", "M4BA", "MENEGATTO", "M4ZD", "USL"]

DICIONARIO_SETORES = {
    "Setor A": [f"L-{i:02d}" for i in range(1, 29)],
    "Setor B": [f"L-{i:02d}" for i in list(range(29, 47)) + list(range(50, 54))],
    "Setor Látex": ["B-71", "B-72", "B-73", "B-74", "B-75", "B-76", "B-77", "B-78", "B-79", "B-80", "B-83", "B-84", "B-85", "B-86", "B-87", "B-88", "B-89", "B-102", "B-103", "B-104"],
    "Setor Menegatto": ["B-47", "B-48", "B-49", "B-81", "B-82", "B-90", "B-91", "B-92", "B-93", "B-94", "B-95", "B-96", "B-97", "B-98", "B-99", "B-100", "B-101"],
}

if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Painel Fusos"
if "aba_setor_fuso" not in st.session_state:
    st.session_state.aba_setor_fuso = "Geral"

def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina

# Leitura blindada do Excel de Fusos
if os.path.exists(ARQUIVO_FUSOS):
    try:
        df_fusos = pd.read_excel(ARQUIVO_FUSOS)
        if not all(col in df_fusos.columns for col in colunas_fusos):
            df_fusos = pd.DataFrame(columns=colunas_fusos)
            df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
    except Exception:
        df_fusos = pd.DataFrame(columns=colunas_fusos)
        df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
else:
    df_fusos = pd.DataFrame(columns=colunas_fusos)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Leitura blindada do Excel de Correias
if os.path.exists(ARQUIVO_CORREIAS):
    try:
        df_correias = pd.read_excel(ARQUIVO_CORREIAS)
    except Exception:
        df_correias = pd.DataFrame(columns=colunas_correias)
        df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
else:
    df_correias = pd.DataFrame(columns=colunas_correias)
    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

for col in colunas_correias:
    if col not in df_correias.columns:
        df_correias[col] = ""

def formatar_modelo(val):
    if not val or str(val).strip() in ["", "nan", "None"]:
        return ""
    v_str = str(val).strip()
    if "." in v_str:
        partes = v_str.split(".")
        parte_inteira = partes[0]
        parte_decimal = partes[1].ljust(3, "0")[:3]
        return f"{parte_inteira}.{parte_decimal}"
    return v_str

# Barra Lateral (Sidebar)
with st.sidebar:
    st.title("⚙️ Portal PCM")
    st.caption("Planejamento e Controle de Manutenção")
    st.markdown("---")
    
    st.markdown("<div style='font-size:0.85rem; font-weight:700; color:#64748b; margin-bottom:6px;'>📊 PAINÉIS</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔩 Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Fusos" else "secondary"):
            navegar("Painel Fusos")
            st.rerun()
    with c2:
        if st.button("🔄 Correias", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Correias" else "secondary"):
            navegar("Painel Correias")
            st.rerun()

    st.markdown("<div style='font-size:0.85rem; font-weight:700; color:#64748b; margin-top:12px; margin-bottom:6px;'>📝 LANÇAMENTOS</div>", unsafe_allow_html=True)
    c3, c4 = st.columns(2)
    with c3:
        if st.button("🔩 Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Lançamento Fusos" else "secondary"):
            navegar("Lançamento Fusos")
            st.rerun()
    with c4:
        if st.button("🔄 Correias", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Correias" else "secondary"):
            navegar("Correias")
            st.rerun()

tela = st.session_state.pagina_atual
lista_meses_puros = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# ------------------------------------------
# PAINEL DE FUSOS
# ------------------------------------------
if tela == "Painel Fusos":
    st.title("Dashboard Fusos")
    st.markdown("---")
    st.info("Utilize as abas laterais para navegar entre os painéis e lançamentos.")

# ------------------------------------------
# LANÇAMENTO DE FUSOS (MODELO MATRICIAL: MÁQUINAS NA HORIZONTAL, MESES NA VERTICAL)
# ------------------------------------------
elif tela == "Lançamento Fusos":
    st.title("🔩 Lançamento: Fechamento de Fusos")
    st.caption("Modelo matricial: Máquinas nas colunas e meses/dias na vertical[cite: 1].")

    with st.container(border=True):
        col_ano, col_setor = st.columns([2, 3])
        with col_ano:
            ano_selecionado = st.selectbox("📅 Ano de Fechamento:", [2024, 2025, 2026, 2027, 2028], index=2)
        with col_setor:
            setor_selecionado = st.selectbox("🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()))

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(f"Matriz de Apontamento — {setor_selecionado} ({ano_selecionado})")

    df_atual = pd.read_excel(ARQUIVO_FUSOS)
    df_filtrado = df_atual[(df_atual["Ano"] == ano_selecionado) & (df_atual["Setor"] == setor_selecionado)]
    maquinas_do_setor = DICIONARIO_SETORES[setor_selecionado]

    # Criação da matriz pivotada (Meses nas linhas, Máquinas nas colunas)
    dados_matriz = []
    for mes in lista_meses_puros:
        linha_mes = {"Mês": mes}
        df_mes = df_filtrado[df_filtrado["Mes"] == mes]
        for maq in maquinas_do_setor:
            reg = df_mes[df_mes["Maquina_TAG"] == maq]
            linha_mes[maq] = int(reg.iloc[0]["Quantidade_Quebras"]) if not reg.empty else 0
        dados_matriz.append(linha_mes)

    df_matriz = pd.DataFrame(dados_matriz)

    # Configuração dinâmica de colunas para a tabela interativa
    config_colunas = {"Mês": st.column_config.TextColumn("Mês", disabled=True)}
    for maq in maquinas_do_setor:
        config_colunas[maq] = st.column_config.NumberColumn(maq, min_value=0, step=1, format="%d")

    tabela_editada = st.data_editor(
        df_matriz,
        column_config=config_colunas,
        hide_index=True,
        use_container_width=True,
        height=480,
        key=f"matriz_fusos_{setor_selecionado}_{ano_selecionado}"
    )

    if st.button("💾 Salvar Alterações da Matriz", type="primary"):
        # Remove os dados antigos do setor/ano e insere a nova matriz editada
        df_limpo = df_atual[~((df_atual["Ano"] == ano_selecionado) & (df_atual["Setor"] == setor_selecionado))]
        
        novos_registros = []
        for _, linha in tabela_editada.iterrows():
            mes_atual = linha["Mês"]
            for maq in maquinas_do_setor:
                novos_registros.append({
                    "Ano": int(ano_selecionado),
                    "Mes": mes_atual,
                    "Setor": setor_selecionado,
                    "Maquina_TAG": maq,
                    "Quantidade_Quebras": int(linha[maq]),
                    "Tipo_Fuso": OPCOES_TIPO_FUSO[0]
                })

        df_final = pd.concat([df_limpo, pd.DataFrame(novos_registros)], ignore_index=True)
        df_final.to_excel(ARQUIVO_FUSOS, index=False)
        st.success(f"✅ Fechamento do {setor_selecionado} salvo com sucesso no formato matricial!")
        st.rerun()

# ------------------------------------------
# OUTRAS TELAS
# ------------------------------------------
elif tela == "Correias":
    st.title("🔄 Gestão de Correias")
    st.info("Módulo de correias ativo.")
