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

# Ficheiros de dados blindados e separados
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

# Base de Fusos
if not os.path.exists(ARQUIVO_FUSOS):
    pd.DataFrame(columns=colunas_fusos).to_excel(ARQUIVO_FUSOS, index=False)

df_fusos = pd.read_excel(ARQUIVO_FUSOS)
if not all(col in df_fusos.columns for col in colunas_fusos):
    df_fusos = pd.DataFrame(columns=colunas_fusos)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Base de Correias
if not os.path.exists(ARQUIVO_CORREIAS):
    if os.path.exists("lancamentos_correias_v3.xlsx"):
        df_antigo = pd.read_excel("lancamentos_correias_v3.xlsx")
        df_migrado = pd.DataFrame(columns=colunas_correias)
        df_migrado["Setor"] = df_antigo.get("Setor", "")
        df_migrado["Maquina_TAG"] = df_antigo.get("Maquina_TAG", "")
        df_migrado["Tipo_Correia_1"] = df_antigo.get("Tipo_Correia", "")
        df_migrado["Data_Instalacao_1"] = df_antigo.get("Data_Instalacao", "")
        df_migrado["Tipo_Correia_2"] = ""
        df_migrado["Data_Instalacao_2"] = ""
        df_migrado.to_excel(ARQUIVO_CORREIAS, index=False)
        df_correias = df_migrado
    else:
        pd.DataFrame(columns=colunas_correias).to_excel(ARQUIVO_CORREIAS, index=False)
        df_correias = pd.DataFrame(columns=colunas_correias)
else:
    df_correias = pd.read_excel(ARQUIVO_CORREIAS)

for col in colunas_correias:
    if col not in df_correias.columns:
        df_correias[col] = ""


# Função de formatação para garantir sempre 3 dígitos após o ponto (ex: 36.100)
def formatar_modelo(val):
    if not val or str(val).strip() in ["", "nan", "None"]:
        return ""
    v_str = str(val).strip()
    if "." in v_str:
        partes = v_str.split(".")
        parte_inteira = partes[0]
        parte_decimal = partes[1]
        if len(parte_decimal) < 3:
            parte_decimal = parte_decimal.ljust(3, "0")
        elif len(parte_decimal) > 3:
            parte_decimal = parte_decimal[:3]
        return f"{parte_inteira}.{parte_decimal}"
    return v_str


df_correias["Tipo_Correia_1"] = df_correias["Tipo_Correia_1"].apply(formatar_modelo)
df_correias["Data_Instalacao_1"] = df_correias["Data_Instalacao_1"].astype(str)
df_correias["Tipo_Correia_2"] = df_correias["Tipo_Correia_2"].apply(formatar_modelo)
df_correias["Data_Instalacao_2"] = df_correias["Data_Instalacao_2"].astype(str)

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

mapa_setor_maquina = {}
for setor_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m:
        mapa_setor_maquina[m] = setor_nome

if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Painel Correias"

if "maq_clicada_cor" not in st.session_state:
    st.session_state.maq_clicada_cor = None

if "card_selecionado_kpi" not in st.session_state:
    st.session_state.card_selecionado_kpi = None

if "aba_setor_fuso" not in st.session_state:
    st.session_state.aba_setor_fuso = "Geral"


def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina


# ==========================================
# CÁLCULO E ANÁLISE DE CORREIAS (BLINDADO)
# ==========================================
todas_as_maquinas = []
for setor_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m:
        todas_as_maquinas.append(m)

data_hoje = date.today()
dados_maquinas = {}
css_botoes = []

lista_correias_todas = []
lista_correias_novas = []
lista_correias_meia = []
lista_correias_criticas = []


def avaliar_correia(dt_val):
    if not dt_val or dt_val == "nan" or dt_val.strip() == "":
        return None, "Sem registro", "Sem histórico"
    try:
        dt_inst = pd.to_datetime(dt_val).date()
        dt_fmt = dt_inst.strftime("%d/%m/%Y")
        dias = (data_hoje - dt_inst).days
        meses = round(dias / 30.4, 1)
        tempo_str = f"{meses}m ({dias}d)"

        if dias <= 365:
            return 1, dt_fmt, tempo_str
        elif 365 < dias <= 547:
            return 2, dt_fmt, tempo_str
        else:
            return 3, dt_fmt, tempo_str
    except Exception:
        return None, dt_val, "Data inválida"


for maq_tag in todas_as_maquinas:
    setor_m = mapa_setor_maquina[maq_tag]
    reg_maq = df_correias[
        (df_correias["Setor"] == setor_m) & (df_correias["Maquina_TAG"] == maq_tag)
    ]

    t1, d1_str, t1_uso, c1_score = "Não informada", "Sem registro", "Sem histórico", None
    t2, d2_str, t2_uso, c2_score = "Não informada", "Sem registro", "Sem histórico", None

    tem_c1 = False
    tem_c2 = False

    if not reg_maq.empty:
        ultimo = reg_maq.iloc[-1]
        v1 = formatar_modelo(ultimo.get("Tipo_Correia_1", ""))
        dt1_raw = str(ultimo.get("Data_Instalacao_1", "")).strip()
        if (v1 and v1 != "nan") or (dt1_raw and dt1_raw != "nan"):
            t1 = v1 if (v1 and v1 != "nan") else "Não informada"
            c1_score, d1_str, t1_uso = avaliar_correia(dt1_raw)
            tem_c1 = True
            reg_c1 = {"tag": maq_tag, "pos": "Superior", "modelo": t1, "data": d1_str, "uso": t1_uso, "setor": setor_m}
            lista_correias_todas.append(reg_c1)
            if c1_score == 1:
                lista_correias_novas.append(reg_c1)
            elif c1_score == 2:
                lista_correias_meia.append(reg_c1)
            elif c1_score == 3:
                lista_correias_criticas.append(reg_c1)

        v2 = formatar_modelo(ultimo.get("Tipo_Correia_2", ""))
        dt2_raw = str(ultimo.get("Data_Instalacao_2", "")).strip()
        if (v2 and v2 != "nan") or (dt2_raw and dt2_raw != "nan"):
            t2 = v2 if (v2 and v2 != "nan") else "Não informada"
            c2_score, d2_str, t2_uso = avaliar_correia(dt2_raw)
            tem_c2 = True
            reg_c2 = {"tag": maq_tag, "pos": "Inferior", "modelo": t2, "data": d2_str, "uso": t2_uso, "setor": setor_m}
            lista_correias_todas.append(reg_c2)
            if c2_score == 1:
                lista_correias_novas.append(reg_c2)
            elif c2_score == 2:
                lista_correias_meia.append(reg_c2)
            elif c2_score == 3:
                lista_correias_criticas.append(reg_c2)

    scores = [s for s in [c1_score, c2_score] if s is not None]
    if scores:
        pior_score = max(scores)
        if pior_score == 1:
            classe_card = "status-verde"
            cor_grad = "linear-gradient(135deg, #10b981 0%, #059669 100%)"
            cor_borda = "#047857"
            status_label = "Nova"
        elif pior_score == 2:
            classe_card = "status-amarelo"
            cor_grad = "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)"
            cor_borda = "#b45309"
            status_label = "Meia-Vida"
        else:
            classe_card = "status-vermelho"
            cor_grad = "linear-gradient(135deg, #ef4444 0%, #dc2626 100%)"
            cor_borda = "#b91c1c"
            status_label = "Troca Necessária"
    else:
        classe_card = "status-cinza"
        cor_grad = "linear-gradient(135deg, #64748b 0%, #475569 100%)"
        cor_borda = "#334155"
        status_label = "Sem Dados"

    dados_maquinas[maq_tag] = {
        "setor": setor_m,
        "t1": t1,
        "d1": d1_str,
        "uso1": t1_uso,
        "tem_c1": tem_c1,
        "t2": t2,
        "d2": d2_str,
        "uso2": t2_uso,
        "tem_c2": tem_c2,
        "status_label": status_label,
        "classe_card": classe_card,
    }

    chave_btn = f"btn_q_{maq_tag.replace('-', '_')}"
    css_botoes.append(
        f"""
        button[key="{chave_btn}"],
        div.st-key-{chave_btn} button {{
            background: {cor_grad} !important;
            color: #ffffff !important;
            border: 1px solid {cor_borda} !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.15) !important;
        }}
        button[key="{chave_btn}"]:hover,
        div.st-key-{chave_btn} button:hover {{
            filter: brightness(1.15) !important;
            transform: translateY(-2px) !important;
            box-shadow: 0 4px 10px rgba(0,0,0,0.22) !important;
        }}
        button[key="{chave_btn}"] p,
        div.st-key-{chave_btn} button p {{
            color: #ffffff !important;
            font-weight: 800 !important;
            letter-spacing: 0.4px !important;
        }}
        """
    )

regras_css_botoes = "\n".join(css_botoes)

st.markdown(
    f"""
    <style>
        .block-container {{
            padding-top: 4.4rem !important;
            padding-bottom: 1rem !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }}

        div[data-testid="column"] {{
            padding: 1px !important;
            margin: 0px !important;
        }}
        div[data-testid="stHorizontalBlock"] {{
            gap: 4px !important;
            margin-bottom: 4px !important;
        }}

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

        /* Botões do mosaico de máquinas */
        div[data-testid="stButton"] button {{
            padding: 0px !important;
            font-size: 0.8rem !important;
            height: 33px !important;
            min-height: 33px !important;
            line-height: 31px !important;
            border-radius: 7px !important;
            transition: all 0.12s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }}

        {regras_css_botoes}

        /* Overlay Invisível Aperfeiçoado */
        div.st-key-btn_inv_total,
        div.st-key-btn_inv_novas,
        div.st-key-btn_inv_meia,
        div.st-key-btn_inv_crit {{
            margin-top: -68px !important;
            opacity: 0 !important;
            z-index: 999 !important;
        }}
        div.st-key-btn_inv_total button,
        div.st-key-btn_inv_novas button,
        div.st-key-btn_inv_meia button,
        div.st-key-btn_inv_crit button {{
            height: 64px !important;
            width: 100% !important;
            cursor: pointer !important;
        }}

        /* Cards KPI Estilizados */
        .card-kpi-bonito {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 10px 14px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 5px rgba(0,0,0,0.02);
            position: relative;
            overflow: hidden;
            height: 64px;
            box-sizing: border-box;
            transition: all 0.15s ease;
        }}
        .card-kpi-bonito:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(0,0,0,0.08);
            border-color: #cbd5e1;
        }}
        .card-kpi-bonito::after {{
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 5px;
        }}
        .card-kpi-bonito.c-total::after {{ background: #475569; }}
        .card-kpi-bonito.c-ok::after {{ background: #10b981; }}
        .card-kpi-bonito.c-warn::after {{ background: #f59e0b; }}
        .card-kpi-bonito.c-crit::after {{ background: #ef4444; }}

        .kpi-val {{
            font-size: 1.35rem;
            font-weight: 800;
            line-height: 1;
            font-family: ui-monospace, monospace;
        }}
        .kpi-lbl {{
            font-size: 0.68rem;
            font-weight: 700;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 3px;
        }}

        /* Alerta de Manutenção */
        .alerta-manutencao {{
            background: #fef2f2;
            border: 1px solid #fecaca;
            border-left: 6px solid #ef4444;
            border-radius: 8px;
            padding: 8px 14px;
            margin: 6px 0 10px 0;
            font-size: 0.83rem;
            font-weight: 600;
            color: #991b1b;
            display: flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 1px 3px rgba(239,68,68,0.05);
        }}
        .chip-critico {{
            background: #fee2e2;
            border: 1px solid #fca5a5;
            color: #991b1b;
            padding: 2px 8px;
            border-radius: 5px;
            font-weight: 800;
            font-size: 0.8rem;
            display: inline-block;
            margin-right: 4px;
        }}

        .hud-detalhe {{
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            padding: 14px 18px;
            margin: 6px 0 12px 0;
            box-shadow: 0 6px 18px rgba(0,0,0,0.06);
            border-left: 6px solid #64748b;
            animation: fadeIn 0.15s ease-in;
        }}
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(-4px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        .hud-detalhe.status-verde {{ border-left-color: #10b981; }}
        .hud-detalhe.status-amarelo {{ border-left-color: #f59e0b; }}
        .hud-detalhe.status-vermelho {{ border-left-color: #ef4444; }}
        .hud-detalhe.status-cinza {{ border-left-color: #94a3b8; }}

        .tag-pill {{
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            padding: 5px 12px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 700;
            color: #334155;
            display: inline-block;
            margin-right: 8px;
        }}
        .badge-status {{
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .badge-verde {{ background: #d1fae5; color: #065f46; }}
        .badge-amarelo {{ background: #fef3c7; color: #92400e; }}
        .badge-vermelho {{ background: #fee2e2; color: #991b1b; }}
        .badge-cinza {{ background: #e2e8f0; color: #475569; }}

        .pill-legenda {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 0.78rem;
            font-weight: 700;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            padding: 4px 10px;
            border-radius: 20px;
            margin-left: 4px;
        }}
        .dot-legenda {{
            width: 9px;
            height: 9px;
            border-radius: 50%;
            display: inline-block;
        }}

        /* Estilo dos Cards Gráficos dos Setores */
        .chart-box-setor {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 14px 16px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.02);
            margin-bottom: 12px;
        }}
        .chart-box-setor-title {{
            font-size: 0.95rem;
            font-weight: 800;
            color: #0f172a;
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 10px;
            padding-bottom: 6px;
            border-bottom: 1px solid #f1f5f9;
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
# 1. PAINEL GERENCIAL DE CORREIAS (BLINDADO)
# ------------------------------------------
if tela == "Painel Correias":
    col_t, col_f1, col_f2, col_leg = st.columns([3.2, 1.4, 1.4, 5.0])

    with col_t:
        st.markdown(
            """
            <div style="height: 43px; display: flex; align-items: center; font-size: 2rem; font-weight: 900; color: #0f172a; padding-top: 5px;">
                Dashboard Correias
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_f1:
        lista_setores_filtro = ["Todos os Setores"] + list(DICIONARIO_SETORES.keys())
        filtro_setor = st.selectbox("Setor", lista_setores_filtro, key="filtro_setor_painel", label_visibility="collapsed")

    with col_f2:
        modelos_unicos = set()
        for r in lista_correias_todas:
            if r["modelo"] and r["modelo"] != "Não informada":
                modelos_unicos.add(r["modelo"])
        lista_modelos_filtro = ["Todos os Tipos"] + sorted(list(modelos_unicos))
        filtro_modelo = st.selectbox("Tipo", lista_modelos_filtro, key="filtro_modelo_painel", label_visibility="collapsed")

    with col_leg:
        st.markdown(
            """
            <div style="height: 43px; display: flex; align-items: center; justify-content: flex-end; padding-top: 5px;">
                <span class='pill-legenda'><span class='dot-legenda' style='background:#10b981;'></span> Nova (&le; 1a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#f59e0b;'></span> Meia-Vida (1-1,5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#ef4444;'></span> Troca Urgente (&gt; 1,5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#94a3b8;'></span> Sem Dados</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)

    todas_as_maquinas = []
    for setor_nome, lista_m in DICIONARIO_SETORES.items():
        if filtro_setor != "Todos os Setores" and setor_nome != filtro_setor:
            continue
        for m in lista_m:
            if filtro_modelo != "Todos os Tipos":
                reg_m = df_correias[(df_correias["Setor"] == setor_nome) & (df_correias["Maquina_TAG"] == m)]
                tem_mod = False
                if not reg_m.empty:
                    ult = reg_m.iloc[-1]
                    if formatar_modelo(ult.get("Tipo_Correia_1", "")) == filtro_modelo or formatar_modelo(ult.get("Tipo_Correia_2", "")) == filtro_modelo:
                        tem_mod = True
                if not tem_mod:
                    continue
            todas_as_maquinas.append(m)

    k1, k2, k3, k4 = st.columns(4)

    with k1:
        st.markdown(
            f"""
            <div class="card-kpi-bonito c-total">
                <div>
                    <div class="kpi-lbl">Correias Totais</div>
                    <div class="kpi-val" style="color:#0f172a;">{len(lista_correias_todas)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">📦</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button(" ", key="btn_inv_total", use_container_width=True):
            st.session_state.card_selecionado_kpi = "TODAS" if st.session_state.card_selecionado_kpi != "TODAS" else None
            st.rerun()

    with k2:
        st.markdown(
            f"""
            <div class="card-kpi-bonito c-ok">
                <div>
                    <div class="kpi-lbl">Correias Novas</div>
                    <div class="kpi-val" style="color:#059669;">{len(lista_correias_novas)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">🟢</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("  ", key="btn_inv_novas", use_container_width=True):
            st.session_state.card_selecionado_kpi = "NOVAS" if st.session_state.card_selecionado_kpi != "NOVAS" else None
            st.rerun()

    with k3:
        st.markdown(
            f"""
            <div class="card-kpi-bonito c-warn">
                <div>
                    <div class="kpi-lbl">Correias Meia Vida</div>
                    <div class="kpi-val" style="color:#d97706;">{len(lista_correias_meia)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">🟡</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("   ", key="btn_inv_meia", use_container_width=True):
            st.session_state.card_selecionado_kpi = "MEIA" if st.session_state.card_selecionado_kpi != "MEIA" else None
            st.rerun()

    with k4:
        st.markdown(
            f"""
            <div class="card-kpi-bonito c-crit">
                <div>
                    <div class="kpi-lbl">Correias Críticas</div>
                    <div class="kpi-val" style="color:#dc2626;">{len(lista_correias_criticas)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">🔴</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("    ", key="btn_inv_crit", use_container_width=True):
            st.session_state.card_selecionado_kpi = "CRITICAS" if st.session_state.card_selecionado_kpi != "CRITICAS" else None
            st.rerun()

    if st.session_state.card_selecionado_kpi is not None:
        sel = st.session_state.card_selecionado_kpi
        if sel == "TODAS":
            lista_alvo = lista_correias_todas
            titulo_faixa = "📦 Correias Totais Instaladas"
        elif sel == "NOVAS":
            lista_alvo = lista_correias_novas
            titulo_faixa = "🟢 Correias Novas (≤ 1 ano)"
        elif sel == "MEIA":
            lista_alvo = lista_correias_meia
            titulo_faixa = "🟡 Correias Meia Vida (1 a 1,5 anos)"
        else:
            lista_alvo = lista_correias_criticas
            titulo_faixa = "🔴 Correias Críticas (> 1,5 anos)"

        df_kpi_sel = pd.DataFrame(lista_alvo)
        if not df_kpi_sel.empty:
            contagem = df_kpi_sel["modelo"].value_counts().to_dict()
            chips_html = " ".join([
                f"<span class='chip-tt'>🏷️ {formatar_modelo(mod)}: <b>{qtd} un.</b></span>"
                for mod, qtd in contagem.items()
            ])
        else:
            chips_html = "<i>Nenhum registo encontrado</i>"

        c_fx, c_fechar = st.columns([6.2, 0.8])
        with c_fx:
            st.markdown(
                f"""
                <div class="faixa-visualizador ativa">
                    <span><b>{titulo_faixa} ({len(lista_alvo)} un.):</b> &nbsp; {chips_html}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with c_fechar:
            if st.button("✖ Fechar", key="btn_fechar_faixa_kpi"):
                st.session_state.card_selecionado_kpi = None
                st.rerun()
                
    if lista_correias_criticas:
        df_crit = pd.DataFrame(lista_correias_criticas)
        contagem_criticas = df_crit["modelo"].value_counts().to_dict()
        chips_crit_html = " ".join([
            f"<span class='chip-critico'>🏷️ {formatar_modelo(mod)}: <b>{qtd} un.</b></span>"
            for mod, qtd in contagem_criticas.items()
        ])
        st.markdown(
            f"""
            <div class="alerta-manutencao">
                <span>🚨 <b>Alerta de Manutenção:</b> &nbsp; {chips_crit_html}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)

    if st.session_state.maq_clicada_cor is not None:
        maq_sel = st.session_state.maq_clicada_cor
        classe_badge = (
            "badge-verde" if maq_sel["status_label"] == "Nova"
            else "badge-amarelo" if maq_sel["status_label"] == "Meia-Vida"
            else "badge-vermelho" if maq_sel["status_label"] == "Troca Necessária"
            else "badge-cinza"
        )

        c_box, c_close = st.columns([6.2, 0.8])
        with c_box:
            html_linhas = ""
            if maq_sel["tem_c1"]:
                html_linhas += f"""
                <span class="tag-pill">🔼 <b>Superior / Cabeceira:</b> {formatar_modelo(maq_sel['t1'])} &nbsp;|&nbsp; 📅 {maq_sel['d1']} &nbsp;|&nbsp; ⏱️ <b>{maq_sel['uso1']}</b></span>
                """
            else:
                html_linhas += """
                <span class="tag-pill" style="opacity:0.75;">🔼 <b>Superior / Cabeceira:</b> Sem registro</span>
                """

            if maq_sel["tem_c2"]:
                html_linhas += f"""
                <span class="tag-pill">🔽 <b>Inferior / Traseira:</b> {formatar_modelo(maq_sel['t2'])} &nbsp;|&nbsp; 📅 {maq_sel['d2']} &nbsp;|&nbsp; ⏱️ <b>{maq_sel['uso2']}</b></span>
                """
            else:
                html_linhas += """
                <span class="tag-pill" style="opacity:0.75;">🔽 <b>Inferior / Traseira:</b> Sem registro</span>
                """

            st.markdown(
                f"""
                <div class="hud-detalhe {maq_sel['classe_card']}">
                    <div style="display:flex; justify-content:space-between; align-items:center; width:100%; margin-bottom:6px;">
                        <div>
                            <span class="tag-pill" style="font-size:0.95rem; background:#0f172a; color:#ffffff; border-color:#0f172a;">⚙️ {maq_sel['tag']}</span>
                            <span class="tag-pill" style="background:#e2e8f0;">🏭 {maq_sel['setor']}</span>
                        </div>
                        <span class="badge-status {classe_badge}">{maq_sel['status_label']}</span>
                    </div>
                    <div style="display:flex; flex-wrap:wrap; gap:6px;">
                        {html_linhas}
                    </div>
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

    if todas_as_maquinas:
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
    else:
        st.info("Nenhuma máquina encontrada com os filtros selecionados.")

# ------------------------------------------
# 2. LANÇAMENTOS: CORREIAS (SUPERIOR / INFERIOR - BLINDADO)
# ------------------------------------------
elif tela == "Correias":
    st.title("🔄 Lançamento: Gestão de Correias")
    st.caption("Cadastro contínuo de modelos e datas de instalação por máquina (Superior / Cabeceira e Inferior / Traseira)")

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
        
        t1, dt1 = "", None
        t2, dt2 = "", None

        if not reg_existente.empty:
            ultimo = reg_existente.iloc[-1]
            
            c1 = formatar_modelo(ultimo.get("Tipo_Correia_1", ""))
            t1 = c1 if c1 != "nan" else ""
            d1_val = ultimo.get("Data_Instalacao_1", "")
            try:
                if pd.notna(d1_val) and str(d1_val).strip() not in ["", "nan"]:
                    dt1 = pd.to_datetime(d1_val).date()
            except Exception:
                dt1 = None

            c2 = formatar_modelo(ultimo.get("Tipo_Correia_2", ""))
            t2 = c2 if c2 != "nan" else ""
            d2_val = ultimo.get("Data_Instalacao_2", "")
            try:
                if pd.notna(d2_val) and str(d2_val).strip() not in ["", "nan"]:
                    dt2 = pd.to_datetime(d2_val).date()
            except Exception:
                dt2 = None

        dados_grade_cor.append(
            {
                "Máquina": maq,
                "Modelo (Superior / Cabeceira)": t1,
                "Data (Superior / Cabeceira)": dt1,
                "Modelo (Inferior / Traseira)": t2,
                "Data (Inferior / Traseira)": dt2,
            }
        )

    df_grade_cor = pd.DataFrame(dados_grade_cor)

    configuracao_colunas_cor = {
        "Máquina": st.column_config.TextColumn("Máquina", disabled=True),
        "Modelo (Superior / Cabeceira)": st.column_config.TextColumn("Modelo (Superior / Cabeceira)"),
        "Data (Superior / Cabeceira)": st.column_config.DateColumn("Data (Superior / Cabeceira)", format="DD/MM/YYYY"),
        "Modelo (Inferior / Traseira)": st.column_config.TextColumn("Modelo (Inferior / Traseira)"),
        "Data (Inferior / Traseira)": st.column_config.DateColumn("Data (Inferior / Traseira)", format="DD/MM/YYYY"),
    }

    tabela_editada_cor = st.data_editor(
        df_grade_cor,
        column_config=configuracao_colunas_cor,
        hide_index=True,
        use_container_width=True,
        height=440,
        key=f"editor_cor_superior_inferior_{setor_selecionado}",
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
            d1 = linha["Data (Superior / Cabeceira)"]
            dt1_str = ""
            if pd.notna(d1) and d1 is not None and str(d1).strip() != "":
                try:
                    dt1_str = pd.to_datetime(d1).strftime("%Y-%m-%d")
                except Exception:
                    dt1_str = ""

            d2 = linha["Data (Inferior / Traseira)"]
            dt2_str = ""
            if pd.notna(d2) and d2 is not None and str(d2).strip() != "":
                try:
                    dt2_str = pd.to_datetime(d2).strftime("%Y-%m-%d")
                except Exception:
                    dt2_str = ""

            m1 = formatar_modelo(linha["Modelo (Superior / Cabeceira)"]) if pd.notna(linha["Modelo (Superior / Cabeceira)"]) and str(linha["Modelo (Superior / Cabeceira)"]).strip() != "None" else ""
            m2 = formatar_modelo(linha["Modelo (Inferior / Traseira)"]) if pd.notna(linha["Modelo (Inferior / Traseira)"]) and str(linha["Modelo (Inferior / Traseira)"]).strip() != "None" else ""

            novos_registros_cor.append(
                {
                    "Setor": setor_selecionado,
                    "Maquina_TAG": linha["Máquina"],
                    "Tipo_Correia_1": m1,
                    "Data_Instalacao_1": dt1_str,
                    "Tipo_Correia_2": m2,
                    "Data_Instalacao_2": dt2_str,
                }
            )

        df_final_cor = pd.concat(
            [df_limpo_cor, pd.DataFrame(novos_registros_cor)], ignore_index=True
        )
        df_final_cor.to_excel(ARQUIVO_CORREIAS, index=False)
        st.success(f"✅ Dados de correias do {setor_selecionado} salvos com sucesso!")
        st.rerun()

# ------------------------------------------
# 3. PAINEL GERENCIAL DE FUSOS (NOVO PADRÃO EXECUTIVO DINÂMICO)
# ------------------------------------------
elif tela == "Painel Fusos":
    # Cabeçalho Superior Alinhado
    col_tf, col_ano_f = st.columns([3.8, 1.4])

    with col_tf:
        st.markdown(
            """
            <div style="height: 43px; display: flex; align-items: center; font-size: 2rem; font-weight: 900; color: #0f172a; padding-top: 5px;">
                Dashboard Fusos
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_ano_f:
        lista_anos_painel = [2024, 2025, 2026, 2027, 2028]
        ano_painel = st.selectbox("Ano", lista_anos_painel, index=2, key="filtro_ano_fusos_dash", label_visibility="collapsed")

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    # Botões de Setor para Exibir Informações Diferentes (Primeira opção: Geral)
    col_b_geral, col_b_sa, col_b_sb, col_b_latex, col_b_men = st.columns(5)
    
    with col_b_geral:
        tipo_geral = "primary" if st.session_state.aba_setor_fuso == "Geral" else "secondary"
        if st.button("🌐 Geral (Fábrica)", key="btn_fuso_aba_geral", use_container_width=True, type=tipo_geral):
            st.session_state.aba_setor_fuso = "Geral"
            st.rerun()

    with col_b_sa:
        tipo_sa = "primary" if st.session_state.aba_setor_fuso == "Setor A" else "secondary"
        if st.button("🏭 Setor A", key="btn_fuso_aba_sa", use_container_width=True, type=tipo_sa):
            st.session_state.aba_setor_fuso = "Setor A"
            st.rerun()

    with col_b_sb:
        tipo_sb = "primary" if st.session_state.aba_setor_fuso == "Setor B" else "secondary"
        if st.button("🏭 Setor B", key="btn_fuso_aba_sb", use_container_width=True, type=tipo_sb):
            st.session_state.aba_setor_fuso = "Setor B"
            st.rerun()

    with col_b_latex:
        tipo_latex = "primary" if st.session_state.aba_setor_fuso == "Setor Látex" else "secondary"
        if st.button("🌿 Setor Látex", key="btn_fuso_aba_latex", use_container_width=True, type=tipo_latex):
            st.session_state.aba_setor_fuso = "Setor Látex"
            st.rerun()

    with col_b_men:
        tipo_men = "primary" if st.session_state.aba_setor_fuso == "Setor Menegatto" else "secondary"
        if st.button("⚙️ Setor Menegatto", key="btn_fuso_aba_men", use_container_width=True, type=tipo_men):
            st.session_state.aba_setor_fuso = "Setor Menegatto"
            st.rerun()

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    # Carregamento e Filtro da Base de Fusos
    df_dados_fusos = pd.read_excel(ARQUIVO_FUSOS)
    df_fuso_ano = df_dados_fusos[df_dados_fusos["Ano"] == int(ano_painel)].copy()

    # ==========================================
    # CASO 1: ABA GERAL (FÁBRICA COMPLETA COM 4 GRÁFICOS)
    # ==========================================
    if st.session_state.aba_setor_fuso == "Geral":
        total_geral_quebras = int(df_fuso_ano["Quantidade_Quebras"].sum()) if not df_fuso_ano.empty else 0
        media_mensal_fabrica = round(total_geral_quebras / 12, 1)

        # Determinar Setor Mais Crítico e Máquina Mais Ofensora
        if not df_fuso_ano.empty and total_geral_quebras > 0:
            setor_ofensor = df_fuso_ano.groupby("Setor")["Quantidade_Quebras"].sum().sort_values(ascending=False).index[0]
            maq_ofensora = df_fuso_ano.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().sort_values(ascending=False).index[0]
            qtd_maq_ofensora = int(df_fuso_ano.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().max())
        else:
            setor_ofensor = "Nenhum"
            maq_ofensora = "Nenhuma"
            qtd_maq_ofensora = 0

        # Cards KPI Consolidados da Fábrica
        kf1, kf2, kf3, kf4 = st.columns(4)
        with kf1:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-total">
                    <div>
                        <div class="kpi-lbl">Total Quebras (Fábrica)</div>
                        <div class="kpi-val" style="color:#0f172a;">{total_geral_quebras}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">🔩</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kf2:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-ok">
                    <div>
                        <div class="kpi-lbl">Média Mensal Fábrica</div>
                        <div class="kpi-val" style="color:#059669;">{media_mensal_fabrica}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">📈</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kf3:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-warn">
                    <div>
                        <div class="kpi-lbl">Setor Mais Crítico</div>
                        <div class="kpi-val" style="color:#d97706; font-size:1.1rem;">{setor_ofensor}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">🏭</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kf4:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-crit">
                    <div>
                        <div class="kpi-lbl">Maior Ofensor Global</div>
                        <div class="kpi-val" style="color:#dc2626; font-size:1.1rem;">{maq_ofensora} ({qtd_maq_ofensora})</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">⚠️</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # Função geradora de gráfico por setor (Jan - Dez)
        def gerar_grafico_setor(nome_setor, cor_primaria):
            df_s = df_fuso_ano[df_fuso_ano["Setor"] == nome_setor]
            agrup_s = df_s.groupby("Mes")["Quantidade_Quebras"].sum().reset_index()
            df_base_meses = pd.DataFrame({"Mes": lista_meses_puros})
            df_consolidado = pd.merge(df_base_meses, agrup_s, on="Mes", how="left").fillna(0)
            df_consolidado["Quantidade_Quebras"] = df_consolidado["Quantidade_Quebras"].astype(int)
            total_setor = df_consolidado["Quantidade_Quebras"].sum()

            chart = (
                alt.Chart(df_consolidado)
                .mark_line(
                    point=alt.OverlayMarkDef(color=cor_primaria, size=50),
                    color=cor_primaria,
                    strokeWidth=2.5,
                )
                .encode(
                    x=alt.X("Mes:N", sort=lista_meses_puros, title=None, axis=alt.Axis(labelAngle=-45)),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                    tooltip=["Mes", "Quantidade_Quebras"],
                )
                .properties(height=210)
            )
            return chart, total_setor

        # Grelha 2x2 com os gráficos de cada um dos 4 setores
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            chart_a, tot_a = gerar_grafico_setor("Setor A", "#2563eb")
            st.markdown(
                f"""
                <div class="chart-box-setor">
                    <div class="chart-box-setor-title">
                        <span>🏭 Setor A</span>
                        <span style="color:#2563eb; font-size:0.85rem;">Total: {tot_a} fusos</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.altair_chart(chart_a, use_container_width=True)

            chart_latex, tot_latex = gerar_grafico_setor("Setor Látex", "#059669")
            st.markdown(
                f"""
                <div class="chart-box-setor">
                    <div class="chart-box-setor-title">
                        <span>🌿 Setor Látex</span>
                        <span style="color:#059669; font-size:0.85rem;">Total: {tot_latex} fusos</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.altair_chart(chart_latex, use_container_width=True)

        with col_g2:
            chart_b, tot_b = gerar_grafico_setor("Setor B", "#d97706")
            st.markdown(
                f"""
                <div class="chart-box-setor">
                    <div class="chart-box-setor-title">
                        <span>🏭 Setor B</span>
                        <span style="color:#d97706; font-size:0.85rem;">Total: {tot_b} fusos</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.altair_chart(chart_b, use_container_width=True)

            chart_men, tot_men = gerar_grafico_setor("Setor Menegatto", "#dc2626")
            st.markdown(
                f"""
                <div class="chart-box-setor">
                    <div class="chart-box-setor-title">
                        <span>⚙️ Setor Menegatto</span>
                        <span style="color:#dc2626; font-size:0.85rem;">Total: {tot_men} fusos</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.altair_chart(chart_men, use_container_width=True)

    # ==========================================
    # CASO 2: VISÃO ESPECÍFICA DE CADA SETOR
    # ==========================================
    else:
        setor_ativo = st.session_state.aba_setor_fuso
        df_setor = df_fuso_ano[df_fuso_ano["Setor"] == setor_ativo].copy()

        total_setor_quebras = int(df_setor["Quantidade_Quebras"].sum()) if not df_setor.empty else 0
        maquinas_setor_lista = DICIONARIO_SETORES[setor_ativo]
        qtd_maquinas_setor = len(maquinas_setor_lista)
        media_mensal_setor = round(total_setor_quebras / 12, 1)

        if not df_setor.empty and total_setor_quebras > 0:
            df_reais = df_setor[df_setor["Quantidade_Quebras"] > 0]
            maquinas_falharam_setor = df_reais["Maquina_TAG"].nunique()
            agrup_maq_setor = df_reais.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().reset_index().sort_values(by="Quantidade_Quebras", ascending=False)
            top_maq_setor = agrup_maq_setor.iloc[0]["Maquina_TAG"]
            qtd_top_setor = int(agrup_maq_setor.iloc[0]["Quantidade_Quebras"])
        else:
            maquinas_falharam_setor = 0
            agrup_maq_setor = pd.DataFrame(columns=["Maquina_TAG", "Quantidade_Quebras"])
            top_maq_setor = "Nenhuma"
            qtd_top_setor = 0

        # Cards KPI do Setor
        ks1, ks2, ks3, ks4 = st.columns(4)
        with ks1:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-total">
                    <div>
                        <div class="kpi-lbl">Total Quebras ({setor_ativo})</div>
                        <div class="kpi-val" style="color:#0f172a;">{total_setor_quebras}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">🔩</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with ks2:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-ok">
                    <div>
                        <div class="kpi-lbl">Média Mensal do Setor</div>
                        <div class="kpi-val" style="color:#059669;">{media_mensal_setor}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">📅</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with ks3:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-warn">
                    <div>
                        <div class="kpi-lbl">Ativos com Quebra</div>
                        <div class="kpi-val" style="color:#d97706;">{maquinas_falharam_setor} / {qtd_maquinas_setor}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">🏭</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with ks4:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-crit">
                    <div>
                        <div class="kpi-lbl">Maior Gargalo ({setor_ativo})</div>
                        <div class="kpi-val" style="color:#dc2626; font-size:1.1rem;">{top_maq_setor} ({qtd_top_setor})</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">⚠️</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # Gráfico de Linha Mensal do Setor + Gráfico por Tipo de Fuso
        c_linha_s, c_tipo_s = st.columns([1.6, 1.0])

        with c_linha_s:
            st.markdown(
                f"""
                <div class="chart-box-setor">
                    <div class="chart-box-setor-title">
                        <span>📈 Evolução Cronológica de Quebras ({setor_ativo} - {ano_painel})</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            df_base_meses = pd.DataFrame({"Mes": lista_meses_puros})
            agrup_mes_setor = df_setor.groupby("Mes")["Quantidade_Quebras"].sum().reset_index()
            df_evol_setor = pd.merge(df_base_meses, agrup_mes_setor, on="Mes", how="left").fillna(0)
            df_evol_setor["Quantidade_Quebras"] = df_evol_setor["Quantidade_Quebras"].astype(int)

            chart_linha_setor = (
                alt.Chart(df_evol_setor)
                .mark_line(
                    point=alt.OverlayMarkDef(color="#2563eb", size=60),
                    color="#2563eb",
                    strokeWidth=3,
                )
                .encode(
                    x=alt.X("Mes:N", sort=lista_meses_puros, title="Mês", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                    tooltip=["Mes", "Quantidade_Quebras"],
                )
                .properties(height=280)
            )
            st.altair_chart(chart_linha_setor, use_container_width=True)

        with c_tipo_s:
            st.markdown(
                f"""
                <div class="chart-box-setor">
                    <div class="chart-box-setor-title">
                        <span>🔩 Distribuição por Tipo de Fuso</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if not df_setor.empty and total_setor_quebras > 0:
                df_tipo_agrup = df_setor[df_setor["Quantidade_Quebras"] > 0].groupby("Tipo_Fuso")["Quantidade_Quebras"].sum().reset_index()
                chart_tipos = (
                    alt.Chart(df_tipo_agrup)
                    .mark_bar(color="#0f172a", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                    .encode(
                        x=alt.X("Quantidade_Quebras:Q", title="Total"),
                        y=alt.Y("Tipo_Fuso:N", sort="-x", title=None),
                        tooltip=["Tipo_Fuso", "Quantidade_Quebras"],
                    )
                    .properties(height=280)
                )
                st.altair_chart(chart_tipos, use_container_width=True)
            else:
                st.info(f"Sem registos de tipos de fuso para {setor_ativo} em {ano_painel}.")

        # Ranking e Tabela Detalhada do Setor
        if not agrup_maq_setor.empty:
            c_rk, c_tb = st.columns([1.5, 1.5])
            with c_rk:
                st.markdown(
                    f"""
                    <div class="chart-box-setor">
                        <div class="chart-box-setor-title">
                            <span>📊 Ranking de Máquinas com Falhas ({setor_ativo})</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                chart_ranking_s = (
                    alt.Chart(agrup_maq_setor)
                    .mark_bar(color="#ef4444", cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                    .encode(
                        x=alt.X("Maquina_TAG:N", sort="-y", title="Máquina"),
                        y=alt.Y("Quantidade_Quebras:Q", title="Total de Quebras"),
                        tooltip=["Maquina_TAG", "Quantidade_Quebras"],
                    )
                    .properties(height=280)
                )
                st.altair_chart(chart_ranking_s, use_container_width=True)

            with c_tb:
                st.markdown(
                    f"""
                    <div class="chart-box-setor">
                        <div class="chart-box-setor-title">
                            <span>📋 Consolidado de Ocorrências</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                df_tabela_s = agrup_maq_setor.rename(
                    columns={"Maquina_TAG": "Equipamento", "Quantidade_Quebras": "Total Quebras"}
                )
                st.dataframe(df_tabela_s, use_container_width=True, hide_index=True, height=280)

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
