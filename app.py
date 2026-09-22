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

df_correias["Tipo_Correia_1"] = df_correias["Tipo_Correia_1"].astype(str)
df_correias["Data_Instalacao_1"] = df_correias["Data_Instalacao_1"].astype(str)
df_correias["Tipo_Correia_2"] = df_correias["Tipo_Correia_2"].astype(str)
df_correias["Data_Instalacao_2"] = df_correias["Data_Instalacao_2"].astype(str)

# Mapeamento oficial de ativos por setor[cite: 4, 5]
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


def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina


# ==========================================
# CÁLCULO E ANÁLISE DE CORREIAS
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
            return 1, dt_fmt, tempo_str  # Verde
        elif 365 < dias <= 547:
            return 2, dt_fmt, tempo_str  # Amarelo
        else:
            return 3, dt_fmt, tempo_str  # Vermelho
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
        v1 = str(ultimo.get("Tipo_Correia_1", "")).strip()
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

        v2 = str(ultimo.get("Tipo_Correia_2", "")).strip()
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

def montar_chips_categoria(lista_alvo, titulo):
    if not lista_alvo:
        return f"<b>{titulo}:</b> <i>Nenhuma correia registrada</i>"
    df_temp = pd.DataFrame(lista_alvo)
    contagem = df_temp["modelo"].value_counts().to_dict()
    chips = " ".join([
        f"<span class='chip-tt'>🏷️ {mod}: <b>{qtd} un.</b></span>"
        for mod, qtd in contagem.items()
    ])
    return f"<b>{titulo} ({len(lista_alvo)} un.):</b> &nbsp; {chips}"

chips_total = montar_chips_categoria(lista_correias_todas, "📦 TOTAL DE CORREIAS")
chips_novas = montar_chips_categoria(lista_correias_novas, "🟢 OPERAÇÃO NORMAL")
chips_meia = montar_chips_categoria(lista_correias_meia, "🟡 ATENÇÃO (MEIA-VIDA)")
chips_crit = montar_chips_categoria(lista_correias_criticas, "🔴 TROCA NECESSÁRIA")

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

        .header-bar {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 10px 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03);
            margin-bottom: 8px;
        }}
        .header-title {{
            font-size: 1.15rem;
            font-weight: 800;
            color: #0f172a;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        /* Container Principal dos Cards KPI */
        .card-kpi-container {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 10px 14px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
            position: relative;
            cursor: pointer;
            transition: all 0.15s ease;
        }}
        .card-kpi-container:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            border-color: #cbd5e1;
        }}
        .card-kpi-container::after {{
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 4px;
        }}
        .card-kpi-container.c-total::after {{ background: #475569; }}
        .card-kpi-container.c-ok::after {{ background: #10b981; }}
        .card-kpi-container.c-warn::after {{ background: #f59e0b; }}
        .card-kpi-container.c-crit::after {{ background: #ef4444; }}

        .kpi-val {{
            font-size: 1.45rem;
            font-weight: 800;
            line-height: 1;
            font-family: ui-monospace, monospace;
        }}
        .kpi-lbl {{
            font-size: 0.72rem;
            font-weight: 700;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 3px;
        }}

        /* Barra HUD que responde ao Hover dos Cards superiores */
        .hud-hover-display {{
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 9px;
            padding: 8px 16px;
            margin: 6px 0 8px 0;
            font-size: 0.84rem;
            font-weight: 600;
            color: #475569;
            min-height: 42px;
            display: flex;
            align-items: center;
            box-shadow: 0 2px 6px rgba(0,0,0,0.02);
        }}
        
        .chip-tt {{
            background: #f1f5f9;
            border: 1px solid #e2e8f0;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 0.8rem;
            color: #0f172a;
            font-weight: 700;
            margin-right: 4px;
            display: inline-block;
        }}

        /* Controlos dinâmicos acionados por hover puro */
        .hud-content {{ display: none; width: 100%; align-items: center; }}
        .hud-padrao {{ display: flex; width: 100%; align-items: center; gap: 8px; }}

        .kpi-section:hover .hud-padrao {{ display: none !important; }}
        .kpi-section .card-total:hover ~ .hud-hover-display .hud-total {{ display: flex !important; }}
        .kpi-section .card-novas:hover ~ .hud-hover-display .hud-novas {{ display: flex !important; }}
        .kpi-section .card-meia:hover ~ .hud-hover-display .hud-meia {{ display: flex !important; }}
        .kpi-section .card-crit:hover ~ .hud-hover-display .hud-crit {{ display: flex !important; }}

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
        .hud-detalhe.status-cinza {{ border-left-color: #64748b; }}

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
        .badge-cinza {{ background: #f1f5f9; color: #475569; }}

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
# 1. PAINEL GERENCIAL DE CORREIAS
# ------------------------------------------
if tela == "Painel Correias":
    st.markdown(
        """
        <div class="header-bar">
            <div class="header-title">
                <span>🔄</span>
                <span>Controle de correias</span>
            </div>
            <div>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#10b981;'></span> Nova (&le; 1a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#f59e0b;'></span> Meia-Vida (1-1,5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#ef4444;'></span> Troca Urgente (&gt; 1,5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#64748b;'></span> Sem Dados</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Bloco integrado dos Cards + Faixa Dinâmica por Hover (100% livre de sobreposição com máquinas)
    html_painel_kpi = f"""
    <div class="kpi-section">
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px;">
            <div class="card-kpi-container c-total card-total">
                <div>
                    <div class="kpi-lbl">Total de Correias</div>
                    <div class="kpi-val" style="color:#0f172a;">{len(lista_correias_todas)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">📦</div>
            </div>
            <div class="card-kpi-container c-ok card-novas">
                <div>
                    <div class="kpi-lbl">Operação Normal</div>
                    <div class="kpi-val" style="color:#059669;">{len(lista_correias_novas)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">🟢</div>
            </div>
            <div class="card-kpi-container c-warn card-meia">
                <div>
                    <div class="kpi-lbl">Atenção (Meia-Vida)</div>
                    <div class="kpi-val" style="color:#d97706;">{len(lista_correias_meia)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">🟡</div>
            </div>
            <div class="card-kpi-container c-crit card-crit">
                <div>
                    <div class="kpi-lbl">Troca Necessária</div>
                    <div class="kpi-val" style="color:#dc2626;">{len(lista_correias_criticas)}</div>
                </div>
                <div style="font-size:1.5rem; opacity:0.8;">🔴</div>
            </div>
        </div>
        <div class="hud-hover-display">
            <div class="hud-padrao">
                <span>💡</span>
                <span><b>Visualizador Operacional:</b> Passe o cursor pelos cards acima para discriminar as quantidades por modelo de correia ou clique numa máquina abaixo.</span>
            </div>
            <div class="hud-content hud-total">{chips_total}</div>
            <div class="hud-content hud-novas">{chips_novas}</div>
            <div class="hud-content hud-meia">{chips_meia}</div>
            <div class="hud-content hud-crit">{chips_crit}</div>
        </div>
    </div>
    """
    st.markdown(html_painel_kpi, unsafe_allow_html=True)

    # Balão HUD ao Clicar numa Máquina
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
                <span class="tag-pill">🔼 <b>Superior / Cabeceira:</b> {maq_sel['t1']} &nbsp;|&nbsp; 📅 {maq_sel['d1']} &nbsp;|&nbsp; ⏱️ <b>{maq_sel['uso1']}</b></span>
                """
            else:
                html_linhas += """
                <span class="tag-pill" style="opacity:0.75;">🔼 <b>Superior / Cabeceira:</b> Sem registro</span>
                """

            if maq_sel["tem_c2"]:
                html_linhas += f"""
                <span class="tag-pill">🔽 <b>Inferior / Traseira:</b> {maq_sel['t2']} &nbsp;|&nbsp; 📅 {maq_sel['d2']} &nbsp;|&nbsp; ⏱️ <b>{maq_sel['uso2']}</b></span>
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

    # Grelha de Máquinas em Mosaico Refinado (12 Colunas)
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
# 2. LANÇAMENTOS: CORREIAS (SUPERIOR / INFERIOR)
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
            
            c1 = str(ultimo.get("Tipo_Correia_1", ""))
            t1 = c1 if c1 != "nan" else ""
            d1_val = ultimo.get("Data_Instalacao_1", "")
            try:
                if pd.notna(d1_val) and str(d1_val).strip() not in ["", "nan"]:
                    dt1 = pd.to_datetime(d1_val).date()
            except Exception:
                dt1 = None

            c2 = str(ultimo.get("Tipo_Correia_2", ""))
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

            m1 = str(linha["Modelo (Superior / Cabeceira)"]).strip() if pd.notna(linha["Modelo (Superior / Cabeceira)"]) and str(linha["Modelo (Superior / Cabeceira)"]).strip() != "None" else ""
            m2 = str(linha["Modelo (Inferior / Traseira)"]).strip() if pd.notna(linha["Modelo (Inferior / Traseira)"]) and str(linha["Modelo (Inferior / Traseira)"]).strip() != "None" else ""

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
