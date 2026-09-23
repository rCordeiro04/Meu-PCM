import calendar
import io
import os
import shutil
from datetime import date, datetime
import altair as alt
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Portal PCM - Gestão de Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Arquivos de dados principais
ARQUIVO_FUSOS = "lancamentos_fusos_v5.xlsx"
ARQUIVO_CORREIAS = "lancamentos_correias_v4.xlsx"
ARQUIVO_PARADAS = "lancamentos_paradas_v1.xlsx"
ARQUIVO_PENDENCIAS = "lancamentos_pendencias_v1.xlsx"

COLUNAS_FUSOS = [
    "Ano",
    "Mes",
    "Dia",
    "Setor",
    "Maquina_TAG",
    "Quantidade_Quebras",
    "Tipo_Fuso",
]

COLUNAS_CORREIAS = [
    "Setor",
    "Maquina_TAG",
    "Tipo_Correia_1",
    "Data_Instalacao_1",
    "Tipo_Correia_2",
    "Data_Instalacao_2",
]

COLUNAS_PARADAS = [
    "Data",
    "Setor",
    "Maquina_TAG",
    "Tipo_Manutencao",
    "Descricao_Servico",
    "Tempo_Parado_Horas",
]

COLUNAS_PENDENCIAS = [
    "Setor",
    "Maquina_TAG",
    "Descricao_Pendencia",
    "Prioridade",
    "Status",
]

OPCOES_TIPO_FUSO = ["FAG", "TEP", "M4BA", "MENEGATTO", "M4ZD", "USL"]

LISTA_MESES_PUROS = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

MAPA_MES_ABREV = {
    "Janeiro": "JAN", "Fevereiro": "FEV", "Março": "MAR", "Abril": "ABR",
    "Maio": "MAI", "Junho": "JUN", "Julho": "JUL", "Agosto": "AGO",
    "Setembro": "SET", "Outubro": "OUT", "Novembro": "NOV", "Dezembro": "DEZ"
}
ORDEM_MESES_ABREV = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

DICIONARIO_SETORES = {
    "Setor A": [f"L-{i:02d}" for i in range(1, 29)],
    "Setor B": [
        "L-29", "L-30", "L-31", "L-32", "L-33", "L-34", "L-35", "L-36", "L-37",
        "L-38", "L-39", "L-40", "L-41", "L-42", "L-43", "L-44", "L-45", "L-46",
        "L-50", "L-51", "L-52", "L-53"
    ],
    "Setor Látex": [
        "B-71", "B-72", "B-73", "B-74", "B-75", "B-76", "B-77", "B-78", "B-79",
        "B-80", "B-83", "B-84", "B-85", "B-86", "B-87", "B-88", "B-89", "B-102",
        "B-103", "B-104"
    ],
    "Setor Menegatto": [
        "B-47", "B-48", "B-49", "B-81", "B-82", "B-90", "B-91", "B-92", "B-93",
        "B-94", "B-95", "B-96", "B-97", "B-98", "B-99", "B-100", "B-101", "B-107", "B-108"
    ],
}

def gerar_backup_seguro(caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        os.makedirs("backups", exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arq = os.path.basename(caminho_arquivo)
        try:
            shutil.copy2(caminho_arquivo, os.path.join("backups", f"{ts}_{nome_arq}"))
        except Exception:
            pass

def formatar_modelo(val):
    if val is None or pd.isna(val):
        return ""
    v_str = str(val).strip()
    if v_str.lower() in ["", "nan", "none", "nat"]:
        return ""
    if "." in v_str:
        partes = v_str.split(".")
        p_dec = partes[1][:3].ljust(3, "0")
        return f"{partes[0]}.{p_dec}"
    return v_str

# =========================================================================
# LEITURA OTIMIZADA COM CACHE
# =========================================================================
@st.cache_data(show_spinner=False)
def carregar_dados():
    if os.path.exists(ARQUIVO_FUSOS):
        try:
            df_f = pd.read_excel(ARQUIVO_FUSOS)
            if "Dia" not in df_f.columns:
                df_f["Dia"] = 1
        except Exception:
            df_f = pd.DataFrame(columns=COLUNAS_FUSOS)
    else:
        df_f = pd.DataFrame(columns=COLUNAS_FUSOS)
        df_f.to_excel(ARQUIVO_FUSOS, index=False)

    if os.path.exists(ARQUIVO_CORREIAS):
        try:
            df_c = pd.read_excel(ARQUIVO_CORREIAS)
        except Exception:
            df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)
    else:
        df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)
        df_c.to_excel(ARQUIVO_CORREIAS, index=False)

    for col in COLUNAS_CORREIAS:
        if col not in df_c.columns:
            df_c[col] = ""
        df_c[col] = df_c[col].astype(object)

    df_c["Tipo_Correia_1"] = df_c["Tipo_Correia_1"].apply(formatar_modelo)
    df_c["Data_Instalacao_1"] = df_c["Data_Instalacao_1"].astype(str).replace({"nan": "", "NaT": "", "None": ""})
    df_c["Tipo_Correia_2"] = df_c["Tipo_Correia_2"].apply(formatar_modelo)
    df_c["Data_Instalacao_2"] = df_c["Data_Instalacao_2"].astype(str).replace({"nan": "", "NaT": "", "None": ""})

    if os.path.exists(ARQUIVO_PARADAS):
        try:
            df_p = pd.read_excel(ARQUIVO_PARADAS)
        except Exception:
            df_p = pd.DataFrame(columns=COLUNAS_PARADAS)
    else:
        df_p = pd.DataFrame(columns=COLUNAS_PARADAS)
        df_p.to_excel(ARQUIVO_PARADAS, index=False)

    for col in COLUNAS_PARADAS:
        if col not in df_p.columns:
            df_p[col] = 0.0 if col == "Tempo_Parado_Horas" else ""

    if os.path.exists(ARQUIVO_PENDENCIAS):
        try:
            df_pend = pd.read_excel(ARQUIVO_PENDENCIAS)
        except Exception:
            df_pend = pd.DataFrame(columns=COLUNAS_PENDENCIAS)
    else:
        df_pend = pd.DataFrame(columns=COLUNAS_PENDENCIAS)
        df_pend.to_excel(ARQUIVO_PENDENCIAS, index=False)

    for col in COLUNAS_PENDENCIAS:
        if col not in df_pend.columns:
            df_pend[col] = ""

    return df_f, df_c, df_p, df_pend

df_fusos, df_correias, df_paradas, df_pendencias = carregar_dados()

def invalidar_cache():
    st.cache_data.clear()

def obter_maquinas_setor(setor_nome, df_c=None, df_f=None):
    base = set(DICIONARIO_SETORES.get(setor_nome, []))
    if df_c is not None and not df_c.empty:
        base.update(df_c[df_c["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique())
    if df_f is not None and not df_f.empty:
        base.update(df_f[df_f["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique())
    return sorted(list(base))

mapa_setor_maquina = {}
for s_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m:
        mapa_setor_maquina[m] = s_nome

for _, r in df_correias.iterrows():
    if pd.notna(r.get("Maquina_TAG")) and pd.notna(r.get("Setor")):
        mapa_setor_maquina[str(r["Maquina_TAG"]).strip()] = str(r["Setor"]).strip()
for _, r in df_fusos.iterrows():
    if pd.notna(r.get("Maquina_TAG")) and pd.notna(r.get("Setor")):
        mapa_setor_maquina[str(r["Maquina_TAG"]).strip()] = str(r["Setor"]).strip()

if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Painel Fusos"
if "maq_clicada_cor" not in st.session_state:
    st.session_state.maq_clicada_cor = None
if "aba_setor_fuso" not in st.session_state:
    st.session_state.aba_setor_fuso = "Geral"

def navegar(p):
    st.session_state.pagina_atual = p

# ==========================================
# PROCESSAMENTO DE CORREIAS OTIMIZADO
# ==========================================
data_hoje = date.today()
dados_maquinas = {}

lista_correias_todas = []
lista_correias_novas = []
lista_correias_meia = []
lista_correias_criticas = []

def avaliar_correia(dt_val):
    if not dt_val or str(dt_val).strip() in ["", "nan", "NaT", "None"]:
        return None, "Sem registro", "Sem histórico"
    try:
        dt_inst = pd.to_datetime(dt_val).date()
        dt_fmt = dt_inst.strftime("%d/%m/%Y")
        dias = (data_hoje - dt_inst).days
        meses = round(dias / 30.4, 1)
        tempo_str = f"{meses}m ({dias}d)"
        return (1 if dias <= 365 else 2 if dias <= 547 else 3), dt_fmt, tempo_str
    except Exception:
        return None, str(dt_val), "Data inválida"

todas_maquinas_totais = []
for s_nome in DICIONARIO_SETORES.keys():
    for m in obter_maquinas_setor(s_nome, df_correias, df_fusos):
        if m not in todas_maquinas_totais:
            todas_maquinas_totais.append(m)

ult_correias_dict = {}
if not df_correias.empty:
    for (s_idx, m_idx), sub_c in df_correias.groupby(["Setor", "Maquina_TAG"]):
        ult_correias_dict[(s_idx, m_idx)] = sub_c.iloc[-1]

for maq_tag in todas_maquinas_totais:
    setor_m = mapa_setor_maquina.get(maq_tag, "Setor A")
    ult = ult_correias_dict.get((setor_m, maq_tag))

    t1, d1_str, t1_uso, c1_score, tem_c1 = "Não informada", "Sem registro", "Sem histórico", None, False
    t2, d2_str, t2_uso, c2_score, tem_c2 = "Não informada", "Sem registro", "Sem histórico", None, False

    if ult is not None:
        v1, dt1_raw = formatar_modelo(ult.get("Tipo_Correia_1", "")), str(ult.get("Data_Instalacao_1", "")).strip()
        if v1 or (dt1_raw not in ["", "nan", "NaT", "None"]):
            t1 = v1 or "Não informada"
            c1_score, d1_str, t1_uso = avaliar_correia(dt1_raw)
            tem_c1 = True
            r1 = {"tag": maq_tag, "pos": "Superior", "modelo": t1, "data": d1_str, "uso": t1_uso, "setor": setor_m}
            lista_correias_todas.append(r1)
            (lista_correias_novas if c1_score == 1 else lista_correias_meia if c1_score == 2 else lista_correias_criticas).append(r1)

        v2, dt2_raw = formatar_modelo(ult.get("Tipo_Correia_2", "")), str(ult.get("Data_Instalacao_2", "")).strip()
        if v2 or (dt2_raw not in ["", "nan", "NaT", "None"]):
            t2 = v2 or "Não informada"
            c2_score, d2_str, t2_uso = avaliar_correia(dt2_raw)
            tem_c2 = True
            r2 = {"tag": maq_tag, "pos": "Inferior", "modelo": t2, "data": d2_str, "uso": t2_uso, "setor": setor_m}
            lista_correias_todas.append(r2)
            (lista_correias_novas if c2_score == 1 else lista_correias_meia if c2_score == 2 else lista_correias_criticas).append(r2)

    scores = [s for s in [c1_score, c2_score] if s is not None]
    if scores:
        pior = max(scores)
        classe_card = "status-verde" if pior == 1 else "status-amarelo" if pior == 2 else "status-vermelho"
        status_label = "Nova" if pior == 1 else "Meia-Vida" if pior == 2 else "Troca Necessária"
        dot_simbolo = "🟢" if pior == 1 else "🟡" if pior == 2 else "🔴"
    else:
        classe_card, status_label, dot_simbolo = "status-cinza", "Sem Dados", "⚪"

    dados_maquinas[maq_tag] = {
        "setor": setor_m, "t1": t1, "d1": d1_str, "uso1": t1_uso, "tem_c1": tem_c1,
        "t2": t2, "d2": d2_str, "uso2": t2_uso, "tem_c2": tem_c2,
        "status_label": status_label, "classe_card": classe_card, "dot": dot_simbolo
    }

st.markdown(
    """
    <style>
        .block-container { padding: 3.5rem 1.6rem 1rem 1.6rem !important; }
        [data-testid="stSidebar"] { background-color: #0f172a !important; border-right: 1px solid #1e293b !important; }
        [data-testid="stSidebar"] h2, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span { color: #f8fafc; }
        
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] {
            background-color: #1e293b !important; color: #e2e8f0 !important; border: 1px solid #334155 !important;
            border-radius: 8px !important; font-weight: 700 !important; height: 38px !important;
        }
        [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover { background-color: #334155 !important; }
        
        [data-testid="stSidebar"] .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: #ffffff !important;
            border: 1px solid #b91c1c !important; border-radius: 8px !important; font-weight: 800 !important; height: 38px !important;
        }

        div.stButton > button {
            background: #ffffff !important;
            color: #0f172a !important;
            border: 1px solid #cbd5e1 !important;
            padding: 0px 4px !important;
            font-size: 0.88rem !important;
            font-weight: 800 !important;
            height: 30px !important;
            min-height: 30px !important;
            line-height: 28px !important;
            border-radius: 6px !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04) !important;
            transition: all 0.1s ease-in-out !important;
        }
        div.stButton > button:hover {
            border-color: #2563eb !important;
            background: #f8fafc !important;
            color: #2563eb !important;
            transform: translateY(-1px);
        }
        
        div[data-testid="column"] { padding: 1px !important; margin: 0px !important; }
        div[data-testid="stHorizontalBlock"] { gap: 4px !important; margin-bottom: 3px !important; }
        div[data-testid="stVegaLiteChart"] summary, div[data-testid="stVegaLiteChart"] .vega-actions { display: none !important; }
        
        .card-kpi-bonito {
            background: #ffffff; border: 1px solid #e2e8f0; border-radius: 9px; padding: 7px 12px;
            display: flex; align-items: center; justify-content: space-between; height: 56px; box-sizing: border-box;
            position: relative; overflow: hidden;
        }
        .card-kpi-bonito::after { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; }
        .card-kpi-bonito.c-total::after { background: #475569; }
        .card-kpi-bonito.c-ok::after { background: #10b981; }
        .card-kpi-bonito.c-warn::after { background: #f59e0b; }
        .card-kpi-bonito.c-crit::after { background: #ef4444; }
        .kpi-val { font-size: 1.25rem; font-weight: 800; line-height: 1; font-family: ui-monospace, monospace; }
        .kpi-lbl { font-size: 0.65rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px; }
        
        .alerta-manutencao {
            background: #fef2f2; border: 1px solid #fecaca; border-left: 5px solid #ef4444; border-radius: 7px;
            padding: 5px 12px; margin: 4px 0 8px 0; font-size: 0.8rem; font-weight: 600; color: #991b1b;
        }
        .chip-critico { background: #fee2e2; border: 1px solid #fca5a5; color: #991b1b; padding: 1px 7px; border-radius: 4px; font-weight: 800; font-size: 0.76rem; }
        
        .hud-detalhe {
            background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 10px 14px; margin: 4px 0 8px 0;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05); border-left: 5px solid #64748b;
        }
        .hud-detalhe.status-verde { border-left-color: #10b981; }
        .hud-detalhe.status-amarelo { border-left-color: #f59e0b; }
        .hud-detalhe.status-vermelho { border-left-color: #ef4444; }
        .hud-detalhe.status-cinza { border-left-color: #94a3b8; }
        
        .tag-pill { background: #f8fafc; border: 1px solid #e2e8f0; padding: 3px 10px; border-radius: 5px; font-size: 0.78rem; font-weight: 700; color: #334155; }
        .badge-status { padding: 3px 8px; border-radius: 5px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase; }
        .badge-verde { background: #d1fae5; color: #065f46; }
        .badge-amarelo { background: #fef3c7; color: #92400e; }
        .badge-vermelho { background: #fee2e2; color: #991b1b; }
        .badge-cinza { background: #e2e8f0; color: #475569; }
        
        .pill-legenda { display: inline-flex; align-items: center; gap: 5px; font-size: 0.74rem; font-weight: 700; background: #f8fafc; border: 1px solid #e2e8f0; padding: 3px 8px; border-radius: 16px; }
        .dot-legenda { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }

        .header-setor-dash {
            font-size: 0.86rem;
            font-weight: 800;
            color: #0f172a;
            border-left: 3px solid #2563eb;
            padding-left: 7px;
            margin: 7px 0 3px 0;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# BARRA LATERAL (SIDEBAR) REESTRUTURADA
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='font-size:1.6rem; font-weight:900; margin:0 0 14px 0;'>⚙️ Portal PCM</h2>", unsafe_allow_html=True)
    
    st.markdown("<p style='font-size:0.72rem; font-weight:800; color:#94a3b8; margin:6px 0;'>PAINÉIS GERENCIAIS</p>", unsafe_allow_html=True)
    st.button("🔩 Painel de Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Fusos" else "secondary", on_click=navegar, args=("Painel Fusos",))
    st.button("🔄 Painel de Correias", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Correias" else "secondary", on_click=navegar, args=("Painel Correias",))
    st.button("🏭 Setores", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Setores" else "secondary", on_click=navegar, args=("Painel Setores",))
    st.button("⚙️ Máquinas", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Maquinas" else "secondary", on_click=navegar, args=("Painel Maquinas",))

    st.markdown("---")
    st.markdown("<p style='font-size:0.72rem; font-weight:800; color:#94a3b8; margin:6px 0;'>SISTEMA & DADOS</p>", unsafe_allow_html=True)
    st.button(
        "🗄️ Banco de Dados",
        key="btn_nav_banco_dados",
        use_container_width=True,
        type="primary" if st.session_state.pagina_atual == "Banco de Dados" else "secondary",
        on_click=navegar,
        args=("Banco de Dados",),
    )
    st.button(
        "🏭 Máquinas",
        key="btn_nav_cad_maquinas",
        use_container_width=True,
        type="primary" if st.session_state.pagina_atual == "Gestao Maquinas" else "secondary",
        on_click=navegar,
        args=("Gestao Maquinas",),
    )

    st.markdown("---")
    st.markdown("<div style='text-align:center; font-size:0.72rem; color:#64748b;'>PCM • Versão Gerencial</div>", unsafe_allow_html=True)

# ==========================================
# ÁREA PRINCIPAL
# ==========================================
tela = st.session_state.pagina_atual

# ------------------------------------------
# 1. PAINEL GERENCIAL DE CORREIAS
# ------------------------------------------
if tela == "Painel Correias":
    c_t, c_f, c_leg = st.columns([3.5, 2.0, 5.5])
    with c_t:
        st.markdown("<h3 style='margin:0; font-weight:900;'>Dashboard Correias</h3>", unsafe_allow_html=True)
    with c_f:
        mods_un = sorted(list({r["modelo"] for r in lista_correias_todas if r["modelo"] and r["modelo"] != "Não informada"}))
        filtro_modelo = st.selectbox("Tipo", ["Todos os Tipos"] + mods_un, label_visibility="collapsed")
    with c_leg:
        st.markdown("""
            <div style="height:36px; display:flex; align-items:center; justify-content:flex-end; gap:6px;">
                <span class='pill-legenda'><span class='dot-legenda' style='background:#10b981;'></span> Nova (&le; 1a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#f59e0b;'></span> Meia (1-1.5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#ef4444;'></span> Urgente (&gt; 1.5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#94a3b8;'></span> Sem Dados</span>
            </div>
        """, unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Correias</div><div class='kpi-val'>{len(lista_correias_todas)}</div></div><div>📦</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Novas</div><div class='kpi-val' style='color:#059669;'>{len(lista_correias_novas)}</div></div><div>🟢</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Meia-Vida</div><div class='kpi-val' style='color:#d97706;'>{len(lista_correias_meia)}</div></div><div>🟡</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Críticas</div><div class='kpi-val' style='color:#dc2626;'>{len(lista_correias_criticas)}</div></div><div>🔴</div></div>", unsafe_allow_html=True)

    if lista_correias_criticas:
        chips = " ".join([f"<span class='chip-critico'>🏷️ {k}: <b>{v} un.</b></span>" for k, v in pd.DataFrame(lista_correias_criticas)["modelo"].value_counts().items()])
        st.markdown(f"<div class='alerta-manutencao'>🚨 <b>Alerta de Troca Necessária:</b> &nbsp; {chips}</div>", unsafe_allow_html=True)

    if st.session_state.maq_clicada_cor:
        sel = st.session_state.maq_clicada_cor
        b_cor = "badge-verde" if sel["status_label"] == "Nova" else "badge-amarelo" if sel["status_label"] == "Meia-Vida" else "badge-vermelho" if sel["status_label"] == "Troca Necessária" else "badge-cinza"
        c_hud, c_cls = st.columns([6.2, 0.8])
        with c_hud:
            st.markdown(f"""
                <div class="hud-detalhe {sel['classe_card']}">
                    <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
                        <span class="tag-pill" style="background:#0f172a; color:#ffffff;">⚙️ {sel['tag']} - {sel['setor']}</span>
                        <span class="badge-status {b_cor}">{sel['status_label']}</span>
                    </div>
                    <div style="display:flex; gap:6px;">
                        <span class="tag-pill">🔼 <b>Superior / Cabeceira:</b> {sel['t1']} | {sel['d1']} | {sel['uso1']}</span>
                        <span class="tag-pill">🔽 <b>Inferior / Traseira:</b> {sel['t2']} | {sel['d2']} | {sel['uso2']}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with c_cls:
            if st.button("✖ Fechar", key="btn_cls_hud"):
                st.session_state.maq_clicada_cor = None
                st.rerun()

    for s_nome in DICIONARIO_SETORES.keys():
        maquinas_do_setor = obter_maquinas_setor(s_nome, df_correias, df_fusos)

        if filtro_modelo != "Todos os Tipos":
            filtradas = []
            for m in maquinas_do_setor:
                r_m = df_correias[(df_correias["Setor"] == s_nome) & (df_correias["Maquina_TAG"] == m)]
                if not r_m.empty:
                    ult_m = r_m.iloc[-1]
                    if formatar_modelo(ult_m.get("Tipo_Correia_1")) == filtro_modelo or formatar_modelo(ult_m.get("Tipo_Correia_2")) == filtro_modelo:
                        filtradas.append(m)
            maquinas_do_setor = filtradas

        if not maquinas_do_setor:
            continue

        st.markdown(f"<div class='header-setor-dash'><span>🏭 {s_nome}</span> <span style='font-size:0.75rem; color:#64748b;'>{len(maquinas_do_setor)} máquinas</span></div>", unsafe_allow_html=True)

        cols_g = 14
        for chunk in [maquinas_do_setor[i:i + cols_g] for i in range(0, len(maquinas_do_setor), cols_g)]:
            cols = st.columns(cols_g)
            for i, m in enumerate(chunk):
                info_m = dados_maquinas.get(m, {})
                dot_m = info_m.get("dot", "⚪")
                btn_label = f"{dot_m} {m}"

                if cols[i].button(btn_label, key=f"btn_c_{m.replace('-', '_')}", use_container_width=True):
                    st.session_state.maq_clicada_cor = {"tag": m, **info_m}
                    st.rerun()

# ------------------------------------------
# 2. PAINEL GERENCIAL DE FUSOS
# ------------------------------------------
elif tela == "Painel Fusos":
    cf_t, cf_a = st.columns([3.8, 1.4])
    cf_t.markdown("<h2 style='margin:0; font-weight:900;'>Dashboard Fusos</h2>", unsafe_allow_html=True)
    ano_f = cf_a.selectbox("Ano", [2024, 2025, 2026, 2027], index=2, label_visibility="collapsed")

    data_hoje_ref = date.today()
    ano_atual_ref = data_hoje_ref.year
    mes_atual_num_ref = data_hoje_ref.month

    if int(ano_f) < ano_atual_ref:
        div_meses = 12
        desc_divisor = "12 meses"
    elif int(ano_f) == ano_atual_ref:
        div_meses = max(1, mes_atual_num_ref)
        desc_divisor = f"Jan a {ORDEM_MESES_ABREV[div_meses - 1]}"
    else:
        div_meses = 1
        desc_divisor = "Previsto"

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    b_geral, b_sa, b_sb, b_lat, b_men = st.columns(5)
    for b_col, nome_aba in zip([b_geral, b_sa, b_sb, b_lat, b_men], ["Geral", "Setor A", "Setor B", "Setor Látex", "Setor Menegatto"]):
        if b_col.button(nome_aba, use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == nome_aba else "secondary"):
            st.session_state.aba_setor_fuso = nome_aba
            st.rerun()

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    df_ano_f = df_fusos[df_fusos["Ano"] == int(ano_f)].copy()
    if df_ano_f.empty:
        df_ano_f = pd.DataFrame(columns=COLUNAS_FUSOS)

    if st.session_state.aba_setor_fuso == "Geral":
        tot_fabrica = int(df_ano_f["Quantidade_Quebras"].sum())
        med_fabrica = round(tot_fabrica / div_meses, 1)

        ult_mes_fab = "Nenhum"
        tot_ult_mes = 0
        setor_ofensor = "Nenhum"
        qtd_setor_ofensor = 0

        if not df_ano_f.empty and tot_fabrica > 0:
            df_reais = df_ano_f[df_ano_f["Quantidade_Quebras"] > 0]
            for m_teste in reversed(LISTA_MESES_PUROS):
                sub_m = df_reais[df_reais["Mes"] == m_teste]
                if not sub_m.empty and sub_m["Quantidade_Quebras"].sum() > 0:
                    ult_mes_fab = m_teste
                    tot_ult_mes = int(sub_m["Quantidade_Quebras"].sum())
                    agrup_s = sub_m.groupby("Setor")["Quantidade_Quebras"].sum().sort_values(ascending=False)
                    setor_ofensor = agrup_s.index[0]
                    qtd_setor_ofensor = int(agrup_s.iloc[0])
                    break

        kf1, kf2, kf3, kf4 = st.columns(4)
        kf1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Fábrica</div><div class='kpi-val'>{tot_fabrica}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
        kf2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal ({desc_divisor})</div><div class='kpi-val' style='color:#059669;'>{med_fabrica}</div></div><div>📈</div></div>", unsafe_allow_html=True)
        kf3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Setor Crítico ({ult_mes_fab})</div><div class='kpi-val' style='color:#d97706; font-size:1.1rem;'>{setor_ofensor} ({qtd_setor_ofensor})</div></div><div>🏭</div></div>", unsafe_allow_html=True)
        kf4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Quebras no Mês ({ult_mes_fab})</div><div class='kpi-val' style='color:#dc2626;'>{tot_ult_mes}</div></div><div>🚨</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        def gerar_chart_setor(nome_s, cor_s):
            sub_s = df_ano_f[df_ano_f["Setor"] == nome_s].groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
            sub_s["Mes_Abrev"] = sub_s["Mes"].map(MAPA_MES_ABREV)
            tot_s = int(sub_s["Quantidade_Quebras"].sum())

            bars = alt.Chart(sub_s).mark_bar(color=cor_s, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                y=alt.Y("Quantidade_Quebras:Q", title="Quebras")
            )
            txt = alt.Chart(sub_s).mark_text(dy=-6, fontSize=11, fontWeight=700).encode(
                x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV),
                y=alt.Y("Quantidade_Quebras:Q"),
                text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
            )
            return (bars + txt).properties(height=210), tot_s

        cg1, cg2 = st.columns(2)
        cores_map = {"Setor A": "#2563eb", "Setor B": "#d97706", "Setor Látex": "#059669", "Setor Menegatto": "#dc2626"}

        with cg1:
            with st.container(border=True):
                c_a, t_a = gerar_chart_setor("Setor A", cores_map["Setor A"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:0.95rem; margin-bottom:4px;'><span>🏭 Setor A</span><span style='color:{cores_map['Setor A']}'>Total: {t_a} fusos</span></div>", unsafe_allow_html=True)
                st.altair_chart(c_a, use_container_width=True)

            with st.container(border=True):
                c_lat, t_lat = gerar_chart_setor("Setor Látex", cores_map["Setor Látex"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:0.95rem; margin-bottom:4px;'><span>🌿 Setor Látex</span><span style='color:{cores_map['Setor Látex']}'>Total: {t_lat} fusos</span></div>", unsafe_allow_html=True)
                st.altair_chart(c_lat, use_container_width=True)

        with cg2:
            with st.container(border=True):
                c_b, t_b = gerar_chart_setor("Setor B", cores_map["Setor B"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:0.95rem; margin-bottom:4px;'><span>🏭 Setor B</span><span style='color:{cores_map['Setor B']}'>Total: {t_b} fusos</span></div>", unsafe_allow_html=True)
                st.altair_chart(c_b, use_container_width=True)

            with st.container(border=True):
                c_men, t_men = gerar_chart_setor("Setor Menegatto", cores_map["Setor Menegatto"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:0.95rem; margin-bottom:4px;'><span>⚙️ Setor Menegatto</span><span style='color:{cores_map['Setor Menegatto']}'>Total: {t_men} fusos</span></div>", unsafe_allow_html=True)
                st.altair_chart(c_men, use_container_width=True)

    else:
        s_ativo = st.session_state.aba_setor_fuso
        df_sa = df_ano_f[df_ano_f["Setor"] == s_ativo].copy()

        tot_s = int(df_sa["Quantidade_Quebras"].sum()) if not df_sa.empty else 0
        maqs_setor = obter_maquinas_setor(s_ativo, df_correias, df_fusos)
        qtd_maqs_s = len(maqs_setor)
        med_s = round(tot_s / div_meses, 1)

        ult_mes_s = "Nenhum"
        top_maq_s = "Nenhuma"
        qtd_top_s = 0
        quebras_ult_mes = 0

        if not df_sa.empty and tot_s > 0:
            df_reais_s = df_sa[df_sa["Quantidade_Quebras"] > 0]
            for m_teste in reversed(LISTA_MESES_PUROS):
                sub_m = df_reais_s[df_reais_s["Mes"] == m_teste]
                if not sub_m.empty and sub_m["Quantidade_Quebras"].sum() > 0:
                    ult_mes_s = m_teste
                    quebras_ult_mes = int(sub_m["Quantidade_Quebras"].sum())
                    agrup_m = sub_m.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().sort_values(ascending=False)
                    top_maq_s = agrup_m.index[0]
                    qtd_top_s = int(agrup_m.iloc[0])
                    break

        quebras_por_maq = round(quebras_ult_mes / qtd_maqs_s, 1) if (qtd_maqs_s > 0 and quebras_ult_mes > 0) else 0.0

        ks1, ks2, ks3, ks4 = st.columns(4)
        ks1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Quebras ({s_ativo})</div><div class='kpi-val'>{tot_s}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
        ks2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal ({desc_divisor})</div><div class='kpi-val' style='color:#059669;'>{med_s}</div></div><div>📅</div></div>", unsafe_allow_html=True)
        ks3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Quebras / Máq ({ult_mes_s})</div><div class='kpi-val' style='color:#d97706;'>{quebras_por_maq}</div></div><div>⚙️</div></div>", unsafe_allow_html=True)
        ks4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Maior Quebra ({ult_mes_s})</div><div class='kpi-val' style='color:#dc2626; font-size:1.1rem;'>{top_maq_s} ({qtd_top_s})</div></div><div>⚠️</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        c_evol, c_rosca = st.columns([1.55, 1.45])

        with c_evol:
            with st.container(border=True):
                st.markdown(f"<div style='font-size:0.96rem; font-weight:800; margin-bottom:6px;'>📊 Evolução Mensal de Quebras ({s_ativo} - {ano_f})</div>", unsafe_allow_html=True)
                df_evol = df_sa.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                df_evol["Mes_Abrev"] = df_evol["Mes"].map(MAPA_MES_ABREV)

                barras_s = alt.Chart(df_evol).mark_bar(color="#2563eb", cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title="Mês", axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                    tooltip=[alt.Tooltip("Mes:N", title="Mês"), alt.Tooltip("Quantidade_Quebras:Q", title="Quebras")]
                )
                rotulos_s = alt.Chart(df_evol).mark_text(dy=-8, fontSize=11, fontWeight=700).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV),
                    y=alt.Y("Quantidade_Quebras:Q"),
                    text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                )
                st.altair_chart((barras_s + rotulos_s).properties(height=280), use_container_width=True)

        with c_rosca:
            with st.container(border=True):
                st.markdown("<div style='font-size:0.96rem; font-weight:800; margin-bottom:6px;'>🍩 Distribuição por Tipo de Fuso</div>", unsafe_allow_html=True)
                if not df_sa.empty and tot_s > 0:
                    df_tipos = df_sa[df_sa["Quantidade_Quebras"] > 0].groupby("Tipo_Fuso")["Quantidade_Quebras"].sum().reset_index()
                    chart_donut = alt.Chart(df_tipos).mark_arc(innerRadius=60, outerRadius=110, stroke="#ffffff", strokeWidth=2).encode(
                        theta=alt.Theta("Quantidade_Quebras:Q", stack=True),
                        color=alt.Color("Tipo_Fuso:N", title="Tipo / Marca", scale=alt.Scale(scheme="category10"), legend=alt.Legend(orient="right")),
                        tooltip=[alt.Tooltip("Tipo_Fuso:N", title="Tipo"), alt.Tooltip("Quantidade_Quebras:Q", title="Quebras")]
                    ).properties(height=280)
                    st.altair_chart(chart_donut, use_container_width=True)
                else:
                    st.info(f"Sem registros de tipos de fuso para {s_ativo} em {ano_f}.")

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.96rem; font-weight:800; margin-bottom:6px;'>🔥 Mapa de Calor Operacional — Quebras por Máquina x Mês ({s_ativo} - {ano_f})</div>", unsafe_allow_html=True)
            
            idx_grid = pd.MultiIndex.from_product([maqs_setor, LISTA_MESES_PUROS], names=["MAQ", "Mes"]).to_frame().reset_index(drop=True)
            idx_grid["MES"] = idx_grid["Mes"].map(MAPA_MES_ABREV)
            agrup_c = df_sa.groupby(["Maquina_TAG", "Mes"])["Quantidade_Quebras"].sum().reset_index()
            m_calor = pd.merge(idx_grid, agrup_c, left_on=["MAQ", "Mes"], right_on=["Maquina_TAG", "Mes"], how="left").fillna(0)

            rect = alt.Chart(m_calor).mark_rect(stroke="#fff", strokeWidth=1).encode(
                x=alt.X("MES:N", sort=ORDEM_MESES_ABREV, title="Mês", axis=alt.Axis(orient="top", labelAngle=0, labelFontWeight="bold")),
                y=alt.Y("MAQ:N", sort=maqs_setor, title="Máquina", axis=alt.Axis(labelFontWeight="bold")),
                color=alt.Color("Quantidade_Quebras:Q", scale=alt.Scale(domain=[0, 3, 8, 15], range=["#dcfce7", "#fef08a", "#f97316", "#dc2626"]), legend=alt.Legend(title="Quebras")),
                tooltip=[alt.Tooltip("MAQ:N", title="Máquina"), alt.Tooltip("MES:N", title="Mês"), alt.Tooltip("Quantidade_Quebras:Q", title="Quebras")]
            )
            txt = alt.Chart(m_calor).mark_text(baseline="middle", fontSize=11, fontWeight=700).encode(
                x=alt.X("MES:N", sort=ORDEM_MESES_ABREV),
                y=alt.Y("MAQ:N", sort=maqs_setor),
                text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")),
                color=alt.condition("datum.Quantidade_Quebras >= 10", alt.value("#ffffff"), alt.value("#0f172a")),
            )
            st.altair_chart((rect + txt).properties(height=max(320, len(maqs_setor) * 23)), use_container_width=True)

# ------------------------------------------
# 3. PAINEL GERENCIAL DE SETORES (COMPLETO COM TODAS AS NOVAS SOLICITAÇÕES)
# ------------------------------------------
elif tela == "Painel Setores":
    c_ts1, c_ts2, c_ts3 = st.columns([2.5, 1.8, 1.8])
    with c_ts1:
        st.markdown("<h2 style='margin:0; font-weight:900;'>🏭 Painel Executivo de Setores</h2>", unsafe_allow_html=True)
    with c_ts2:
        setores_filtro_painel = list(DICIONARIO_SETORES.keys())
        setor_selecionado_exec = st.selectbox("Filtrar Setor:", setores_filtro_painel, label_visibility="collapsed")
    with c_ts3:
        meses_filtro_painel = ["Todos os Meses"] + LISTA_MESES_PUROS
        mes_selecionado_exec = st.selectbox("Filtrar Mês:", meses_filtro_painel, label_visibility="collapsed")

    st.caption(f"Análise executiva detalhada para o **{setor_selecionado_exec}** ({mes_selecionado_exec}).")

    # Coleta de dados exclusiva do setor selecionado
    s_nome = setor_selecionado_exec
    maqs_s = obter_maquinas_setor(s_nome, df_correias, df_fusos)
    qtd_maqs_setor = len(maqs_s)

    # 1. Gráfico e dados de Fusos do Setor
    df_fusos_setor = df_fusos[df_fusos["Setor"] == s_nome].copy() if not df_fusos.empty else pd.DataFrame()
    if mes_selecionado_exec != "Todos os Meses" and not df_fusos_setor.empty:
        df_fusos_setor = df_fusos_setor[df_fusos_setor["Mes"] == mes_selecionado_exec]
    tot_q_fusos = int(df_fusos_setor["Quantidade_Quebras"].sum()) if not df_fusos_setor.empty else 0

    # 2. Horas Paradas por Corretivas no Setor
    df_paradas_setor = df_paradas[df_paradas["Setor"] == s_nome].copy() if not df_paradas.empty else pd.DataFrame()
    if mes_selecionado_exec != "Todos os Meses" and not df_paradas_setor.empty:
        df_paradas_setor["Mes_Nome"] = pd.to_datetime(df_paradas_setor["Data"], errors="coerce").dt.month.map(lambda x: LISTA_MESES_PUROS[x-1] if pd.notna(x) and 1 <= x <= 12 else "")
        df_paradas_setor = df_paradas_setor[df_paradas_setor["Mes_Nome"] == mes_selecionado_exec]
    tot_horas_paradas_setor = float(df_paradas_setor["Tempo_Parado_Horas"].sum()) if not df_paradas_setor.empty else 0.0

    # 3. Cálculo de Rendimento da Manutenção (%)
    # Possível de produzir: qtd de máquinas * 24 horas (se mês específico, multiplica pelos dias do mês; senão, base anual padrão de 30 dias/mês ou 720h por máquina)
    horas_possiveis_total = qtd_maqs_setor * 720.0 if mes_selecionado_exec == "Todos os Meses" else qtd_maqs_setor * 24.0 * calendar.monthrange(2026, LISTA_MESES_PUROS.index(mes_selecionado_exec)+1)[1]
    rendimento_manutencao = max(0.0, round(100.0 * (1.0 - (tot_horas_paradas_setor / horas_possiveis_total)), 1)) if horas_possiveis_total > 0 else 100.0

    # 4. Condições das Correias do Setor
    novas_s = len([r for r in lista_correias_novas if r["setor"] == s_nome])
    meia_s = len([r for r in lista_correias_meia if r["setor"] == s_nome])
    crit_s = len([r for r in lista_correias_criticas if r["setor"] == s_nome])
    total_cor_s = novas_s + meia_s + crit_s

    # 5. Top 5 Máquinas com mais Corretivas no Setor
    if not df_paradas_setor.empty:
        top_corretivas = df_paradas_setor.groupby("Maquina_TAG")["Tempo_Parado_Horas"].sum().reset_index()
        top_corretivas.columns = ["Máquina", "Horas Paradas"]
        top_corretivas = top_corretivas.sort_values(by="Horas Paradas", ascending=False).head(5)
    else:
        top_corretivas = pd.DataFrame(columns=["Máquina", "Horas Paradas"])

    # 6. Máquinas que precisam trocar correia no Setor
    maqs_troca_correia = [r["tag"] for r in lista_correias_criticas if r["setor"] == s_nome]

    # KPIs Superiores para o Setor Selecionado
    cs1, cs2, cs3, cs4 = st.columns(4)
    cs1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Rendimento Manutenção</div><div class='kpi-val' style='color:#059669;'>{rendimento_manutencao}%</div></div><div>📈</div></div>", unsafe_allow_html=True)
    cs2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Total Horas Paradas</div><div class='kpi-val' style='color:#dc2626;'>{round(tot_horas_paradas_setor, 1)}h</div></div><div>⏱️</div></div>", unsafe_allow_html=True)
    cs3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Quebras de Fusos</div><div class='kpi-val' style='color:#d97706;'>{tot_q_fusos}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
    cs4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Trocas de Correia Pendentes</div><div class='kpi-val' style='color:#dc2626;'>{len(maqs_troca_correia)}</div></div><div>🚨</div></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    # Gráficos do Setor Selecionado: Quebras de Fusos (Evolução) + Rosca de Correias
    cg_set1, cg_set2 = st.columns(2)

    with cg_set1:
        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🔩 Evolução de Quebras de Fusos — {s_nome}</div>", unsafe_allow_html=True)
            if not df_fusos_setor.empty:
                df_f_graf = df_fusos_setor.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                df_f_graf["Mes_Abrev"] = df_f_graf["Mes"].map(MAPA_MES_ABREV)

                bar_fs = alt.Chart(df_f_graf).mark_bar(color="#2563eb", cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                    tooltip=["Mes", "Quantidade_Quebras"]
                )
                txt_fs = bar_fs.mark_text(dy=-8, fontSize=11, fontWeight=700).encode(
                    text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                )
                st.altair_chart((bar_fs + txt_fs).properties(height=240), use_container_width=True)
            else:
                st.info(f"Sem registros de quebras de fusos para o {s_nome}.")

    with cg_set2:
        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🍩 Condições das Correias — {s_nome}</div>", unsafe_allow_html=True)
            if total_cor_s > 0:
                df_donut_s = pd.DataFrame([
                    {"Condicao": "Novas (≤ 1a)", "Quantidade": novas_s, "Cor": "#10b981"},
                    {"Condicao": "Meia-Vida (1-1.5a)", "Quantidade": meia_s, "Cor": "#f59e0b"},
                    {"Condicao": "Troca Urgente (> 1.5a)", "Quantidade": crit_s, "Cor": "#ef4444"},
                ])
                chart_don_s = alt.Chart(df_donut_s).mark_arc(innerRadius=65, outerRadius=110, stroke="#ffffff", strokeWidth=2).encode(
                    theta=alt.Theta("Quantidade:Q", stack=True),
                    color=alt.Color("Condicao:N", scale=alt.Scale(domain=["Novas (≤ 1a)", "Meia-Vida (1-1.5a)", "Troca Urgente (> 1.5a)"], range=["#10b981", "#f59e0b", "#ef4444"]), legend=alt.Legend(orient="right")),
                    tooltip=["Condicao", "Quantidade"]
                ).properties(height=240)
                st.altair_chart(chart_don_s, use_container_width=True)
            else:
                st.info(f"Sem dados de correias cadastradas para o {s_nome}.")

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # Tabelas Informativas do Setor: Top 5 Corretivas e Máquinas para Trocar Correia
    ct_set1, ct_set2 = st.columns(2)

    with ct_set1:
        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🏆 Top Máquinas com Mais Manutenções Corretivas</div>", unsafe_allow_html=True)
            if not top_corretivas.empty:
                st.dataframe(top_corretivas, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhuma corretiva registrada neste setor.")

    with ct_set2:
        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🚨 Máquinas que Precisam Trocar Correia</div>", unsafe_allow_html=True)
            if maqs_troca_correia:
                df_troca_c = pd.DataFrame([{"Máquina (TAG)": t, "Setor": s_nome, "Status": "Troca Urgente (> 1.5 anos)"} for t in maqs_troca_correia])
                st.dataframe(df_troca_c, use_container_width=True, hide_index=True)
            else:
                st.success("Nenhuma máquina com correia crítica neste setor no momento!")

# ------------------------------------------
# 4. PAINEL GERENCIAL DE MÁQUINAS
# ------------------------------------------
elif tela == "Painel Maquinas":
    st.markdown("<h2 style='margin:0; font-weight:900;'>⚙️ Prontuário Individual da Máquina</h2>", unsafe_allow_html=True)
    st.caption("Consulte o histórico detalhado, dados de correias, manutenções corretivas, pendências e quebras de fusos por TAG.")

    col_sm, col_mq = st.columns([1.5, 2.0])
    with col_sm:
        setor_selecionado_maq = st.selectbox("Filtrar por Setor:", list(DICIONARIO_SETORES.keys()), key="sel_maq_painel_set")
    
    maqs_disponiveis = obter_maquinas_setor(setor_selecionado_maq, df_correias, df_fusos)
    with col_mq:
        tag_selecionada = st.selectbox("Selecione a TAG da Máquina:", maqs_disponiveis, key="sel_maq_painel_tag")

    if tag_selecionada:
        info_cor_maq = dados_maquinas.get(tag_selecionada, {})
        sub_fusos_maq = df_fusos[df_fusos["Maquina_TAG"] == tag_selecionada]
        tot_falhas_maq = int(sub_fusos_maq["Quantidade_Quebras"].sum()) if not sub_fusos_maq.empty else 0

        sub_paradas_maq = df_paradas[df_paradas["Maquina_TAG"] == tag_selecionada]
        tot_horas_paradas = float(sub_paradas_maq["Tempo_Parado_Horas"].sum()) if not sub_paradas_maq.empty else 0.0

        sub_pend_maq = df_pendencias[(df_pendencias["Maquina_TAG"] == tag_selecionada) & (df_pendencias["Status"].astype(str).str.lower() != "concluído")]
        tot_pendencias_abertas = len(sub_pend_maq)

        cm1, cm2, cm3, cm4 = st.columns(4)
        cm1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Setor Ativo</div><div class='kpi-val' style='font-size:1.05rem;'>{setor_selecionado_maq}</div></div><div>🏭</div></div>", unsafe_allow_html=True)
        cm2.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Status Correia</div><div class='kpi-val' style='font-size:1.05rem;'>{info_cor_maq.get('dot', '⚪')} {info_cor_maq.get('status_label', 'Sem Dados')}</div></div><div>🔄</div></div>", unsafe_allow_html=True)
        cm3.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Horas Paradas (Corretivas)</div><div class='kpi-val' style='color:#dc2626;'>{round(tot_horas_paradas, 1)}h</div></div><div>⏱️</div></div>", unsafe_allow_html=True)
        cm4.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Pendências Abertas</div><div class='kpi-val' style='color:#059669;'>{tot_pendencias_abertas}</div></div><div>📋</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🔄 Correias Instaladas na Máquina {tag_selecionada}</div>", unsafe_allow_html=True)
            c_cor1, c_cor2 = st.columns(2)
            with c_cor1:
                st.markdown(f"""
                    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:10px 14px;">
                        <div style="font-weight:800; font-size:0.85rem; color:#0f172a; margin-bottom:4px;">🔼 Superior / Cabeceira</div>
                        <div style="font-size:0.8rem; color:#475569;"><b>Modelo:</b> {info_cor_maq.get('t1', 'Não informada')}</div>
                        <div style="font-size:0.8rem; color:#475569;"><b>Instalação:</b> {info_cor_maq.get('d1', 'Sem registro')}</div>
                        <div style="font-size:0.8rem; color:#475569;"><b>Tempo de Uso:</b> {info_cor_maq.get('uso1', 'Sem histórico')}</div>
                    </div>
                """, unsafe_allow_html=True)
            with c_cor2:
                st.markdown(f"""
                    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:10px 14px;">
                        <div style="font-weight:800; font-size:0.85rem; color:#0f172a; margin-bottom:4px;">🔽 Inferior / Traseira</div>
                        <div style="font-size:0.8rem; color:#475569;"><b>Modelo:</b> {info_cor_maq.get('t2', 'Não informada')}</div>
                        <div style="font-size:0.8rem; color:#475569;"><b>Instalação:</b> {info_cor_maq.get('d2', 'Sem registro')}</div>
                        <div style="font-size:0.8rem; color:#475569;"><b>Tempo de Uso:</b> {info_cor_maq.get('uso2', 'Sem histórico')}</div>
                    </div>
                """, unsafe_allow_html=True)

        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>📋 Manutenções Pendentes — {tag_selecionada}</div>", unsafe_allow_html=True)
            if not sub_pend_maq.empty:
                df_exibe_pend = sub_pend_maq[["Descricao_Pendencia", "Prioridade", "Status"]]
                st.dataframe(df_exibe_pend, use_container_width=True, hide_index=True)
            else:
                st.info(f"Nenhuma manutenção pendente cadastrada para a máquina {tag_selecionada}.")

        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🛠️ Manutenções Corretivas Executadas — {tag_selecionada}</div>", unsafe_allow_html=True)
            if not sub_paradas_maq.empty:
                df_exibe_p = sub_paradas_maq[["Data", "Tipo_Manutencao", "Descricao_Servico", "Tempo_Parado_Horas"]].sort_values("Data", ascending=False)
                st.dataframe(df_exibe_p, use_container_width=True, hide_index=True)
            else:
                st.info(f"Nenhuma manutenção corretiva registrada para a máquina {tag_selecionada}.")

        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown(f"<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>📊 Evolução de Quebras de Fusos — {tag_selecionada}</div>", unsafe_allow_html=True)
            if not sub_fusos_maq.empty:
                df_f_maq = sub_fusos_maq.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                df_f_maq["Mes_Abrev"] = df_f_maq["Mes"].map(MAPA_MES_ABREV)

                bar_maq = alt.Chart(df_f_maq).mark_bar(color="#2563eb", cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title="Mês", axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                    tooltip=["Mes", "Quantidade_Quebras"]
                )
                txt_maq = bar_maq.mark_text(dy=-6, fontSize=11, fontWeight=700).encode(
                    text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                )
                st.altair_chart((bar_maq + txt_maq).properties(height=240), use_container_width=True)
            else:
                st.info(f"Sem registros de quebras de fusos para a máquina {tag_selecionada}.")

# ------------------------------------------
# 5. BANCO DE DADOS & GESTÃO DE ARQUIVOS
# ------------------------------------------
elif tela == "Banco de Dados":
    st.title("🗄️ Banco de Dados & Gestão de Arquivos")
    st.caption("Central de importação e exportação de dados mestres em formato Excel (.xlsx)")

    tab_fusos_db, tab_correias_db, tab_paradas_db, tab_pendencias_db, tab_backups_db = st.tabs([
        "🔩 Base de Fusos",
        "🔄 Base de Correias",
        "🛠️ Corretivas & Paradas",
        "📋 Manutenções Pendentes",
        "🛡️ Histórico de Backups"
    ])

    with tab_fusos_db:
        st.markdown("### 📅 Gestão de Fusos Anual Consolidada")
        c_ano_db, c_set_db = st.columns([1.5, 2.5])
        with c_ano_db:
            ano_db_fuso = st.selectbox("Ano de Trabalho:", [2024, 2025, 2026, 2027], index=2, key="sel_ano_db_fuso")
        with c_set_db:
            setor_db_fuso = st.selectbox("Setor:", list(DICIONARIO_SETORES.keys()), key="sel_setor_db_fuso")

        maquinas_set_db = obter_maquinas_setor(setor_db_fuso, df_correias, df_fusos)

        buffer_excel_ano = io.BytesIO()
        with pd.ExcelWriter(buffer_excel_ano, engine="openpyxl") as writer:
            for idx_m, nome_mes_aba in enumerate(LISTA_MESES_PUROS):
                num_mes = idx_m + 1
                _, dias_no_mes = calendar.monthrange(int(ano_db_fuso), num_mes)
                cols_dias_aba = [str(d) for d in range(1, dias_no_mes + 1)]

                df_mes_fuso = df_fusos[
                    (df_fusos["Ano"] == int(ano_db_fuso))
                    & (df_fusos["Mes"] == nome_mes_aba)
                    & (df_fusos["Setor"] == setor_db_fuso)
                ]

                grade_aba = []
                for maq in maquinas_set_db:
                    sub_maq = df_mes_fuso[df_mes_fuso["Maquina_TAG"] == maq]
                    linha_aba = {"MAQUINA": maq}
                    for d in range(1, dias_no_mes + 1):
                        sub_d = sub_maq[sub_maq["Dia"] == d]
                        linha_aba[str(d)] = int(sub_d.iloc[0]["Quantidade_Quebras"]) if not sub_d.empty else 0
                    grade_aba.append(linha_aba)

                df_mes_planilha = pd.DataFrame(grade_aba)[["MAQUINA"] + cols_dias_aba]
                df_mes_planilha.to_excel(writer, index=False, sheet_name=nome_mes_aba[:31])

        buffer_excel_ano.seek(0)

        st.markdown("---")
        col_down_ano, col_up_ano = st.columns([1.5, 2.5])

        with col_down_ano:
            st.markdown("#### 📥 Descarregar Livro de 12 Meses")
            st.download_button(
                f"📥 Baixar Ano {ano_db_fuso} Completo (.xlsx)",
                data=buffer_excel_ano,
                file_name=f"fusos_{setor_db_fuso.replace(' ', '_')}_{ano_db_fuso}_12_meses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"btn_down_ano_completo_{ano_db_fuso}_{setor_db_fuso}",
                use_container_width=True,
            )

        with col_up_ano:
            st.markdown("#### 📤 Importar Livro de 12 Meses")
            upload_ano = st.file_uploader(
                f"Enviar ficheiro completo de {ano_db_fuso} (.xlsx)",
                type=["xlsx"],
                key=f"upload_ano_db_{ano_db_fuso}_{setor_db_fuso}",
            )
            if upload_ano is not None:
                if st.button(f"Confirmar e Atualizar Ano {ano_db_fuso}", key=f"btn_conf_up_ano_{ano_db_fuso}", type="primary"):
                    try:
                        excel_importado = pd.ExcelFile(upload_ano)
                        abas_encontradas = excel_importado.sheet_names
                        novos_registros_ano = []
                        meses_atualizados = []

                        fuso_padrao = "FAG" if setor_db_fuso == "Setor A" else "TEP" if setor_db_fuso == "Setor B" else "M4BA" if setor_db_fuso == "Setor Látex" else "MENEGATTO"

                        for nome_mes_oficial in LISTA_MESES_PUROS:
                            aba_alvo = None
                            for sh in abas_encontradas:
                                if sh.strip().lower() == nome_mes_oficial.lower():
                                    aba_alvo = sh
                                    break

                            if aba_alvo:
                                df_aba = pd.read_excel(excel_importado, sheet_name=aba_alvo)
                                col_maq = None
                                for cand in ["MAQUINA", "Máquina", "Maquina", "Maquina_TAG"]:
                                    if cand in df_aba.columns:
                                        col_maq = cand
                                        break

                                if col_maq:
                                    num_mes = LISTA_MESES_PUROS.index(nome_mes_oficial) + 1
                                    _, dias_no_mes = calendar.monthrange(int(ano_db_fuso), num_mes)
                                    meses_atualizados.append(nome_mes_oficial)

                                    for _, r_up in df_aba.iterrows():
                                        m_val = str(r_up[col_maq]).strip()
                                        for d in range(1, dias_no_mes + 1):
                                            qtd_val = 0
                                            if d in df_aba.columns:
                                                qtd_val = int(r_up[d]) if pd.notna(r_up[d]) else 0
                                            elif str(d) in df_aba.columns:
                                                qtd_val = int(r_up[str(d)]) if pd.notna(r_up[str(d)]) else 0
                                            elif f"{d:02d}" in df_aba.columns:
                                                qtd_val = int(r_up[f"{d:02d}"]) if pd.notna(r_up[f"{d:02d}"]) else 0

                                            novos_registros_ano.append({
                                                "Ano": int(ano_db_fuso),
                                                "Mes": nome_mes_oficial,
                                                "Dia": d,
                                                "Setor": setor_db_fuso,
                                                "Maquina_TAG": m_val,
                                                "Quantidade_Quebras": qtd_val,
                                                "Tipo_Fuso": fuso_padrao,
                                            })

                        if novos_registros_ano:
                            df_limpo_fusos = df_fusos[
                                ~(
                                    (df_fusos["Ano"] == int(ano_db_fuso))
                                    & (df_fusos["Setor"] == setor_db_fuso)
                                    & (df_fusos["Mes"].isin(meses_atualizados))
                                )
                            ]
                            df_atualizado_fusos = pd.concat([df_limpo_fusos, pd.DataFrame(novos_registros_ano)], ignore_index=True)
                            gerar_backup_seguro(ARQUIVO_FUSOS)
                            df_atualizado_fusos.to_excel(ARQUIVO_FUSOS, index=False)
                            invalidar_cache()
                            st.success(f"✅ {len(meses_atualizados)} meses atualizados com sucesso!")
                            st.rerun()
                        else:
                            st.warning("Nenhuma aba com nome de mês correspondente ou coluna 'MAQUINA' foi encontrada.")

                    except Exception as erro_up:
                        st.error(f"Erro ao processar o ficheiro anual: {erro_up}")

    with tab_correias_db:
        st.markdown("### 🔄 Troca de Dados de Correias")
        col_d_cor, col_u_cor = st.columns([1.5, 2.5])

        with col_d_cor:
            st.markdown("#### 📥 Descarregar Planilha")
            filtro_export_setor = st.selectbox(
                "Exportar Setor:",
                ["Todos os Setores"] + list(DICIONARIO_SETORES.keys()),
                key="sel_export_setor_cor"
            )

            setores_alvo = list(DICIONARIO_SETORES.keys()) if filtro_export_setor == "Todos os Setores" else [filtro_export_setor]
            linhas_export = []

            for s_nome in setores_alvo:
                maqs_set = obter_maquinas_setor(s_nome, df_correias, df_fusos)
                for maq in maqs_set:
                    reg = df_correias[(df_correias["Setor"] == s_nome) & (df_correias["Maquina_TAG"] == maq)]
                    t1, d1, t2, d2 = "", "", "", ""
                    if not reg.empty:
                        ult = reg.iloc[-1]
                        t1 = formatar_modelo(ult.get("Tipo_Correia_1", ""))
                        d1 = str(ult.get("Data_Instalacao_1", "")).replace("nan", "").replace("NaT", "").strip()
                        t2 = formatar_modelo(ult.get("Tipo_Correia_2", ""))
                        d2 = str(ult.get("Data_Instalacao_2", "")).replace("nan", "").replace("NaT", "").strip()

                    linhas_export.append({
                        "Setor": s_nome,
                        "Maquina_TAG": maq,
                        "Tipo_Correia_1": t1,
                        "Data_Instalacao_1": d1,
                        "Tipo_Correia_2": t2,
                        "Data_Instalacao_2": d2,
                    })

            df_export_pronto = pd.DataFrame(linhas_export)
            buf_down_cor = io.BytesIO()
            with pd.ExcelWriter(buf_down_cor, engine="openpyxl") as writer_cor:
                df_export_pronto.to_excel(writer_cor, index=False, sheet_name="Correias")
            buf_down_cor.seek(0)

            nome_arquivo_down = (
                f"correias_fabrica_completa_{date.today().strftime('%Y%m%d')}.xlsx"
                if filtro_export_setor == "Todos os Setores"
                else f"correias_{filtro_export_setor.replace(' ', '_')}_{date.today().strftime('%Y%m%d')}.xlsx"
            )

            st.download_button(
                label="📥 Baixar Planilha de Correias (.xlsx)",
                data=buf_down_cor,
                file_name=nome_arquivo_down,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_down_base_correias",
                use_container_width=True,
            )

        with col_u_cor:
            st.markdown("#### 📤 Enviar Dados Atualizados")
            up_arquivo_cor = st.file_uploader("Carregar nova planilha de correias (.xlsx)", type=["xlsx"], key="uploader_novas_correias")
            modo_gravacao = st.radio("Modo de Atualização:", ["Mesclar e Atualizar", "Substituição Completa"], key="radio_modo_up_cor")

            if up_arquivo_cor is not None:
                if st.button("🚀 Confirmar e Atualizar Base de Correias", type="primary", key="btn_executar_up_cor"):
                    try:
                        df_novo_cor = pd.read_excel(up_arquivo_cor)
                        mapa_cols_upload = {}
                        for c in df_novo_cor.columns:
                            c_norm = str(c).strip().lower().replace(" ", "_")
                            if "setor" in c_norm:
                                mapa_cols_upload[c] = "Setor"
                            elif "maquina" in c_norm or "tag" in c_norm:
                                mapa_cols_upload[c] = "Maquina_TAG"
                            elif "tipo" in c_norm and ("1" in c_norm or "sup" in c_norm or "cab" in c_norm):
                                mapa_cols_upload[c] = "Tipo_Correia_1"
                            elif "data" in c_norm and ("1" in c_norm or "sup" in c_norm or "cab" in c_norm):
                                mapa_cols_upload[c] = "Data_Instalacao_1"
                            elif "tipo" in c_norm and ("2" in c_norm or "inf" in c_norm or "tras" in c_norm):
                                mapa_cols_upload[c] = "Tipo_Correia_2"
                            elif "data" in c_norm and ("2" in c_norm or "inf" in c_norm or "tras" in c_norm):
                                mapa_cols_upload[c] = "Data_Instalacao_2"

                        df_novo_cor.rename(columns=mapa_cols_upload, inplace=True)
                        for c in COLUNAS_CORREIAS:
                            if c not in df_novo_cor.columns:
                                df_novo_cor[c] = ""
                        df_novo_cor["Setor"] = df_novo_cor["Setor"].astype(str).str.strip()
                        df_novo_cor["Maquina_TAG"] = df_novo_cor["Maquina_TAG"].astype(str).str.strip().str.upper()

                        def tratar_data_str(val):
                            if pd.isna(val) or val is None or str(val).strip().lower() in ["", "nan", "nat", "none"]:
                                return ""
                            try:
                                return pd.to_datetime(val).strftime("%Y-%m-%d")
                            except Exception:
                                return ""

                        df_novo_cor["Tipo_Correia_1"] = df_novo_cor["Tipo_Correia_1"].apply(formatar_modelo)
                        df_novo_cor["Data_Instalacao_1"] = df_novo_cor["Data_Instalacao_1"].apply(tratar_data_str)
                        df_novo_cor["Tipo_Correia_2"] = df_novo_cor["Tipo_Correia_2"].apply(formatar_modelo)
                        df_novo_cor["Data_Instalacao_2"] = df_novo_cor["Data_Instalacao_2"].apply(tratar_data_str)

                        df_novo_cor = df_novo_cor[COLUNAS_CORREIAS].drop_duplicates(subset=["Setor", "Maquina_TAG"], keep="last")
                        gerar_backup_seguro(ARQUIVO_CORREIAS)

                        if "Substituição Completa" in modo_gravacao:
                            df_final_up_c = df_novo_cor
                        else:
                            chaves_enviadas = set(zip(df_novo_cor["Setor"], df_novo_cor["Maquina_TAG"]))
                            mascara_manter = [(r["Setor"], r["Maquina_TAG"]) not in chaves_enviadas for _, r in df_correias.iterrows()]
                            df_base_restante = df_correias[mascara_manter]
                            df_final_up_c = pd.concat([df_base_restante, df_novo_cor], ignore_index=True)

                        df_final_up_c.to_excel(ARQUIVO_CORREIAS, index=False)
                        invalidar_cache()
                        st.success(f"✅ Base de Correias atualizada com sucesso!")
                        st.rerun()
                    except Exception as erro_proc:
                        st.error(f"Erro ao processar: {erro_proc}")

    with tab_paradas_db:
        st.markdown("### 🛠️ Gestão de Manutenções Corretivas")
        col_d_par, col_u_par = st.columns([1.5, 2.5])

        with col_d_par:
            st.markdown("#### 📥 Descarregar Registos")
            filtro_setor_par = st.selectbox("Filtrar Setor:", ["Todos os Setores"] + list(DICIONARIO_SETORES.keys()), key="sel_export_setor_parada")
            df_export_par = df_paradas if filtro_setor_par == "Todos os Setores" else df_paradas[df_paradas["Setor"] == filtro_setor_par]
            
            if df_export_par.empty:
                df_export_par = pd.DataFrame(columns=COLUNAS_PARADAS)
            else:
                df_export_par = df_export_par[COLUNAS_PARADAS]

            buf_down_p = io.BytesIO()
            with pd.ExcelWriter(buf_down_p, engine="openpyxl") as writer_p:
                df_export_par.to_excel(writer_p, index=False, sheet_name="Corretivas")
            buf_down_p.seek(0)

            st.download_button(
                label="📥 Baixar Planilha de Corretivas (.xlsx)",
                data=buf_down_p,
                file_name=f"manutencoes_corretivas_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_down_base_paradas",
                use_container_width=True,
            )

        with col_u_par:
            st.markdown("#### 📤 Enviar Registos de Corretivas")
            up_arquivo_par = st.file_uploader("Carregar planilha de corretivas (.xlsx)", type=["xlsx"], key="uploader_novas_paradas")
            modo_gravacao_par = st.radio("Modo de Gravação:", ["Incrementar Registos", "Substituição Completa"], key="radio_modo_up_par")

            if up_arquivo_par is not None:
                if st.button("🚀 Confirmar e Atualizar Corretivas", type="primary", key="btn_executar_up_par"):
                    try:
                        df_novo_p = pd.read_excel(up_arquivo_par)
                        mapa_cols_up_p = {}
                        for c in df_novo_p.columns:
                            c_norm = str(c).strip().lower().replace(" ", "_")
                            if "data" in c_norm:
                                mapa_cols_up_p[c] = "Data"
                            elif "setor" in c_norm:
                                mapa_cols_up_p[c] = "Setor"
                            elif "maquina" in c_norm or "tag" in c_norm:
                                mapa_cols_up_p[c] = "Maquina_TAG"
                            elif "tipo" in c_norm:
                                mapa_cols_up_p[c] = "Tipo_Manutencao"
                            elif "servico" in c_norm or "desc" in c_norm:
                                mapa_cols_up_p[c] = "Descricao_Servico"
                            elif "tempo" in c_norm or "hora" in c_norm or "parado" in c_norm:
                                mapa_cols_up_p[c] = "Tempo_Parado_Horas"

                        df_novo_p.rename(columns=mapa_cols_up_p, inplace=True)
                        for c in COLUNAS_PARADAS:
                            if c not in df_novo_p.columns:
                                df_novo_p[c] = 0.0 if c == "Tempo_Parado_Horas" else ""
                        df_novo_p["Setor"] = df_novo_p["Setor"].astype(str).str.strip()
                        df_novo_p["Maquina_TAG"] = df_novo_p["Maquina_TAG"].astype(str).str.strip().str.upper()

                        def tratar_data_p(val):
                            if pd.isna(val) or val is None or str(val).strip().lower() in ["", "nan", "nat", "none"]:
                                return date.today().strftime("%Y-%m-%d")
                            try:
                                return pd.to_datetime(val).strftime("%Y-%m-%d")
                            except Exception:
                                return date.today().strftime("%Y-%m-%d")

                        df_novo_p["Data"] = df_novo_p["Data"].apply(tratar_data_p)
                        df_novo_p["Tempo_Parado_Horas"] = pd.to_numeric(df_novo_p["Tempo_Parado_Horas"], errors="coerce").fillna(0.0)
                        df_novo_p["Tipo_Manutencao"] = df_novo_p["Tipo_Manutencao"].replace({"": "Corretiva"}).fillna("Corretiva")
                        df_novo_p = df_novo_p[(df_novo_p["Tempo_Parado_Horas"] > 0) | (df_novo_p["Descricao_Servico"].astype(str).str.strip() != "")]

                        gerar_backup_seguro(ARQUIVO_PARADAS)
                        if "Substituição Completa" in modo_gravacao_par:
                            df_final_p = df_novo_p[COLUNAS_PARADAS]
                        else:
                            df_final_p = pd.concat([df_paradas, df_novo_p[COLUNAS_PARADAS]], ignore_index=True)

                        df_final_p.to_excel(ARQUIVO_PARADAS, index=False)
                        invalidar_cache()
                        st.success("✅ Manutenções corretivas registradas com sucesso!")
                        st.rerun()
                    except Exception as erro_proc_p:
                        st.error(f"Erro ao processar: {erro_proc_p}")

    with tab_pendencias_db:
        st.markdown("### 📋 Gestão de Manutenções Pendentes")
        col_d_pend, col_u_pend = st.columns([1.5, 2.5])

        with col_d_pend:
            st.markdown("#### 📥 Descarregar Pendências")
            filtro_setor_pend = st.selectbox("Filtrar Setor:", ["Todos os Setores"] + list(DICIONARIO_SETORES.keys()), key="sel_export_setor_pendencia")
            df_export_pend = df_pendencias if filtro_setor_pend == "Todos os Setores" else df_pendencias[df_pendencias["Setor"] == filtro_setor_pend]

            if df_export_pend.empty:
                df_export_pend = pd.DataFrame(columns=COLUNAS_PENDENCIAS)
            else:
                df_export_pend = df_export_pend[COLUNAS_PENDENCIAS]

            buf_down_pend = io.BytesIO()
            with pd.ExcelWriter(buf_down_pend, engine="openpyxl") as writer_pend:
                df_export_pend.to_excel(writer_pend, index=False, sheet_name="Pendencias")
            buf_down_pend.seek(0)

            st.download_button(
                label="📥 Baixar Planilha de Pendências (.xlsx)",
                data=buf_down_pend,
                file_name=f"manutencoes_pendentes_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_down_base_pendencias",
                use_container_width=True,
            )

        with col_u_pend:
            st.markdown("#### 📤 Enviar Backlog de Pendências")
            up_arquivo_pend = st.file_uploader("Carregar planilha de pendências (.xlsx)", type=["xlsx"], key="uploader_novas_pendencias")
            modo_gravacao_pend = st.radio("Modo de Gravação:", ["Incrementar Pendências", "Substituição Completa"], key="radio_modo_up_pend")

            if up_arquivo_pend is not None:
                if st.button("🚀 Confirmar e Atualizar Pendências", type="primary", key="btn_executar_up_pend"):
                    try:
                        df_novo_pend = pd.read_excel(up_arquivo_pend)
                        mapa_cols_up_pend = {}
                        for c in df_novo_pend.columns:
                            c_norm = str(c).strip().lower().replace(" ", "_")
                            if "setor" in c_norm:
                                mapa_cols_up_pend[c] = "Setor"
                            elif "maquina" in c_norm or "tag" in c_norm:
                                mapa_cols_up_pend[c] = "Maquina_TAG"
                            elif "pendencia" in c_norm or "desc" in c_norm or "servico" in c_norm:
                                mapa_cols_up_pend[c] = "Descricao_Pendencia"
                            elif "prioridade" in c_norm or "prio" in c_norm:
                                mapa_cols_up_pend[c] = "Prioridade"
                            elif "status" in c_norm or "situacao" in c_norm:
                                mapa_cols_up_pend[c] = "Status"

                        df_novo_pend.rename(columns=mapa_cols_up_pend, inplace=True)
                        for c in COLUNAS_PENDENCIAS:
                            if c not in df_novo_pend.columns:
                                df_novo_pend[c] = "Pendente" if c == "Status" else "Média" if c == "Prioridade" else ""
                        df_novo_pend["Setor"] = df_novo_pend["Setor"].astype(str).str.strip()
                        df_novo_pend["Maquina_TAG"] = df_novo_pend["Maquina_TAG"].astype(str).str.strip().str.upper()
                        df_novo_pend["Descricao_Pendencia"] = df_novo_pend["Descricao_Pendencia"].astype(str).str.strip()
                        df_novo_pend["Prioridade"] = df_novo_pend["Prioridade"].replace({"": "Média"}).fillna("Média")
                        df_novo_pend["Status"] = df_novo_pend["Status"].replace({"": "Pendente"}).fillna("Pendente")
                        df_novo_pend = df_novo_pend[df_novo_pend["Descricao_Pendencia"].astype(str).str.strip() != ""]

                        gerar_backup_seguro(ARQUIVO_PENDENCIAS)
                        if "Substituição Completa" in modo_gravacao_pend:
                            df_final_pend = df_novo_pend[COLUNAS_PENDENCIAS]
                        else:
                            df_final_pend = pd.concat([df_pendencias, df_novo_pend[COLUNAS_PENDENCIAS]], ignore_index=True)

                        df_final_pend.to_excel(ARQUIVO_PENDENCIAS, index=False)
                        invalidar_cache()
                        st.success("✅ Pendências atualizadas com sucesso!")
                        st.rerun()
                    except Exception as erro_proc_pend:
                        st.error(f"Erro ao processar: {erro_proc_pend}")

    with tab_backups_db:
        st.markdown("#### Cópias de Segurança Geradas Automaticamente")
        if os.path.exists("backups"):
            arquivos_bkp = sorted(os.listdir("backups"), reverse=True)
            if arquivos_bkp:
                dados_bkp = [{"Arquivo de Backup": bkp, "Tamanho": f"{round(os.path.getsize(os.path.join('backups', bkp)) / 1024, 1)} KB"} for bkp in arquivos_bkp]
                st.dataframe(pd.DataFrame(dados_bkp), use_container_width=True, height=300)
            else:
                st.info("Nenhum backup gerado ainda.")
        else:
            st.info("Pasta de backups ainda não inicializada.")

# ------------------------------------------
# 6. GESTÃO CADASTRAL DE MÁQUINAS
# ------------------------------------------
elif tela == "Gestao Maquinas":
    st.markdown("<h2 style='margin:0; font-weight:900;'>🏭 Gestão Cadastral de Máquinas</h2>", unsafe_allow_html=True)
    st.caption("Parametrize os ativos: configure o tipo de fuso, quantidade de correias e dados instalados.")

    c_f_set, c_f_maq = st.columns([1.5, 2.0])
    with c_f_set:
        setores_disponiveis = list(DICIONARIO_SETORES.keys())
        setor_selecionado = st.selectbox("Setor Operacional:", setores_disponiveis, key="sel_setor_gestao_maq")

    maquinas_do_setor = obter_maquinas_setor(setor_selecionado, df_correias, df_fusos)

    with c_f_maq:
        maq_selecionada = st.selectbox("Máquina (TAG):", maquinas_do_setor, key="sel_maq_gestao_maq")

    fuso_atual = "FAG"
    sub_fuso = df_fusos[(df_fusos["Setor"] == setor_selecionado) & (df_fusos["Maquina_TAG"] == maq_selecionada)]
    if not sub_fuso.empty:
        val_fuso = str(sub_fuso.iloc[-1].get("Tipo_Fuso", "")).strip()
        if val_fuso in OPCOES_TIPO_FUSO:
            fuso_atual = val_fuso
        else:
            fuso_atual = "FAG" if setor_selecionado == "Setor A" else "TEP" if setor_selecionado == "Setor B" else "M4BA" if setor_selecionado == "Setor Látex" else "MENEGATTO"
    else:
        fuso_atual = "FAG" if setor_selecionado == "Setor A" else "TEP" if setor_selecionado == "Setor B" else "M4BA" if setor_selecionado == "Setor Látex" else "MENEGATTO"

    sub_cor = df_correias[(df_correias["Setor"] == setor_selecionado) & (df_correias["Maquina_TAG"] == maq_selecionada)]
    mod1_atual, dt1_atual = "", None
    mod2_atual, dt2_atual = "", None

    if not sub_cor.empty:
        ult_c = sub_cor.iloc[-1]
        mod1_atual = formatar_modelo(ult_c.get("Tipo_Correia_1", ""))
        raw_d1 = ult_c.get("Data_Instalacao_1", "")
        if pd.notna(raw_d1) and str(raw_d1).strip() not in ["", "nan", "NaT", "None"]:
            try:
                dt1_atual = pd.to_datetime(raw_d1).date()
            except Exception:
                dt1_atual = None

        mod2_atual = formatar_modelo(ult_c.get("Tipo_Correia_2", ""))
        raw_d2 = ult_c.get("Data_Instalacao_2", "")
        if pd.notna(raw_d2) and str(raw_d2).strip() not in ["", "nan", "NaT", "None"]:
            try:
                dt2_atual = pd.to_datetime(raw_d2).date()
            except Exception:
                dt2_atual = None

    tem_duas_inicial = bool(mod2_atual or dt2_atual)

    st.markdown("---")

    with st.container(border=True):
        st.markdown(f"#### ⚙️ Parâmetros do Ativo: **{maq_selecionada}** ({setor_selecionado})")
        
        c_fuso, c_qtd_cor = st.columns([1.5, 2.0])
        
        with c_fuso:
            idx_fuso = OPCOES_TIPO_FUSO.index(fuso_atual) if fuso_atual in OPCOES_TIPO_FUSO else 0
            novo_fuso = st.selectbox("Tipo de Fuso:", OPCOES_TIPO_FUSO, index=idx_fuso, key=f"fuso_edit_{maq_selecionada}")

        with c_qtd_cor:
            qtd_correias_opc = st.radio(
                "Quantidade de Correias:",
                ["1 Correia (Única)", "2 Correias (Superior/Cabeceira e Inferior/Traseira)"],
                index=1 if tem_duas_inicial else 0,
                key=f"qtd_cor_edit_{maq_selecionada}",
                horizontal=True
            )

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

        if qtd_correias_opc == "1 Correia (Única)":
            c_c1_mod, c_c1_dt = st.columns(2)
            with c_c1_mod:
                novo_mod1 = st.text_input("Modelo da Correia Única:", value=mod1_atual, key=f"inp_mod1_u_{maq_selecionada}", placeholder="Ex: 36.100")
            with c_c1_dt:
                nova_dt1 = st.date_input("Data de Instalação:", value=dt1_atual, key=f"inp_dt1_u_{maq_selecionada}", format="DD/MM/YYYY")

            novo_mod2 = ""
            nova_dt2 = None
        else:
            c_sup1, c_sup2 = st.columns(2)
            with c_sup1:
                st.markdown("<p style='font-size:0.85rem; font-weight:800; color:#0f172a; margin-bottom:2px;'>🔼 Superior / Cabeceira</p>", unsafe_allow_html=True)
                novo_mod1 = st.text_input("Modelo (Superior / Cabeceira):", value=mod1_atual, key=f"inp_mod1_d_{maq_selecionada}", placeholder="Ex: 19.500")
                nova_dt1 = st.date_input("Data de Instalação (Superior):", value=dt1_atual, key=f"inp_dt1_d_{maq_selecionada}", format="DD/MM/YYYY")

            with c_sup2:
                st.markdown("<p style='font-size:0.85rem; font-weight:800; color:#0f172a; margin-bottom:2px;'>🔽 Inferior / Traseira</p>", unsafe_allow_html=True)
                novo_mod2 = st.text_input("Modelo (Inferior / Traseira):", value=mod2_atual, key=f"inp_mod2_d_{maq_selecionada}", placeholder="Ex: 18.050")
                nova_dt2 = st.date_input("Data de Instalação (Inferior):", value=dt2_atual, key=f"inp_dt2_d_{maq_selecionada}", format="DD/MM/YYYY")

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        if st.button("💾 Salvar Parâmetros da Máquina", type="primary", key=f"btn_salvar_param_{maq_selecionada}"):
            mask_fusos_maq = (df_fusos["Setor"] == setor_selecionado) & (df_fusos["Maquina_TAG"] == maq_selecionada)
            if mask_fusos_maq.any():
                df_fusos.loc[mask_fusos_maq, "Tipo_Fuso"] = novo_fuso
            else:
                novos_ano = [
                    {"Ano": 2026, "Mes": m_n, "Dia": 1, "Setor": setor_selecionado, "Maquina_TAG": maq_selecionada, "Quantidade_Quebras": 0, "Tipo_Fuso": novo_fuso}
                    for m_n in LISTA_MESES_PUROS
                ]
                df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_ano)], ignore_index=True)

            gerar_backup_seguro(ARQUIVO_FUSOS)
            df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

            m1_fmt = formatar_modelo(novo_mod1)
            d1_fmt = str(nova_dt1) if nova_dt1 is not None else ""
            m2_fmt = formatar_modelo(novo_mod2)
            d2_fmt = str(nova_dt2) if nova_dt2 is not None else ""

            mask_cor_maq = (df_correias["Setor"] == setor_selecionado) & (df_correias["Maquina_TAG"] == maq_selecionada)
            if mask_cor_maq.any():
                idx_c = df_correias[mask_cor_maq].index[0]
                df_correias.loc[idx_c, "Tipo_Correia_1"] = m1_fmt
                df_correias.loc[idx_c, "Data_Instalacao_1"] = d1_fmt
                df_correias.loc[idx_c, "Tipo_Correia_2"] = m2_fmt
                df_correias.loc[idx_c, "Data_Instalacao_2"] = d2_fmt
            else:
                novo_registro_cor = {
                    "Setor": setor_selecionado,
                    "Maquina_TAG": maq_selecionada,
                    "Tipo_Correia_1": m1_fmt,
                    "Data_Instalacao_1": d1_fmt,
                    "Tipo_Correia_2": m2_fmt,
                    "Data_Instalacao_2": d2_fmt,
                }
                df_correias = pd.concat([df_correias, pd.DataFrame([novo_registro_cor])], ignore_index=True)

            gerar_backup_seguro(ARQUIVO_CORREIAS)
            df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
            invalidar_cache()
            st.success(f"✅ Configurações da máquina {maq_selecionada} atualizadas com sucesso!")
            st.rerun()
