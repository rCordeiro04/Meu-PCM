import os
import shutil
import calendar
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

ARQUIVO_FUSOS = "lancamentos_fusos_v5.xlsx"
ARQUIVO_CORREIAS = "lancamentos_correias_v4.xlsx"

COLUNAS_FUSOS = [
    "Ano", "Mes", "Dia", "Setor", "Maquina_TAG", "Quantidade_Quebras", "Tipo_Fuso"
]

COLUNAS_CORREIAS = [
    "Setor", "Maquina_TAG", "Tipo_Correia_1", "Data_Instalacao_1", "Tipo_Correia_2", "Data_Instalacao_2"
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

# Mapeamento Oficial dos Ativos
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

# ==========================================
# INICIALIZAÇÃO E CARGA OTIMIZADA DAS BASES
# ==========================================
def carregar_dados_iniciais():
    # 1. Base Fusos
    df_f = None
    if os.path.exists(ARQUIVO_FUSOS):
        try:
            df_f = pd.read_excel(ARQUIVO_FUSOS)
            if "Dia" not in df_f.columns:
                df_f["Dia"] = 1
        except Exception:
            df_f = pd.DataFrame(columns=COLUNAS_FUSOS)
    else:
        df_f = pd.DataFrame(columns=COLUNAS_FUSOS)

    # 2. Base Correias
    df_c = None
    if os.path.exists(ARQUIVO_CORREIAS):
        try:
            df_c = pd.read_excel(ARQUIVO_CORREIAS)
        except Exception:
            df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)
    else:
        df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)

    for c in COLUNAS_CORREIAS:
        if c not in df_c.columns:
            df_c[c] = ""
        df_c[c] = df_c[c].astype(object)

    return df_f, df_c

df_fusos, df_correias = carregar_dados_iniciais()

# ==========================================
# INJEÇÃO HISTÓRICA DE CORREIAS (OTIMIZADA)
# ==========================================
DADOS_HISTORICOS_CORREIAS = [
    ("Setor A", "L-01", "36.100", "2026-08-21", "", ""),
    ("Setor A", "L-02", "", "", "", ""),
    ("Setor A", "L-03", "36.100", "2026-05-04", "", ""),
    ("Setor A", "L-04", "36.100", "2025-01-07", "", ""),
    ("Setor A", "L-05", "36.100", "2026-05-27", "", ""),
    ("Setor A", "L-06", "36.100", "2024-11-19", "", ""),
    ("Setor A", "L-07", "36.100", "2025-11-25", "", ""),
    ("Setor A", "L-08", "36.100", "2025-05-20", "", ""),
    ("Setor A", "L-09", "36.100", "2026-05-13", "", ""),
    ("Setor A", "L-10", "36.100", "2025-10-08", "", ""),
    ("Setor A", "L-11", "19.500", "2025-01-27", "18.050", "2026-08-21"),
    ("Setor A", "L-12", "19.500", "2026-01-17", "18.050", "2026-01-17"),
    ("Setor A", "L-13", "", "", "18.050", "2025-08-11"),
    ("Setor A", "L-14", "19.500", "2025-02-02", "18.050", "2026-03-26"),
    ("Setor A", "L-15", "36.100", "2025-03-07", "", ""),
    ("Setor A", "L-16", "36.100", "2025-02-12", "", ""),
    ("Setor A", "L-17", "36.100", "2025-08-20", "", ""),
    ("Setor A", "L-18", "36.100", "2026-08-24", "", ""),
    ("Setor A", "L-19", "36.100", "2025-04-12", "", ""),
    ("Setor A", "L-20", "36.100", "2025-07-01", "", ""),
    ("Setor A", "L-21", "36.100", "2026-05-20", "", ""),
    ("Setor A", "L-22", "36.100", "2026-05-19", "", ""),
    ("Setor A", "L-23", "36.100", "2025-09-01", "", ""),
    ("Setor A", "L-24", "36.100", "2025-10-27", "", ""),
    ("Setor A", "L-25", "36.100", "2026-06-29", "", ""),
    ("Setor A", "L-26", "", "", "18.050", "2026-09-14"),
    ("Setor A", "L-27", "", "", "18.050", "2025-03-13"),
    ("Setor A", "L-28", "19.500", "2025-03-22", "", ""),
    ("Setor B", "L-29", "36.100", "2025-11-13", "", ""),
    ("Setor B", "L-30", "36.100", "2026-09-10", "", ""),
    ("Setor B", "L-31", "36.100", "2026-08-24", "", ""),
    ("Setor B", "L-33", "36.100", "2026-03-20", "", ""),
    ("Setor B", "L-34", "36.100", "2026-08-06", "", ""),
    ("Setor B", "L-36", "36.100", "2026-02-26", "", ""),
    ("Setor B", "L-38", "36.100", "2025-10-02", "", ""),
    ("Setor B", "L-39", "36.100", "2025-03-21", "", ""),
    ("Setor B", "L-40", "36.100", "2026-03-11", "", ""),
    ("Setor B", "L-42", "", "", "34.870", "2025-10-02"),
    ("Setor B", "L-43", "", "", "34.870", "2025-08-05"),
    ("Setor B", "L-44", "", "", "34.870", "2025-08-08"),
    ("Setor B", "L-50", "19.500", "2026-07-28", "18.050", "2026-03-05"),
    ("Setor B", "L-51", "36.100", "2026-06-30", "", ""),
    ("Setor Látex", "B-72", "33.990", "2026-01-02", "", ""),
    ("Setor Látex", "B-73", "33.990", "2026-06-13", "34.870", "2026-06-13"),
    ("Setor Látex", "B-74", "33.990", "2026-03-10", "34.870", "2025-02-15"),
    ("Setor Látex", "B-78", "", "", "34.870", "2025-12-29"),
    ("Setor Látex", "B-79", "", "", "34.870", "2024-11-30"),
    ("Setor Menegatto", "B-93", "", "", "38.740", "2025-04-16"),
    ("Setor Menegatto", "B-94", "", "", "38.740", "2025-05-01"),
    ("Setor Menegatto", "B-95", "", "", "38.740", "2025-04-17"),
    ("Setor Menegatto", "B-96", "", "", "38.740", "2026-06-25"),
    ("Setor Menegatto", "B-97", "", "", "38.740", "2025-05-05"),
    ("Setor Menegatto", "B-98", "", "", "38.740", "2024-12-11"),
    ("Setor Menegatto", "B-99", "", "", "38.740", "2024-12-06"),
    ("Setor Menegatto", "B-100", "", "2025-05-27", "38.740", "2025-12-04"),
    ("Setor Menegatto", "B-101", "", "2025-10-04", "38.740", "2026-04-03"),
]

salvar_cor_init = False
for s_cor, tag_cor, m1_cor, dt1_cor, m2_cor, dt2_cor in DADOS_HISTORICOS_CORREIAS:
    mask = (df_correias["Setor"] == s_cor) & (df_correias["Maquina_TAG"] == tag_cor)
    if not mask.any():
        novo = {
            "Setor": s_cor, "Maquina_TAG": tag_cor,
            "Tipo_Correia_1": str(m1_cor), "Data_Instalacao_1": str(dt1_cor),
            "Tipo_Correia_2": str(m2_cor), "Data_Instalacao_2": str(dt2_cor),
        }
        df_correias = pd.concat([df_correias, pd.DataFrame([novo])], ignore_index=True)
        salvar_cor_init = True
    else:
        idx = df_correias[mask].index[0]
        if m1_cor and not str(df_correias.at[idx, "Tipo_Correia_1"]).strip():
            df_correias.loc[idx, "Tipo_Correia_1"] = str(m1_cor)
            df_correias.loc[idx, "Data_Instalacao_1"] = str(dt1_cor)
            salvar_cor_init = True
        if m2_cor and not str(df_correias.at[idx, "Tipo_Correia_2"]).strip():
            df_correias.loc[idx, "Tipo_Correia_2"] = str(m2_cor)
            df_correias.loc[idx, "Data_Instalacao_2"] = str(dt2_cor)
            salvar_cor_init = True

if salvar_cor_init:
    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

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

if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Painel Fusos"
if "maq_clicada_cor" not in st.session_state:
    st.session_state.maq_clicada_cor = None
if "card_selecionado_kpi" not in st.session_state:
    st.session_state.card_selecionado_kpi = None
if "aba_setor_fuso" not in st.session_state:
    st.session_state.aba_setor_fuso = "Geral"

def navegar(p):
    st.session_state.pagina_atual = p

# ==========================================
# PROCESSAMENTO DE CORREIAS
# ==========================================
data_hoje = date.today()
dados_maquinas = {}
css_botoes = []

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

for maq_tag in todas_maquinas_totais:
    setor_m = mapa_setor_maquina.get(maq_tag, "Setor A")
    reg_maq = df_correias[(df_correias["Setor"] == setor_m) & (df_correias["Maquina_TAG"] == maq_tag)]

    t1, d1_str, t1_uso, c1_score, tem_c1 = "Não informada", "Sem registro", "Sem histórico", None, False
    t2, d2_str, t2_uso, c2_score, tem_c2 = "Não informada", "Sem registro", "Sem histórico", None, False

    if not reg_maq.empty:
        ult = reg_maq.iloc[-1]
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
        cor_grad = "linear-gradient(135deg, #10b981, #059669)" if pior == 1 else "linear-gradient(135deg, #f59e0b, #d97706)" if pior == 2 else "linear-gradient(135deg, #ef4444, #dc2626)"
        cor_borda = "#047857" if pior == 1 else "#b45309" if pior == 2 else "#b91c1c"
        status_label = "Nova" if pior == 1 else "Meia-Vida" if pior == 2 else "Troca Necessária"
    else:
        classe_card, cor_grad, cor_borda, status_label = "status-cinza", "linear-gradient(135deg, #64748b, #475569)", "#334155", "Sem Dados"

    dados_maquinas[maq_tag] = {
        "setor": setor_m, "t1": t1, "d1": d1_str, "uso1": t1_uso, "tem_c1": tem_c1,
        "t2": t2, "d2": d2_str, "uso2": t2_uso, "tem_c2": tem_c2,
        "status_label": status_label, "classe_card": classe_card
    }

    chave_btn = f"btn_q_{maq_tag.replace('-', '_')}"
    css_botoes.append(f"""
        button[key="{chave_btn}"], div.st-key-{chave_btn} button {{
            background: {cor_grad} !important; color: #ffffff !important; border: 1px solid {cor_borda} !important;
        }}
    """)

regras_css_botoes = "\n".join(css_botoes)

st.markdown(
    f"""
    <style>
        .block-container {{ padding: 4.4rem 2rem 1rem 2rem !important; }}
        [data-testid="stSidebar"] {{ background-color: #0f172a !important; border-right: 1px solid #1e293b !important; }}
        [data-testid="stSidebar"] h2, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{ color: #f8fafc; }}
        
        [data-testid="stSidebar"] div[data-testid="stRadio"] {{ background: #1e293b; padding: 4px; border-radius: 10px; border: 1px solid #334155; margin-bottom: 12px; }}
        [data-testid="stSidebar"] div[data-testid="stRadio"] > div {{ flex-direction: row; justify-content: space-between; }}
        [data-testid="stSidebar"] div[data-testid="stRadio"] label p {{ color: #f1f5f9 !important; font-weight: 700 !important; }}
        
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] {{
            background-color: #1e293b !important; color: #e2e8f0 !important; border: 1px solid #334155 !important;
            border-radius: 8px !important; font-weight: 700 !important; height: 38px !important;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] p {{ color: #e2e8f0 !important; }}
        [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {{ background-color: #334155 !important; }}
        
        [data-testid="stSidebar"] .stButton > button[kind="primary"] {{
            background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: #ffffff !important;
            border: 1px solid #b91c1c !important; border-radius: 8px !important; font-weight: 800 !important; height: 38px !important;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="primary"] p {{ color: #ffffff !important; }}
        
        div[data-testid="column"] {{ padding: 1px !important; margin: 0px !important; }}
        div[data-testid="stHorizontalBlock"] {{ gap: 4px !important; margin-bottom: 4px !important; }}
        div[data-testid="stVegaLiteChart"] summary, div[data-testid="stVegaLiteChart"] .vega-actions {{ display: none !important; }}
        div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {{ overscroll-behavior: contain; }}
        
        div[data-testid="stButton"] button {{
            padding: 0px !important; font-size: 0.8rem !important; height: 33px !important;
            min-height: 33px !important; line-height: 31px !important; border-radius: 7px !important;
        }}
        {regras_css_botoes}
        
        .card-kpi-bonito {{
            background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; padding: 10px 14px;
            display: flex; align-items: center; justify-content: space-between; height: 64px; box-sizing: border-box;
            position: relative; overflow: hidden;
        }}
        .card-kpi-bonito::after {{ content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 5px; }}
        .card-kpi-bonito.c-total::after {{ background: #475569; }}
        .card-kpi-bonito.c-ok::after {{ background: #10b981; }}
        .card-kpi-bonito.c-warn::after {{ background: #f59e0b; }}
        .card-kpi-bonito.c-crit::after {{ background: #ef4444; }}
        .kpi-val {{ font-size: 1.35rem; font-weight: 800; line-height: 1; font-family: ui-monospace, monospace; }}
        .kpi-lbl {{ font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }}
        
        .alerta-manutencao {{
            background: #fef2f2; border: 1px solid #fecaca; border-left: 6px solid #ef4444; border-radius: 8px;
            padding: 8px 14px; margin: 6px 0 10px 0; font-size: 0.83rem; font-weight: 600; color: #991b1b;
        }}
        .chip-critico {{ background: #fee2e2; border: 1px solid #fca5a5; color: #991b1b; padding: 2px 8px; border-radius: 5px; font-weight: 800; }}
        .hud-detalhe {{
            background: #ffffff; border: 1px solid #cbd5e1; border-radius: 10px; padding: 14px 18px; margin: 6px 0 12px 0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05); border-left: 6px solid #64748b;
        }}
        .hud-detalhe.status-verde {{ border-left-color: #10b981; }}
        .hud-detalhe.status-amarelo {{ border-left-color: #f59e0b; }}
        .hud-detalhe.status-vermelho {{ border-left-color: #ef4444; }}
        .hud-detalhe.status-cinza {{ border-left-color: #94a3b8; }}
        
        .tag-pill {{ background: #f8fafc; border: 1px solid #e2e8f0; padding: 5px 12px; border-radius: 6px; font-size: 0.82rem; font-weight: 700; color: #334155; }}
        .badge-status {{ padding: 4px 10px; border-radius: 6px; font-size: 0.75rem; font-weight: 800; text-transform: uppercase; }}
        .badge-verde {{ background: #d1fae5; color: #065f46; }}
        .badge-amarelo {{ background: #fef3c7; color: #92400e; }}
        .badge-vermelho {{ background: #fee2e2; color: #991b1b; }}
        .badge-cinza {{ background: #e2e8f0; color: #475569; }}
        
        .pill-legenda {{ display: inline-flex; align-items: center; gap: 6px; font-size: 0.78rem; font-weight: 700; background: #f8fafc; border: 1px solid #e2e8f0; padding: 4px 10px; border-radius: 20px; }}
        .dot-legenda {{ width: 9px; height: 9px; border-radius: 50%; display: inline-block; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='font-size:1.6rem; font-weight:900; margin:0 0 10px 0;'>⚙️ Portal PCM</h2>", unsafe_allow_html=True)
    is_lancto = st.session_state.pagina_atual in ["Lançamento Fusos", "Correias", "Preventiva", "Máquinas"]
    modo = st.radio("Modo", ["📊 Painéis", "📝 Lançamentos"], index=1 if is_lancto else 0, label_visibility="collapsed")

    if modo == "📊 Painéis":
        st.markdown("<p style='font-size:0.72rem; font-weight:800; color:#94a3b8; margin:6px 0;'>PAINÉIS GERENCIAIS</p>", unsafe_allow_html=True)
        st.button("🔩 Painel de Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Fusos" else "secondary", on_click=navegar, args=("Painel Fusos",))
        st.button("🔄 Painel de Correias", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Correias" else "secondary", on_click=navegar, args=("Painel Correias",))
        if is_lancto:
            st.session_state.pagina_atual = "Painel Fusos"
            st.rerun()
    else:
        st.markdown("<p style='font-size:0.72rem; font-weight:800; color:#94a3b8; margin:6px 0;'>APONTAMENTOS</p>", unsafe_allow_html=True)
        st.button("🔩 Fechamento de Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Lançamento Fusos" else "secondary", on_click=navegar, args=("Lançamento Fusos",))
        st.button("🔄 Gestão de Correias", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Correias" else "secondary", on_click=navegar, args=("Correias",))
        st.button("🛠️ Preventiva", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Preventiva" else "secondary", on_click=navegar, args=("Preventiva",))
        st.button("🏭 Máquinas", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Máquinas" else "secondary", on_click=navegar, args=("Máquinas",))
        if not is_lancto:
            st.session_state.pagina_atual = "Lançamento Fusos"
            st.rerun()

    st.markdown("---")
    st.markdown("<p style='font-size:0.72rem; font-weight:800; color:#94a3b8; margin:4px 0;'>📥 BACKUP DAS BASES</p>", unsafe_allow_html=True)
    if os.path.exists(ARQUIVO_CORREIAS):
        with open(ARQUIVO_CORREIAS, "rb") as fc:
            st.download_button("Baixar Correias (.xlsx)", fc, "lancamentos_correias.xlsx", use_container_width=True)
    if os.path.exists(ARQUIVO_FUSOS):
        with open(ARQUIVO_FUSOS, "rb") as ff:
            st.download_button("Baixar Fusos (.xlsx)", ff, "lancamentos_fusos.xlsx", use_container_width=True)

# ==========================================
# ÁREA PRINCIPAL
# ==========================================
tela = st.session_state.pagina_atual

# 1. PAINEL CORREIAS
if tela == "Painel Correias":
    c_t, c_f1, c_f2, c_leg = st.columns([3.2, 1.4, 1.4, 5.0])
    with c_t:
        st.markdown("<h2 style='margin:0; font-weight:900;'>Dashboard Correias</h2>", unsafe_allow_html=True)
    with c_f1:
        filtro_setor = st.selectbox("Setor", ["Todos os Setores"] + list(DICIONARIO_SETORES.keys()), label_visibility="collapsed")
    with c_f2:
        mods_un = sorted(list({r["modelo"] for r in lista_correias_todas if r["modelo"] and r["modelo"] != "Não informada"}))
        filtro_modelo = st.selectbox("Tipo", ["Todos os Tipos"] + mods_un, label_visibility="collapsed")
    with c_leg:
        st.markdown("""
            <div style="height:40px; display:flex; align-items:center; justify-content:flex-end; gap:6px;">
                <span class='pill-legenda'><span class='dot-legenda' style='background:#10b981;'></span> Nova (&le; 1a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#f59e0b;'></span> Meia (1-1.5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#ef4444;'></span> Urgente (&gt; 1.5a)</span>
            </div>
        """, unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Correias</div><div class='kpi-val'>{len(lista_correias_todas)}</div></div><div>📦</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Novas</div><div class='kpi-val' style='color:#059669;'>{len(lista_correias_novas)}</div></div><div>🟢</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Meia-Vida</div><div class='kpi-val' style='color:#d97706;'>{len(lista_correias_meia)}</div></div><div>🟡</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Críticas</div><div class='kpi-val' style='color:#dc2626;'>{len(lista_correias_criticas)}</div></div><div>🔴</div></div>", unsafe_allow_html=True)

    if lista_correias_criticas:
        chips = " ".join([f"<span class='chip-critico'>🏷️ {k}: <b>{v} un.</b></span>" for k, v in pd.DataFrame(lista_correias_criticas)["modelo"].value_counts().items()])
        st.markdown(f"<div class='alerta-manutencao'>🚨 <b>Alerta:</b> &nbsp; {chips}</div>", unsafe_allow_html=True)

    if st.session_state.maq_clicada_cor:
        sel = st.session_state.maq_clicada_cor
        b_cor = "badge-verde" if sel["status_label"] == "Nova" else "badge-amarelo" if sel["status_label"] == "Meia-Vida" else "badge-vermelho" if sel["status_label"] == "Troca Necessária" else "badge-cinza"
        c_hud, c_cls = st.columns([6.2, 0.8])
        with c_hud:
            st.markdown(f"""
                <div class="hud-detalhe {sel['classe_card']}">
                    <div style="display:flex; justify-content:space-between; margin-bottom:6px;">
                        <span class="tag-pill" style="background:#0f172a; color:#ffffff;">⚙️ {sel['tag']} - {sel['setor']}</span>
                        <span class="badge-status {b_cor}">{sel['status_label']}</span>
                    </div>
                    <div style="display:flex; gap:8px;">
                        <span class="tag-pill">🔼 <b>Superior:</b> {sel['t1']} | {sel['d1']} | {sel['uso1']}</span>
                        <span class="tag-pill">🔽 <b>Inferior:</b> {sel['t2']} | {sel['d2']} | {sel['uso2']}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with c_cls:
            if st.button("✖ Fechar", key="btn_cls_hud"):
                st.session_state.maq_clicada_cor = None
                st.rerun()

    maqs_grid = []
    for s in DICIONARIO_SETORES.keys():
        if filtro_setor != "Todos os Setores" and s != filtro_setor:
            continue
        for m in obter_maquinas_setor(s, df_correias, df_fusos):
            if filtro_modelo != "Todos os Tipos":
                r_m = df_correias[(df_correias["Setor"] == s) & (df_correias["Maquina_TAG"] == m)]
                if r_m.empty:
                    continue
                ult_m = r_m.iloc[-1]
                if formatar_modelo(ult_m.get("Tipo_Correia_1")) != filtro_modelo and formatar_modelo(ult_m.get("Tipo_Correia_2")) != filtro_modelo:
                    continue
            maqs_grid.append(m)

    cols_g = 12
    for chunk in [maqs_grid[i:i + cols_g] for i in range(0, len(maqs_grid), cols_g)]:
        cols = st.columns(cols_g)
        for i, m in enumerate(chunk):
            if cols[i].button(m, key=f"btn_q_{m.replace('-', '_')}", use_container_width=True):
                st.session_state.maq_clicada_cor = {"tag": m, **dados_maquinas[m]}
                st.rerun()

# 2. LANÇAMENTO CORREIAS
elif tela == "Correias":
    st.title("🔄 Lançamento: Gestão de Correias")
    setor_sel = st.selectbox("Setor Operacional", list(DICIONARIO_SETORES.keys()))
    df_cur_c = pd.read_excel(ARQUIVO_CORREIAS)
    maqs_s = obter_maquinas_setor(setor_sel, df_cur_c, df_fusos)

    grade = []
    for m in maqs_s:
        reg = df_cur_c[(df_cur_c["Setor"] == setor_sel) & (df_cur_c["Maquina_TAG"] == m)]
        t1, dt1, t2, dt2 = "", None, "", None
        if not reg.empty:
            u = reg.iloc[-1]
            t1 = formatar_modelo(u.get("Tipo_Correia_1"))
            try:
                dt1 = pd.to_datetime(u.get("Data_Instalacao_1")).date()
            except Exception:
                pass
            t2 = formatar_modelo(u.get("Tipo_Correia_2"))
            try:
                dt2 = pd.to_datetime(u.get("Data_Instalacao_2")).date()
            except Exception:
                pass
        grade.append({"Máquina": m, "Modelo (Superior)": t1, "Data (Superior)": dt1, "Modelo (Inferior)": t2, "Data (Inferior)": dt2})

    editado = st.data_editor(
        pd.DataFrame(grade),
        column_config={
            "Máquina": st.column_config.TextColumn(disabled=True),
            "Data (Superior)": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Data (Inferior)": st.column_config.DateColumn(format="DD/MM/YYYY"),
        },
        hide_index=True,
        use_container_width=True,
        height=450
    )

    if st.button("💾 Salvar Correias", type="primary"):
        limpo = df_cur_c[df_cur_c["Setor"] != setor_sel]
        novos = []
        for _, r in editado.iterrows():
            novos.append({
                "Setor": setor_sel, "Maquina_TAG": r["Máquina"],
                "Tipo_Correia_1": formatar_modelo(r["Modelo (Superior)"]),
                "Data_Instalacao_1": str(r["Data (Superior)"]) if pd.notna(r["Data (Superior)"]) else "",
                "Tipo_Correia_2": formatar_modelo(r["Modelo (Inferior)"]),
                "Data_Instalacao_2": str(r["Data (Inferior)"]) if pd.notna(r["Data (Inferior)"]) else "",
            })
        final = pd.concat([limpo, pd.DataFrame(novos)], ignore_index=True)
        gerar_backup_seguro(ARQUIVO_CORREIAS)
        final.to_excel(ARQUIVO_CORREIAS, index=False)
        st.success("Salvo com sucesso!")
        st.rerun()

# 3. PAINEL FUSOS
elif tela == "Painel Fusos":
    cf_t, cf_a = st.columns([3.8, 1.4])
    cf_t.markdown("<h2 style='margin:0; font-weight:900;'>Dashboard Fusos</h2>", unsafe_allow_html=True)
    ano_f = cf_a.selectbox("Ano", [2024, 2025, 2026, 2027], index=2, label_visibility="collapsed")

    div_meses = 12 if int(ano_f) < date.today().year else max(1, date.today().month)
    b_geral, b_sa, b_sb, b_lat, b_men = st.columns(5)
    for b_col, nome_aba in zip([b_geral, b_sa, b_sb, b_lat, b_men], ["Geral", "Setor A", "Setor B", "Setor Látex", "Setor Menegatto"]):
        if b_col.button(nome_aba, use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == nome_aba else "secondary"):
            st.session_state.aba_setor_fuso = nome_aba
            st.rerun()

    df_ano_f = df_fusos[df_fusos["Ano"] == int(ano_f)].copy()
    if df_ano_f.empty:
        df_ano_f = pd.DataFrame(columns=COLUNAS_FUSOS)

    if st.session_state.aba_setor_fuso == "Geral":
        tot = int(df_ano_f["Quantidade_Quebras"].sum())
        kf1, kf2 = st.columns(2)
        kf1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Fábrica</div><div class='kpi-val'>{tot}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
        kf2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal</div><div class='kpi-val'>{round(tot / div_meses, 1)}</div></div><div>📈</div></div>", unsafe_allow_html=True)

        cg1, cg2 = st.columns(2)
        cores = {"Setor A": "#2563eb", "Setor B": "#d97706", "Setor Látex": "#059669", "Setor Menegatto": "#dc2626"}
        for i, s_nome in enumerate(DICIONARIO_SETORES.keys()):
            with (cg1 if i % 2 == 0 else cg2):
                df_s = df_ano_f[df_ano_f["Setor"] == s_nome].groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                df_s["Mes_Abrev"] = df_s["Mes"].map(MAPA_MES_ABREV)
                c_bar = alt.Chart(df_s).mark_bar(color=cores[s_nome], cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras")
                ).properties(height=180, title=f"🏭 {s_nome}")
                st.altair_chart(c_bar, use_container_width=True)

    else:
        s_ativo = st.session_state.aba_setor_fuso
        df_sa = df_ano_f[df_ano_f["Setor"] == s_ativo]
        tot_s = int(df_sa["Quantidade_Quebras"].sum())
        ks1, ks2 = st.columns(2)
        ks1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total {s_ativo}</div><div class='kpi-val'>{tot_s}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
        ks2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal</div><div class='kpi-val'>{round(tot_s / div_meses, 1)}</div></div><div>📈</div></div>", unsafe_allow_html=True)

        # Mapa de Calor Vetorizado Rápido
        maqs_calor = obter_maquinas_setor(s_ativo, df_correias, df_fusos)
        idx_grid = pd.MultiIndex.from_product([maqs_calor, LISTA_MESES_PUROS], names=["MAQ", "Mes"]).to_frame().reset_index(drop=True)
        idx_grid["MES"] = idx_grid["Mes"].map(MAPA_MES_ABREV)
        agrup_c = df_sa.groupby(["Maquina_TAG", "Mes"])["Quantidade_Quebras"].sum().reset_index()
        m_calor = pd.merge(idx_grid, agrup_c, left_on=["MAQ", "Mes"], right_on=["Maquina_TAG", "Mes"], how="left").fillna(0)

        rect = alt.Chart(m_calor).mark_rect(stroke="#fff").encode(
            x=alt.X("MES:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(orient="top")),
            y=alt.Y("MAQ:N", sort=maqs_calor, title=None),
            color=alt.Color("Quantidade_Quebras:Q", scale=alt.Scale(domain=[0, 3, 8, 15], range=["#dcfce7", "#fef08a", "#f97316", "#dc2626"]))
        )
        txt = alt.Chart(m_calor).mark_text(fontWeight=700).encode(
            x=alt.X("MES:N", sort=ORDEM_MESES_ABREV),
            y=alt.Y("MAQ:N", sort=maqs_calor),
            text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
        )
        st.altair_chart((rect + txt).properties(height=max(260, len(maqs_calor) * 22)), use_container_width=True)

# 4. LANÇAMENTO FUSOS (MATRIZ DIÁRIA)
elif tela == "Lançamento Fusos":
    st.title("🔩 Lançamento: Apontamento Diário de Fusos")
    c_ano, c_set = st.columns([1.2, 2])
    ano_sel = c_ano.selectbox("Ano", [2024, 2025, 2026, 2027], index=2)
    setor_sel = c_set.selectbox("Setor", list(DICIONARIO_SETORES.keys()))

    abas_m = st.tabs(LISTA_MESES_PUROS)
    maqs_f = obter_maquinas_setor(setor_sel, df_correias, df_fusos)

    for idx, m_nome in enumerate(LISTA_MESES_PUROS):
        with abas_m[idx]:
            _, dias_no_mes = calendar.monthrange(int(ano_sel), idx + 1)
            cols_d = [f"{d:02d}" for d in range(1, dias_no_mes + 1)]

            df_sub_f = df_fusos[(df_fusos["Ano"] == ano_sel) & (df_fusos["Mes"] == m_nome) & (df_fusos["Setor"] == setor_sel)]
            grade_f = []
            for m in maqs_f:
                reg_m = df_sub_f[df_sub_f["Maquina_TAG"] == m]
                tipo = str(reg_m.iloc[0]["Tipo_Fuso"]) if not reg_m.empty and reg_m.iloc[0]["Tipo_Fuso"] in OPCOES_TIPO_FUSO else OPCOES_TIPO_FUSO[0]
                linha = {"Máquina": m, "Tipo de Fuso": tipo}
                for d in range(1, dias_no_mes + 1):
                    r_d = reg_m[reg_m["Dia"] == d]
                    linha[f"{d:02d}"] = int(r_d.iloc[0]["Quantidade_Quebras"]) if not r_d.empty else 0
                grade_f.append(linha)

            cfg = {
                "Máquina": st.column_config.TextColumn(disabled=True, width="small"),
                "Tipo de Fuso": st.column_config.SelectboxColumn(options=OPCOES_TIPO_FUSO, width="medium"),
            }
            for cd in cols_d:
                cfg[cd] = st.column_config.NumberColumn(cd, min_value=0, step=1, format="%d", width="small")

            ed_f = st.data_editor(
                pd.DataFrame(grade_f)[["Máquina", "Tipo de Fuso"] + cols_d],
                column_config=cfg,
                hide_index=True,
                use_container_width=True,
                height=420,
                key=f"ed_f_{ano_sel}_{setor_sel}_{m_nome}"
            )

            if st.button(f"💾 Salvar Apontamentos ({m_nome})", key=f"btn_s_{ano_sel}_{setor_sel}_{m_nome}", type="primary"):
                df_limpo = df_fusos[~((df_fusos["Ano"] == ano_sel) & (df_fusos["Mes"] == m_nome) & (df_fusos["Setor"] == setor_sel))]
                novos_f = []
                for _, r in ed_f.iterrows():
                    for d in range(1, dias_no_mes + 1):
                        novos_f.append({
                            "Ano": int(ano_sel), "Mes": m_nome, "Dia": d,
                            "Setor": setor_sel, "Maquina_TAG": r["Máquina"],
                            "Quantidade_Quebras": int(r[f"{d:02d}"]),
                            "Tipo_Fuso": str(r["Tipo de Fuso"]),
                        })
                df_final_f = pd.concat([df_limpo, pd.DataFrame(novos_f)], ignore_index=True)
                gerar_backup_seguro(ARQUIVO_FUSOS)
                df_final_f.to_excel(ARQUIVO_FUSOS, index=False)
                st.success(f"Apontamentos de {m_nome} salvos com sucesso!")
                st.rerun()

elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Máquinas")
