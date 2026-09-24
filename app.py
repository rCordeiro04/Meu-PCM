import calendar
import io
import os
import shutil
from datetime import date, datetime, timedelta
import altair as alt
import pandas as pd
import streamlit as st

# ==========================================
# CONFIGURAÇÃO DA PÁGINA & TEMA GLOBAL
# ==========================================
st.set_page_config(
    page_title="Portal PCM - Gestão de Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

def tema_altair():
    return {
        "config": {
            "view": {"strokeWidth": 0},
            "axis": {
                "domainColor": "#cbd5e1",
                "gridColor": "#f8fafc",
                "labelColor": "#64748b",
                "labelFontWeight": 600,
                "titleColor": "#0f172a",
                "titleFontWeight": 800,
                "tickColor": "#cbd5e1"
            },
            "legend": {
                "labelColor": "#475569",
                "titleColor": "#0f172a",
                "titleFontWeight": 800
            }
        }
    }
alt.themes.register("tema_pcm", tema_altair)
alt.themes.enable("tema_pcm")

# ==========================================
# CONSTANTES & ARQUIVOS
# ==========================================
ARQUIVO_FUSOS = "lancamentos_fusos_v5.xlsx"
ARQUIVO_CORREIAS = "lancamentos_correias_v4.xlsx"
ARQUIVO_PARADAS = "lancamentos_paradas_v1.xlsx"
ARQUIVO_PENDENCIAS = "lancamentos_pendencias_v1.xlsx"

COLUNAS_FUSOS = ["Ano", "Mes", "Dia", "Setor", "Maquina_TAG", "Quantidade_Quebras", "Tipo_Fuso"]
COLUNAS_CORREIAS = ["Setor", "Maquina_TAG", "Tipo_Correia_1", "Data_Instalacao_1", "Tipo_Correia_2", "Data_Instalacao_2"]
COLUNAS_PARADAS = ["Data", "Setor", "Maquina_TAG", "Tipo_Manutencao", "Descricao_Servico", "Tempo_Parado_Horas"]
COLUNAS_PENDENCIAS = ["Setor", "Maquina_TAG", "Nome_Servico", "Descricao_Pendencia", "Prioridade", "Status"]

OPCOES_TIPO_FUSO = ["FAG", "TEP", "M4BA", "MENEGATTO", "M4ZD", "USL"]

LISTA_MESES_PUROS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MAPA_MES_ABREV = {"Janeiro": "JAN", "Fevereiro": "FEV", "Março": "MAR", "Abril": "ABR", "Maio": "MAI", "Junho": "JUN", "Julho": "JUL", "Agosto": "AGO", "Setembro": "SET", "Outubro": "OUT", "Novembro": "NOV", "Dezembro": "DEZ"}
ORDEM_MESES_ABREV = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

DICIONARIO_SETORES = {
    "Setor A": [f"L-{i:02d}" for i in range(1, 29)],
    "Setor B": ["L-29", "L-30", "L-31", "L-32", "L-33", "L-34", "L-35", "L-36", "L-37", "L-38", "L-39", "L-40", "L-41", "L-42", "L-43", "L-44", "L-45", "L-46", "L-50", "L-51", "L-52", "L-53"],
    "Setor Látex": ["B-71", "B-72", "B-73", "B-74", "B-75", "B-76", "B-77", "B-78", "B-79", "B-80", "B-83", "B-84", "B-85", "B-86", "B-87", "B-88", "B-89", "B-102", "B-103", "B-104"],
    "Setor Menegatto": ["B-47", "B-48", "B-49", "B-81", "B-82", "B-93", "B-94", "B-95", "B-96", "B-97", "B-98", "B-99", "B-100", "B-101", "B-107", "B-108"],
}

# ==========================================
# FUNÇÕES DE SUPORTE & CACHE
# ==========================================
def obter_fuso_padrao(maq_tag, setor_nome):
    maq_str = str(maq_tag).strip().upper()
    MAQUINAS_FAG = [f"L-{i:02d}" for i in range(1, 29)] + ["L-52", "L-53", "B-47", "B-48", "B-49"]
    MAQUINAS_MENEGATTO = ["L-29", "L-30", "L-31", "L-35", "L-38", "L-50", "L-51", "L-41", "L-42", "L-43", "L-44", "L-45", "L-46"]
    
    if maq_str in MAQUINAS_FAG: return "FAG"
    if maq_str in MAQUINAS_MENEGATTO: return "MENEGATTO"
    
    if setor_nome == "Setor A": return "FAG"
    elif setor_nome == "Setor B": return "TEP"
    elif setor_nome == "Setor Látex": return "M4BA"
    return "MENEGATTO"

def gerar_backup_seguro(caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        os.makedirs("backups", exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arq = os.path.basename(caminho_arquivo)
        try: shutil.copy2(caminho_arquivo, os.path.join("backups", f"{ts}_{nome_arq}"))
        except: pass

def formatar_modelo(val):
    if val is None or pd.isna(val): return ""
    v_str = str(val).strip()
    if v_str.lower() in ["", "nan", "none", "nat"]: return ""
    if "." in v_str:
        partes = v_str.split(".")
        p_dec = partes[1][:3].ljust(3, "0")
        return f"{partes[0]}.{p_dec}"
    return v_str

@st.cache_data(show_spinner=False)
def carregar_dados():
    if os.path.exists(ARQUIVO_FUSOS):
        try:
            df_f = pd.read_excel(ARQUIVO_FUSOS)
            if "Dia" not in df_f.columns: df_f["Dia"] = 1
        except: df_f = pd.DataFrame(columns=COLUNAS_FUSOS)
    else:
        df_f = pd.DataFrame(columns=COLUNAS_FUSOS)
        df_f.to_excel(ARQUIVO_FUSOS, index=False)

    if not df_f.empty:
        def force_fuso(row):
            t = str(row.get("Tipo_Fuso", "")).strip()
            m = str(row.get("Maquina_TAG", "")).strip().upper()
            s = str(row.get("Setor", "")).strip()
            
            MAQUINAS_FAG = [f"L-{i:02d}" for i in range(1, 29)] + ["L-52", "L-53", "B-47", "B-48", "B-49"]
            MAQUINAS_MENEGATTO = ["L-29", "L-30", "L-31", "L-35", "L-38", "L-50", "L-51", "L-41", "L-42", "L-43", "L-44", "L-45", "L-46"]
            
            if m in MAQUINAS_FAG: return "FAG"
            if m in MAQUINAS_MENEGATTO: return "MENEGATTO"
            if t in OPCOES_TIPO_FUSO: return t
            return obter_fuso_padrao(m, s)

        novo_tipo = df_f.apply(force_fuso, axis=1)
        if not df_f["Tipo_Fuso"].equals(novo_tipo):
            df_f["Tipo_Fuso"] = novo_tipo
            df_f.to_excel(ARQUIVO_FUSOS, index=False)

    if os.path.exists(ARQUIVO_CORREIAS):
        try: df_c = pd.read_excel(ARQUIVO_CORREIAS)
        except: df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)
    else:
        df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)
        df_c.to_excel(ARQUIVO_CORREIAS, index=False)

    for col in COLUNAS_CORREIAS:
        if col not in df_c.columns: df_c[col] = ""
        df_c[col] = df_c[col].astype(object)

    df_c["Tipo_Correia_1"] = df_c["Tipo_Correia_1"].apply(formatar_modelo)
    df_c["Data_Instalacao_1"] = df_c["Data_Instalacao_1"].astype(str).replace({"nan": "", "NaT": "", "None": ""})
    df_c["Tipo_Correia_2"] = df_c["Tipo_Correia_2"].apply(formatar_modelo)
    df_c["Data_Instalacao_2"] = df_c["Data_Instalacao_2"].astype(str).replace({"nan": "", "NaT": "", "None": ""})

    if os.path.exists(ARQUIVO_PARADAS):
        try: df_p = pd.read_excel(ARQUIVO_PARADAS)
        except: df_p = pd.DataFrame(columns=COLUNAS_PARADAS)
    else:
        df_p = pd.DataFrame(columns=COLUNAS_PARADAS)
        df_p.to_excel(ARQUIVO_PARADAS, index=False)
    for col in COLUNAS_PARADAS:
        if col not in df_p.columns: df_p[col] = 0.0 if col == "Tempo_Parado_Horas" else ""

    if os.path.exists(ARQUIVO_PENDENCIAS):
        try: df_pend = pd.read_excel(ARQUIVO_PENDENCIAS)
        except: df_pend = pd.DataFrame(columns=COLUNAS_PENDENCIAS)
    else:
        df_pend = pd.DataFrame(columns=COLUNAS_PENDENCIAS)
        df_pend.to_excel(ARQUIVO_PENDENCIAS, index=False)
    for col in COLUNAS_PENDENCIAS:
        if col not in df_pend.columns: df_pend[col] = ""

    return df_f, df_c, df_p, df_pend

df_fusos, df_correias, df_paradas, df_pendencias = carregar_dados()

if "Nome_Servico" not in df_pendencias.columns:
    df_pendencias["Nome_Servico"] = ""

if "Amortecedores fusos" not in df_pendencias["Nome_Servico"].values:
    test_pend = []
    for m in DICIONARIO_SETORES["Setor A"]:
        status = "Concluído" if m in ["L-09", "L-20", "L-21"] else "Pendente"
        test_pend.append({
            "Setor": "Setor A",
            "Maquina_TAG": m,
            "Nome_Servico": "Amortecedores fusos",
            "Descricao_Pendencia": "troca dos amortecedores para diminuir indice de quebra de fusos",
            "Prioridade": "Alta",
            "Status": status
        })
    df_pendencias = pd.concat([df_pendencias, pd.DataFrame(test_pend)], ignore_index=True)
    df_pendencias.to_excel(ARQUIVO_PENDENCIAS, index=False)
    st.cache_data.clear()

def invalidar_cache():
    st.cache_data.clear()

def obter_maquinas_setor(setor_nome, df_c=None, df_f=None):
    base = set(DICIONARIO_SETORES.get(setor_nome, []))
    if df_c is not None and not df_c.empty: base.update(df_c[df_c["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique())
    if df_f is not None and not df_f.empty: base.update(df_f[df_f["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique())
    return sorted(list(base))

mapa_setor_maquina = {}
for s_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m: mapa_setor_maquina[m] = s_nome
for _, r in df_correias.iterrows():
    if pd.notna(r.get("Maquina_TAG")) and pd.notna(r.get("Setor")): mapa_setor_maquina[str(r["Maquina_TAG"]).strip()] = str(r["Setor"]).strip()
for _, r in df_fusos.iterrows():
    if pd.notna(r.get("Maquina_TAG")) and pd.notna(r.get("Setor")): mapa_setor_maquina[str(r["Maquina_TAG"]).strip()] = str(r["Setor"]).strip()

# ==========================================
# ESTADOS DA SESSÃO
# ==========================================
if "pagina_atual" not in st.session_state: st.session_state.pagina_atual = "Painel Fusos"
if "maq_clicada_cor" not in st.session_state: st.session_state.maq_clicada_cor = None
if "aba_setor_fuso" not in st.session_state: st.session_state.aba_setor_fuso = "Geral"

def navegar(p): st.session_state.pagina_atual = p

# ==========================================
# MOTOR DE REGRAS - CORREIAS
# ==========================================
data_hoje = date.today()
dados_maquinas, lista_correias_todas, lista_correias_novas, lista_correias_meia, lista_correias_criticas = {}, [], [], [], []

def avaliar_correia(dt_val):
    if not dt_val or str(dt_val).strip() in ["", "nan", "NaT", "None"]: return None, "Sem registro", "Sem histórico"
    try:
        dt_inst = pd.to_datetime(dt_val).date()
        dias = (data_hoje - dt_inst).days
        meses = round(dias / 30.4, 1)
        return (1 if dias <= 365 else 2 if dias <= 547 else 3), dt_inst.strftime("%d/%m/%Y"), f"{meses}m ({dias}d)"
    except: return None, str(dt_val), "Data inválida"

todas_maquinas_totais = []
for s_nome in DICIONARIO_SETORES.keys():
    for m in obter_maquinas_setor(s_nome, df_correias, df_fusos):
        if m not in todas_maquinas_totais: todas_maquinas_totais.append(m)

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

    dados_maquinas[maq_tag] = {"setor": setor_m, "t1": t1, "d1": d1_str, "uso1": t1_uso, "tem_c1": tem_c1, "t2": t2, "d2": d2_str, "uso2": t2_uso, "tem_c2": tem_c2, "status_label": status_label, "classe_card": classe_card, "dot": dot_simbolo}

# ==========================================
# CSS CUSTOMIZADO LAPIDADO (UI/UX PREMIUM)
# ==========================================
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800;900&display=swap');
        html, body, [class*="css"]  { font-family: 'Inter', sans-serif !important; }
        .block-container { padding: 3rem 2rem 1.5rem 2rem !important; max-width: 1400px; }
        [data-testid="stSidebar"] { background-color: #0b1120 !important; border-right: 1px solid #1e293b !important; }
        [data-testid="stSidebar"] h2, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span { color: #f8fafc; }
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] {
            background-color: transparent !important; color: #94a3b8 !important; border: 1px solid transparent !important;
            border-radius: 8px !important; font-weight: 600 !important; height: 42px !important; justify-content: flex-start; padding-left: 14px; transition: all 0.2s;
        }
        [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover { background-color: #1e293b !important; color: #f8fafc !important; }
        [data-testid="stSidebar"] .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #2563eb, #1d4ed8) !important; color: #ffffff !important;
            border: 1px solid transparent !important; border-radius: 8px !important; font-weight: 800 !important; height: 42px !important; justify-content: flex-start; padding-left: 14px; box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
        }
        div.stButton > button {
            background: #ffffff !important; color: #0f172a !important; border: 1px solid #e2e8f0 !important;
            padding: 0px 4px !important; font-size: 0.85rem !important; font-weight: 800 !important;
            height: 32px !important; min-height: 32px !important; line-height: 28px !important;
            border-radius: 6px !important; box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important; transition: all 0.15s ease-in-out !important;
        }
        div.stButton > button:hover { border-color: #3b82f6 !important; background: #f8fafc !important; color: #1d4ed8 !important; transform: translateY(-1px); box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important; }
        div[data-testid="column"] { padding: 0 6px !important; margin: 0px !important; }
        div[data-testid="stHorizontalBlock"] { gap: 0px !important; margin-bottom: 4px !important; }
        div[data-testid="stVegaLiteChart"] summary, div[data-testid="stVegaLiteChart"] .vega-actions { display: none !important; }
        .card-kpi-bonito {
            background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 10px 16px;
            display: flex; align-items: center; justify-content: space-between; height: 72px; box-sizing: border-box;
            position: relative; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.04); transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .card-kpi-bonito:hover { transform: translateY(-3px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.08); }
        .card-kpi-bonito::after { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 6px; }
        .card-kpi-bonito.c-total::after { background: #3b82f6; }
        .card-kpi-bonito.c-ok::after { background: #10b981; }
        .card-kpi-bonito.c-warn::after { background: #f59e0b; }
        .card-kpi-bonito.c-crit::after { background: #ef4444; }
        .kpi-val { font-size: 1.45rem; font-weight: 900; line-height: 1; font-family: 'Inter', sans-serif; color: #0f172a;}
        .kpi-lbl { font-size: 0.7rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }
        .alerta-manutencao {
            background: #fef2f2; border: 1px solid #fecaca; border-left: 6px solid #ef4444; border-radius: 8px;
            padding: 8px 14px; margin: 6px 0 12px 0; font-size: 0.85rem; font-weight: 600; color: #991b1b; box-shadow: 0 2px 4px rgba(239, 68, 68, 0.05);
        }
        .chip-critico { background: #fee2e2; border: 1px solid #fca5a5; color: #991b1b; padding: 2px 8px; border-radius: 6px; font-weight: 800; font-size: 0.78rem; }
        .hud-detalhe {
            background: #ffffff; border: 1px solid #cbd5e1; border-radius: 12px; padding: 12px 16px; margin: 4px 0 12px 0;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05); border-left: 6px solid #64748b; transition: all 0.2s;
        }
        .hud-detalhe.status-verde { border-left-color: #10b981; }
        .hud-detalhe.status-amarelo { border-left-color: #f59e0b; }
        .hud-detalhe.status-vermelho { border-left-color: #ef4444; }
        .hud-detalhe.status-cinza { border-left-color: #94a3b8; }
        .tag-pill { background: #f8fafc; border: 1px solid #e2e8f0; padding: 4px 12px; border-radius: 6px; font-size: 0.8rem; font-weight: 700; color: #334155; }
        .badge-status { padding: 4px 10px; border-radius: 6px; font-size: 0.75rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;}
        .badge-verde { background: #d1fae5; color: #065f46; }
        .badge-amarelo { background: #fef3c7; color: #92400e; }
        .badge-vermelho { background: #fee2e2; color: #991b1b; }
        .badge-cinza { background: #e2e8f0; color: #475569; }
        .pill-legenda { display: inline-flex; align-items: center; gap: 6px; font-size: 0.76rem; font-weight: 700; background: #ffffff; border: 1px solid #e2e8f0; padding: 4px 10px; border-radius: 16px; box-shadow: 0 1px 2px rgba(0,0,0,0.02);}
        .dot-legenda { width: 9px; height: 9px; border-radius: 50%; display: inline-block; }
        .header-setor-dash {
            font-size: 0.9rem; font-weight: 800; color: #0f172a; border-left: 4px solid #2563eb;
            padding-left: 10px; margin: 12px 0 6px 0; display: flex; align-items: center; justify-content: space-between;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='font-size:1.7rem; font-weight:900; margin:0 0 18px 0; color:#ffffff;'>⚙️ Portal PCM</h2>", unsafe_allow_html=True)
    
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#64748b; margin:8px 0 4px 0; letter-spacing:1px;'>PAINÉIS GERENCIAIS</p>", unsafe_allow_html=True)
    st.button("🔩 Painel de Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Fusos" else "secondary", on_click=navegar, args=("Painel Fusos",))
    st.button("🔄 Painel de Correias", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Correias" else "secondary", on_click=navegar, args=("Painel Correias",))
    st.button("🏭 Painel dos Setores", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Setores" else "secondary", on_click=navegar, args=("Painel Setores",))
    st.button("⚙️ Visão por Máquina", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Painel Maquinas" else "secondary", on_click=navegar, args=("Painel Maquinas",))

    st.markdown("<hr style='border-color:#1e293b; margin:15px 0;'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#64748b; margin:8px 0 4px 0; letter-spacing:1px;'>SISTEMA E DADOS</p>", unsafe_allow_html=True)
    st.button("🗄️ Banco de Dados", key="btn_nav_banco_dados", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Banco de Dados" else "secondary", on_click=navegar, args=("Banco de Dados",))

    st.markdown("<br><div style='text-align:center; font-size:0.7rem; color:#475569; font-weight:600;'>Portal PCM • Versão 1.0 Pro</div>", unsafe_allow_html=True)

# ==========================================
# ÁREA PRINCIPAL
# ==========================================
tela = st.session_state.pagina_atual

# ------------------------------------------
# 1. PAINEL DE CORREIAS
# ------------------------------------------
if tela == "Painel Correias":
    c_t, c_f, c_leg = st.columns([3.5, 2.0, 5.5])
    with c_t: st.markdown("<h2 style='margin:0; font-weight:900;'>🔄 Dashboard de Correias</h2>", unsafe_allow_html=True)
    with c_f:
        mods_un = sorted(list({r["modelo"] for r in lista_correias_todas if r["modelo"] and r["modelo"] != "Não informada"}))
        filtro_modelo = st.selectbox("Tipo", ["Todos os Tipos"] + mods_un, label_visibility="collapsed")
    with c_leg:
        st.markdown("""
            <div style="height:36px; display:flex; align-items:center; justify-content:flex-end; gap:8px;">
                <span class='pill-legenda'><span class='dot-legenda' style='background:#10b981;'></span> Nova (&le; 1a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#f59e0b;'></span> Meia (1-1.5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#ef4444;'></span> Urgente (&gt; 1.5a)</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#94a3b8;'></span> S/ Dados</span>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Correias</div><div class='kpi-val'>{len(lista_correias_todas)}</div></div><div style='font-size:1.8rem;'>📦</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Vida Útil OK</div><div class='kpi-val' style='color:#059669;'>{len(lista_correias_novas)}</div></div><div style='font-size:1.8rem;'>🟢</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Meia-Vida Ativa</div><div class='kpi-val' style='color:#d97706;'>{len(lista_correias_meia)}</div></div><div style='font-size:1.8rem;'>🟡</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Críticas / Urgentes</div><div class='kpi-val' style='color:#dc2626;'>{len(lista_correias_criticas)}</div></div><div style='font-size:1.8rem;'>🔴</div></div>", unsafe_allow_html=True)

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
                    <div style="display:flex; justify-content:space-between; margin-bottom:8px;">
                        <span class="tag-pill" style="background:#0f172a; color:#ffffff; font-size:0.85rem;">⚙️ <b>{sel['tag']}</b> &nbsp;|&nbsp; {sel['setor']}</span>
                        <span class="badge-status {b_cor}">{sel['status_label']}</span>
                    </div>
                    <div style="display:flex; gap:10px;">
                        <span class="tag-pill">🔼 <b>Superior/Cabeceira:</b> {sel['t1']} &nbsp;•&nbsp; {sel['d1']} &nbsp;•&nbsp; {sel['uso1']}</span>
                        <span class="tag-pill">🔽 <b>Inferior/Traseira:</b> {sel['t2']} &nbsp;•&nbsp; {sel['d2']} &nbsp;•&nbsp; {sel['uso2']}</span>
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
                    if formatar_modelo(ult_m.get("Tipo_Correia_1")) == filtro_modelo or formatar_modelo(ult_m.get("Tipo_Correia_2")) == filtro_modelo: filtradas.append(m)
            maquinas_do_setor = filtradas
        if not maquinas_do_setor: continue

        st.markdown(f"<div class='header-setor-dash'><span>🏭 {s_nome}</span> <span style='font-size:0.75rem; color:#64748b; font-weight:700;'>{len(maquinas_do_setor)} ativos vinculados</span></div>", unsafe_allow_html=True)
        cols_g = 14
        for chunk in [maquinas_do_setor[i:i + cols_g] for i in range(0, len(maquinas_do_setor), cols_g)]:
            cols = st.columns(cols_g)
            for i, m in enumerate(chunk):
                info_m = dados_maquinas.get(m, {})
                dot_m = info_m.get("dot", "⚪")
                if cols[i].button(f"{dot_m} {m}", key=f"btn_c_{m.replace('-', '_')}", use_container_width=True):
                    st.session_state.maq_clicada_cor = {"tag": m, **info_m}
                    st.rerun()

# ------------------------------------------
# 2. PAINEL DE FUSOS
# ------------------------------------------
elif tela == "Painel Fusos":
    cf_t, cf_a = st.columns([3.8, 1.4])
    cf_t.markdown("<h2 style='margin:0; font-weight:900;'>🔩 Dashboard de Fusos</h2>", unsafe_allow_html=True)
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

    df_ano_f = df_fusos[df_fusos["Ano"] == int(ano_f)].copy() if not df_fusos.empty else pd.DataFrame(columns=COLUNAS_FUSOS)

    if st.session_state.aba_setor_fuso == "Geral":
        tot_fabrica = int(df_ano_f["Quantidade_Quebras"].sum()) if not df_ano_f.empty else 0
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

        projecao_mes = "-"
        lbl_projecao = "Projeção Mês"
        if int(ano_f) == ano_atual_ref:
            nome_mes_atual = LISTA_MESES_PUROS[mes_atual_num_ref - 1]
            df_mes_atual = df_ano_f[df_ano_f["Mes"] == nome_mes_atual]
            quebras_mes_ate_hoje = int(df_mes_atual["Quantidade_Quebras"].sum()) if not df_mes_atual.empty else 0
            dias_passados = data_hoje_ref.day
            dias_totais_mes = calendar.monthrange(ano_atual_ref, mes_atual_num_ref)[1]

            if dias_passados > 0:
                projecao_mes = int(round((quebras_mes_ate_hoje / dias_passados) * dias_totais_mes, 0))
            lbl_projecao = f"Projeção ({MAPA_MES_ABREV[nome_mes_atual]})"

        kf1, kf2, kf3, kf4, kf5 = st.columns(5)
        kf1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Fábrica</div><div class='kpi-val'>{tot_fabrica}</div></div><div style='font-size:1.8rem;'>🔩</div></div>", unsafe_allow_html=True)
        kf2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal ({desc_divisor})</div><div class='kpi-val' style='color:#059669;'>{med_fabrica}</div></div><div style='font-size:1.8rem;'>📈</div></div>", unsafe_allow_html=True)
        kf3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Setor Crítico ({ult_mes_fab})</div><div class='kpi-val' style='color:#d97706; font-size:1.2rem;'>{setor_ofensor} ({qtd_setor_ofensor})</div></div><div style='font-size:1.8rem;'>🏭</div></div>", unsafe_allow_html=True)
        kf4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Quebras no Mês ({ult_mes_fab})</div><div class='kpi-val' style='color:#dc2626;'>{tot_ult_mes}</div></div><div style='font-size:1.8rem;'>🚨</div></div>", unsafe_allow_html=True)
        kf5.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>{lbl_projecao}</div><div class='kpi-val' style='color:#6366f1;'>{projecao_mes}</div></div><div style='font-size:1.8rem;'>🔮</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        def gerar_chart_setor(nome_s, cor_s):
            sub_s = df_ano_f[df_ano_f["Setor"] == nome_s].groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
            sub_s["Mes_Abrev"] = sub_s["Mes"].map(MAPA_MES_ABREV)
            tot_s = int(sub_s["Quantidade_Quebras"].sum())
            bars = alt.Chart(sub_s).mark_bar(color=cor_s, cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0)), y=alt.Y("Quantidade_Quebras:Q", title="Quebras"))
            txt = alt.Chart(sub_s).mark_text(dy=-8, fontSize=12, fontWeight=800).encode(x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV), y=alt.Y("Quantidade_Quebras:Q"), text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")))
            return (bars + txt).properties(height=240), tot_s

        cg1, cg2 = st.columns(2)
        cores_map = {"Setor A": "#3b82f6", "Setor B": "#f59e0b", "Setor Látex": "#10b981", "Setor Menegatto": "#ef4444"}

        with cg1:
            with st.container(border=True):
                c_a, t_a = gerar_chart_setor("Setor A", cores_map["Setor A"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:1rem; margin-bottom:10px;'><span>🏭 Setor A</span><span style='color:{cores_map['Setor A']}'>Total: {t_a} fusos</span></div>", unsafe_allow_html=True)
                st.altair_chart(c_a, use_container_width=True)
            with st.container(border=True):
                c_lat, t_lat = gerar_chart_setor("Setor Látex", cores_map["Setor Látex"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:1rem; margin-bottom:10px;'><span>🌿 Setor Látex</span><span style='color:{cores_map['Setor Látex']}'>Total: {t_lat} fusos</span></div>", unsafe_allow_html=True)
                st.altair_chart(c_lat, use_container_width=True)
        with cg2:
            with st.container(border=True):
                c_b, t_b = gerar_chart_setor("Setor B", cores_map["Setor B"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:1rem; margin-bottom:10px;'><span>🏭 Setor B</span><span style='color:{cores_map['Setor B']}'>Total: {t_b} fusos</span></div>", unsafe_allow_html=True)
                st.altair_chart(c_b, use_container_width=True)
            with st.container(border=True):
                c_men, t_men = gerar_chart_setor("Setor Menegatto", cores_map["Setor Menegatto"])
                st.markdown(f"<div style='display:flex; justify-content:space-between; font-weight:800; font-size:1rem; margin-bottom:10px;'><span>⚙️ Setor Menegatto</span><span style='color:{cores_map['Setor Menegatto']}'>Total: {t_men} fusos</span></div>", unsafe_allow_html=True)
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

        projecao_mes_s = "-"
        lbl_projecao_s = "Projeção Mês"
        if int(ano_f) == ano_atual_ref:
            nome_mes_atual = LISTA_MESES_PUROS[mes_atual_num_ref - 1]
            df_mes_atual_s = df_sa[df_sa["Mes"] == nome_mes_atual]
            quebras_mes_ate_hoje_s = int(df_mes_atual_s["Quantidade_Quebras"].sum()) if not df_mes_atual_s.empty else 0
            dias_passados = data_hoje_ref.day
            dias_totais_mes = calendar.monthrange(ano_atual_ref, mes_atual_num_ref)[1]

            if dias_passados > 0:
                projecao_mes_s = int(round((quebras_mes_ate_hoje_s / dias_passados) * dias_totais_mes, 0))
            lbl_projecao_s = f"Projeção ({MAPA_MES_ABREV[nome_mes_atual]})"

        ks1, ks2, ks3, ks4, ks5 = st.columns(5)
        ks1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Quebras ({s_ativo})</div><div class='kpi-val'>{tot_s}</div></div><div style='font-size:1.8rem;'>🔩</div></div>", unsafe_allow_html=True)
        ks2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal ({desc_divisor})</div><div class='kpi-val' style='color:#059669;'>{med_s}</div></div><div style='font-size:1.8rem;'>📅</div></div>", unsafe_allow_html=True)
        ks3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Tx Falha/Máq ({ult_mes_s})</div><div class='kpi-val' style='color:#d97706;'>{quebras_por_maq}</div></div><div style='font-size:1.8rem;'>⚙️</div></div>", unsafe_allow_html=True)
        ks4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Maior Ofensor ({ult_mes_s})</div><div class='kpi-val' style='color:#dc2626; font-size:1.2rem;'>{top_maq_s} ({qtd_top_s})</div></div><div style='font-size:1.8rem;'>⚠️</div></div>", unsafe_allow_html=True)
        ks5.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>{lbl_projecao_s}</div><div class='kpi-val' style='color:#6366f1;'>{projecao_mes_s}</div></div><div style='font-size:1.8rem;'>🔮</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

        c_evol, c_rosca = st.columns([1.5, 1.5])
        with c_evol:
            with st.container(border=True):
                st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>📊 Evolução de Quebras ({ano_f})</div>", unsafe_allow_html=True)
                df_evol = df_sa.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                df_evol["Mes_Abrev"] = df_evol["Mes"].map(MAPA_MES_ABREV)
                barras_s = alt.Chart(df_evol).mark_bar(color="#3b82f6", cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None), y=alt.Y("Quantidade_Quebras:Q", title="Quebras"), tooltip=["Mes", "Quantidade_Quebras"])
                rotulos_s = alt.Chart(df_evol).mark_text(dy=-8, fontSize=12, fontWeight=800).encode(x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV), y=alt.Y("Quantidade_Quebras:Q"), text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")))
                st.altair_chart((barras_s + rotulos_s).properties(height=280), use_container_width=True)

        with c_rosca:
            with st.container(border=True):
                cr_col1, cr_col2 = st.columns([2.0, 1.5])
                with cr_col1: st.markdown(f"<div style='font-size:1rem; font-weight:800;'>🍩 Distribuição por Tipo de Fuso</div>", unsafe_allow_html=True)
                with cr_col2: mes_filtro_rosca_fuso = st.selectbox("Mês:", ["Todos os Meses"] + LISTA_MESES_PUROS, key=f"sel_mes_rosca_fuso_{s_ativo}", label_visibility="collapsed")

                df_sa_rosca = df_sa.copy()
                if mes_filtro_rosca_fuso != "Todos os Meses": df_sa_rosca = df_sa_rosca[df_sa_rosca["Mes"] == mes_filtro_rosca_fuso]
                if not df_sa_rosca.empty and df_sa_rosca["Quantidade_Quebras"].sum() > 0:
                    df_tipos = df_sa_rosca[df_sa_rosca["Quantidade_Quebras"] > 0].groupby("Tipo_Fuso")["Quantidade_Quebras"].sum().reset_index()
                    base_don_fuso = alt.Chart(df_tipos).encode(
                        theta=alt.Theta("Quantidade_Quebras:Q", stack=True),
                        color=alt.Color("Tipo_Fuso:N", scale=alt.Scale(scheme="category10"), legend=alt.Legend(orient="right", title="Tipo de Fuso", labelFontSize=13, titleFontSize=14)),
                        tooltip=["Tipo_Fuso", "Quantidade_Quebras"]
                    )
                    arc_fuso = base_don_fuso.mark_arc(innerRadius=65, outerRadius=120, stroke="#ffffff", strokeWidth=2)
                    text_fuso = base_don_fuso.mark_text(radius=92, fontSize=13, fontWeight=800, fill="#ffffff").encode(text=alt.Text("Quantidade_Quebras:Q"))
                    st.altair_chart((arc_fuso + text_fuso).properties(height=280), use_container_width=True)
                else:
                    st.info(f"Sem registros para o período.")

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            ch_col1, ch_col2 = st.columns([2.5, 1.5])
            with ch_col1: st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>🔥 Mapa de Calor Operacional ({s_ativo})</div>", unsafe_allow_html=True)
            with ch_col2:
                tipos_disp = sorted(df_sa["Tipo_Fuso"].dropna().unique().tolist()) if not df_sa.empty else OPCOES_TIPO_FUSO
                fuso_filtro_mapa = st.selectbox("Filtrar Fuso no Mapa:", ["Todos os Tipos"] + tipos_disp, key=f"sel_fuso_mapa_{s_ativo}")

            df_mapa_calor = df_sa[df_sa["Tipo_Fuso"] == fuso_filtro_mapa] if fuso_filtro_mapa != "Todos os Tipos" else df_sa.copy()
            idx_grid = pd.MultiIndex.from_product([maqs_setor, LISTA_MESES_PUROS], names=["MAQ", "Mes"]).to_frame().reset_index(drop=True)
            idx_grid["MES"] = idx_grid["Mes"].map(MAPA_MES_ABREV)
            agrup_c = df_mapa_calor.groupby(["Maquina_TAG", "Mes"])["Quantidade_Quebras"].sum().reset_index()
            m_calor = pd.merge(idx_grid, agrup_c, left_on=["MAQ", "Mes"], right_on=["Maquina_TAG", "Mes"], how="left").fillna(0)

            rect = alt.Chart(m_calor).mark_rect(stroke="#fff", strokeWidth=1.5).encode(
                x=alt.X("MES:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(orient="top", labelAngle=0)),
                y=alt.Y("MAQ:N", sort=maqs_setor, title=None),
                color=alt.Color("Quantidade_Quebras:Q", scale=alt.Scale(domain=[0, 3, 8, 15], range=["#dcfce7", "#fef08a", "#f97316", "#ef4444"]), legend=alt.Legend(title="Quebras", orient="right")),
                tooltip=["MAQ", "MES", "Quantidade_Quebras"]
            )
            txt = alt.Chart(m_calor).mark_text(baseline="middle", fontSize=12, fontWeight=800).encode(
                x=alt.X("MES:N", sort=ORDEM_MESES_ABREV), y=alt.Y("MAQ:N", sort=maqs_setor),
                text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")),
                color=alt.condition("datum.Quantidade_Quebras >= 10", alt.value("#ffffff"), alt.value("#0f172a"))
            )
            st.altair_chart((rect + txt).properties(height=max(250, len(maqs_setor) * 26)), use_container_width=True)

# ------------------------------------------
# 3. PAINEL DE SETORES
# ------------------------------------------
elif tela == "Painel Setores":
    c_ts1, c_ts2, c_ts3 = st.columns([3.0, 1.5, 1.5])
    with c_ts2:
        setores_filtro_painel = ["Todos os Setores"] + list(DICIONARIO_SETORES.keys())
        setor_selecionado_exec = st.selectbox("Filtrar Setor:", setores_filtro_painel, label_visibility="collapsed")
    with c_ts3:
        mes_filtro_painel = st.selectbox("Mês:", ["Acumulado do Ano"] + LISTA_MESES_PUROS, label_visibility="collapsed")

    with c_ts1: 
        titulo_setor = f"Painel {setor_selecionado_exec}" if setor_selecionado_exec != "Todos os Setores" else "Painel Todos os Setores"
        st.markdown(f"<h2 style='margin:0; font-weight:900;'>🏭 {titulo_setor}</h2>", unsafe_allow_html=True)

    st.caption("Visão consolidada e comparativa de desempenho operacional por setor da fábrica.")

    setores_alvo_exec = list(DICIONARIO_SETORES.keys()) if setor_selecionado_exec == "Todos os Setores" else [setor_selecionado_exec]
    
    maquinas_alvo_totais = []
    for s_nome in setores_alvo_exec:
        maqs_s = obter_maquinas_setor(s_nome, df_correias, df_fusos)
        maquinas_alvo_totais.extend(maqs_s)

    tot_maqs_sel = len(maquinas_alvo_totais)

    # Lógica de Filtro de Tempo Dinâmico
    ano_ref = date.today().year
    mes_ref_num = date.today().month
    dia_ref = date.today().day

    df_paradas_chart = df_paradas.copy()
    df_paradas_chart['Data'] = pd.to_datetime(df_paradas_chart['Data'], errors='coerce')
    df_paradas_alvo = df_paradas_chart[df_paradas_chart['Maquina_TAG'].isin(maquinas_alvo_totais)]
    df_paradas_ano = df_paradas_alvo[df_paradas_alvo['Data'].dt.year == ano_ref]

    df_fusos_alvo = df_fusos[(df_fusos["Setor"].isin(setores_alvo_exec)) & (df_fusos["Ano"] == ano_ref)]

    if mes_filtro_painel == "Acumulado do Ano":
        dias_calculo = (date.today() - date(ano_ref, 1, 1)).days + 1
        if dias_calculo < 1: dias_calculo = 1
        horas_paradas_total = df_paradas_ano['Tempo_Parado_Horas'].sum()
        df_prev_filtro = df_paradas_ano[df_paradas_ano['Tipo_Manutencao'].astype(str).str.contains('Preventiva|Preventivo|Prev', case=False, na=False)]
        tot_fusos_sel = int(df_fusos_alvo["Quantidade_Quebras"].sum())
        df_paradas_filtro = df_paradas_ano
    else:
        num_mes_selecionado = LISTA_MESES_PUROS.index(mes_filtro_painel) + 1
        if num_mes_selecionado == mes_ref_num:
            dias_calculo = dia_ref
        elif num_mes_selecionado > mes_ref_num:
            dias_calculo = 0
        else:
            dias_calculo = calendar.monthrange(ano_ref, num_mes_selecionado)[1]
        
        df_paradas_mes = df_paradas_ano[df_paradas_ano['Data'].dt.month == num_mes_selecionado]
        horas_paradas_total = df_paradas_mes['Tempo_Parado_Horas'].sum()
        df_prev_filtro = df_paradas_mes[df_paradas_mes['Tipo_Manutencao'].astype(str).str.contains('Preventiva|Preventivo|Prev', case=False, na=False)]
        tot_fusos_sel = int(df_fusos_alvo[df_fusos_alvo["Mes"] == mes_filtro_painel]["Quantidade_Quebras"].sum())
        df_paradas_filtro = df_paradas_mes

    horas_totais_disponiveis = tot_maqs_sel * 24 * dias_calculo
    horas_operando_total = max(0, horas_totais_disponiveis - horas_paradas_total)
    total_h = horas_operando_total + horas_paradas_total
    perc_efi = (horas_operando_total / total_h * 100) if total_h > 0 else 0.0

    maquinas_com_prev = df_prev_filtro['Maquina_TAG'].nunique()
    perc_prev = (maquinas_com_prev / tot_maqs_sel * 100) if tot_maqs_sel > 0 else 0.0

    criticas_setor = [r for r in lista_correias_criticas if r["setor"] in setores_alvo_exec]
    tot_crit_sel = len(criticas_setor)
    
    tooltip_correias = "Máquinas Críticas:&#10;"
    if tot_crit_sel > 0:
        maqs_crit = {}
        for r in criticas_setor:
            if r["tag"] not in maqs_crit: maqs_crit[r["tag"]] = []
            maqs_crit[r["tag"]].append(r['pos'])
        for t, p_list in sorted(maqs_crit.items()):
            tooltip_correias += f"• {t} ({', '.join(p_list)})&#10;"
    else:
        tooltip_correias = "Nenhuma correia crítica neste setor."

    cs1, cs2, cs3, cs4 = st.columns(4)
    cs1.markdown(f"<div class='card-kpi-bonito c-ok' title='Eficiência calculada com base no período: {mes_filtro_painel}'><div><div class='kpi-lbl'>Eficiência Mecânica</div><div class='kpi-val' style='color:#059669;'>{perc_efi:.1f}%</div></div><div style='font-size:1.8rem;'>⏱️</div></div>", unsafe_allow_html=True)
    cs2.markdown(f"<div class='card-kpi-bonito c-total' title='Cobertura preventiva no período: {mes_filtro_painel}'><div><div class='kpi-lbl'>Cobertura Preventiva</div><div class='kpi-val' style='color:#3b82f6;'>{perc_prev:.1f}%</div></div><div style='font-size:1.8rem;'>🛠️</div></div>", unsafe_allow_html=True)
    cs3.markdown(f"<div class='card-kpi-bonito c-warn' title='Total registrado em: {mes_filtro_painel}'><div><div class='kpi-lbl'>Quebras de Fusos</div><div class='kpi-val' style='color:#d97706;'>{tot_fusos_sel}</div></div><div style='font-size:1.8rem;'>🔩</div></div>", unsafe_allow_html=True)
    cs4.markdown(f"<div class='card-kpi-bonito c-crit' title='{tooltip_correias}'><div><div class='kpi-lbl'>Correias Críticas</div><div class='kpi-val' style='color:#dc2626;'>{tot_crit_sel}</div></div><div style='font-size:1.8rem;'>🚨</div></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    if setor_selecionado_exec == "Todos os Setores":
        with st.container(border=True):
            if mes_filtro_painel == "Acumulado do Ano":
                st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>🔩 Evolução Mensal de Fusos — {setor_selecionado_exec}</div>", unsafe_allow_html=True)
                df_fusos_setor = df_fusos[(df_fusos["Setor"].isin(setores_alvo_exec)) & (df_fusos["Ano"] == ano_ref)].copy()
                df_f_evol = df_fusos_setor.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                df_f_evol["Mes_Abrev"] = df_f_evol["Mes"].map(MAPA_MES_ABREV)
                
                barras_setor = alt.Chart(df_f_evol).mark_bar(color="#3b82f6", cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                    tooltip=["Mes", "Quantidade_Quebras"]
                )
                rotulos_setor = alt.Chart(df_f_evol).mark_text(dy=-8, fontSize=12, fontWeight=800).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV), 
                    y=alt.Y("Quantidade_Quebras:Q"), 
                    text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                )
                st.altair_chart((barras_setor + rotulos_setor).properties(height=280), use_container_width=True)
            else:
                st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>🔩 Evolução Diária de Fusos ({mes_filtro_painel}) — {setor_selecionado_exec}</div>", unsafe_allow_html=True)
                df_fusos_setor_mes = df_fusos[(df_fusos["Setor"].isin(setores_alvo_exec)) & (df_fusos["Ano"] == ano_ref) & (df_fusos["Mes"] == mes_filtro_painel)].copy()
                
                num_mes_selecionado = LISTA_MESES_PUROS.index(mes_filtro_painel) + 1
                _, dias_no_mes = calendar.monthrange(ano_ref, num_mes_selecionado)
                lista_dias = list(range(1, dias_no_mes + 1))
                
                if df_fusos_setor_mes.empty:
                    df_f_evol_dia = pd.DataFrame({"Dia": lista_dias, "Quantidade_Quebras": 0})
                else:
                    df_f_evol_dia = df_fusos_setor_mes.groupby("Dia")["Quantidade_Quebras"].sum().reindex(lista_dias, fill_value=0).reset_index()
                
                df_f_evol_dia["Dia_Str"] = df_f_evol_dia["Dia"].astype(str)
                
                barras_dia = alt.Chart(df_f_evol_dia).mark_bar(color="#3b82f6", cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
                    x=alt.X("Dia_Str:N", sort=[str(d) for d in lista_dias], title="Dia do Mês", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                    tooltip=["Dia", "Quantidade_Quebras"]
                )
                rotulos_dia = alt.Chart(df_f_evol_dia).mark_text(dy=-8, fontSize=11, fontWeight=800).encode(
                    x=alt.X("Dia_Str:N", sort=[str(d) for d in lista_dias]), 
                    y=alt.Y("Quantidade_Quebras:Q"), 
                    text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                )
                st.altair_chart((barras_dia + rotulos_dia).properties(height=280), use_container_width=True)

            # =========================================================================
            # NOVA SESSÃO: PORCENTAGEM DE EFICIÊNCIA DE CADA SETOR SEPARADO (BEM BONITO)
            # =========================================================================
            st.markdown("""
                <style>
                .card-efi-setor {
                    background: linear-gradient(145deg, #ffffff, #f8fafc); border: 1px solid #e2e8f0; border-radius: 12px; padding: 18px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); transition: all 0.2s ease-in-out;
                }
                .card-efi-setor:hover {
                    transform: translateY(-4px); box-shadow: 0 12px 20px -5px rgba(0,0,0,0.08); border-color: #cbd5e1;
                }
                </style>
                <div style='height:8px;'></div>
                <hr style='border: 0; border-top: 1px dashed #cbd5e1; margin: 15px 0;'>
                <div style='font-size:1.15rem; font-weight:900; margin-bottom:16px; color:#0f172a; display:flex; align-items:center; gap:8px;'>
                    ⚡ Eficiência Mecânica Detalhada por Setor
                </div>
            """, unsafe_allow_html=True)

            cols_efi = st.columns(len(DICIONARIO_SETORES))
            
            for idx, s_nome in enumerate(DICIONARIO_SETORES.keys()):
                maqs_s = obter_maquinas_setor(s_nome, df_correias, df_fusos)
                tot_maqs_s = len(maqs_s)
                
                df_paradas_s = df_paradas_filtro[df_paradas_filtro['Maquina_TAG'].isin(maqs_s)]
                hp_s = df_paradas_s['Tempo_Parado_Horas'].sum() if not df_paradas_s.empty else 0.0
                
                hd_s = tot_maqs_s * 24 * dias_calculo
                ho_s = max(0, hd_s - hp_s)
                th_s = ho_s + hp_s
                efi_s = (ho_s / th_s * 100) if th_s > 0 else 0.0
                
                cor_borda = "#10b981" if efi_s >= 95 else "#f59e0b" if efi_s >= 90 else "#ef4444"
                icone_status = "🟢" if efi_s >= 95 else "🟡" if efi_s >= 90 else "🔴"
                
                with cols_efi[idx]:
                    st.markdown(f"""
                        <div class="card-efi-setor" style="border-bottom: 4px solid {cor_borda};">
                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                                <span style="font-size:0.85rem; font-weight:800; color:#475569; text-transform:uppercase; letter-spacing:0.5px;">{s_nome}</span>
                                <span>{icone_status}</span>
                            </div>
                            <div style="font-size:2rem; font-weight:900; color:#0f172a; margin-bottom:12px; line-height:1;">
                                {efi_s:.1f}<span style="font-size:1.2rem; color:#64748b;">%</span>
                            </div>
                            <div style="width:100%; background-color:#e2e8f0; border-radius:8px; height:8px; overflow:hidden; margin-bottom:8px;">
                                <div style="width:{efi_s}%; background-color:{cor_borda}; height:100%; border-radius:8px; transition: width 1s ease-in-out;"></div>
                            </div>
                            <div style="display:flex; justify-content:space-between; font-size:0.75rem; font-weight:700; color:#64748b;">
                                <span>{tot_maqs_s} máqs</span>
                                <span>{hp_s:.1f}h paradas</span>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            # =========================================================================

    else:
        c_fuso_evol, c_top_efi = st.columns([20, 15])
        with c_fuso_evol:
            with st.container(border=True):
                if mes_filtro_painel == "Acumulado do Ano":
                    st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>🔩 Evolução Mensal de Fusos — {setor_selecionado_exec}</div>", unsafe_allow_html=True)
                    df_fusos_setor = df_fusos[(df_fusos["Setor"].isin(setores_alvo_exec)) & (df_fusos["Ano"] == ano_ref)].copy()
                    df_f_evol = df_fusos_setor.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                    df_f_evol["Mes_Abrev"] = df_f_evol["Mes"].map(MAPA_MES_ABREV)
                    
                    barras_setor = alt.Chart(df_f_evol).mark_bar(color="#3b82f6", cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                        x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0)),
                        y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                        tooltip=["Mes", "Quantidade_Quebras"]
                    )
                    rotulos_setor = alt.Chart(df_f_evol).mark_text(dy=-8, fontSize=12, fontWeight=800).encode(
                        x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV), 
                        y=alt.Y("Quantidade_Quebras:Q"), 
                        text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                    )
                    st.altair_chart((barras_setor + rotulos_setor).properties(height=280), use_container_width=True)
                else:
                    st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>🔩 Evolução Diária de Fusos ({mes_filtro_painel}) — {setor_selecionado_exec}</div>", unsafe_allow_html=True)
                    df_fusos_setor_mes = df_fusos[(df_fusos["Setor"].isin(setores_alvo_exec)) & (df_fusos["Ano"] == ano_ref) & (df_fusos["Mes"] == mes_filtro_painel)].copy()
                    
                    num_mes_selecionado = LISTA_MESES_PUROS.index(mes_filtro_painel) + 1
                    _, dias_no_mes = calendar.monthrange(ano_ref, num_mes_selecionado)
                    lista_dias = list(range(1, dias_no_mes + 1))
                    
                    if df_fusos_setor_mes.empty:
                        df_f_evol_dia = pd.DataFrame({"Dia": lista_dias, "Quantidade_Quebras": 0})
                    else:
                        df_f_evol_dia = df_fusos_setor_mes.groupby("Dia")["Quantidade_Quebras"].sum().reindex(lista_dias, fill_value=0).reset_index()
                    
                    df_f_evol_dia["Dia_Str"] = df_f_evol_dia["Dia"].astype(str)
                    
                    barras_dia = alt.Chart(df_f_evol_dia).mark_bar(color="#3b82f6", cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
                        x=alt.X("Dia_Str:N", sort=[str(d) for d in lista_dias], title="Dia do Mês", axis=alt.Axis(labelAngle=0)),
                        y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                        tooltip=["Dia", "Quantidade_Quebras"]
                    )
                    rotulos_dia = alt.Chart(df_f_evol_dia).mark_text(dy=-8, fontSize=11, fontWeight=800).encode(
                        x=alt.X("Dia_Str:N", sort=[str(d) for d in lista_dias]), 
                        y=alt.Y("Quantidade_Quebras:Q"), 
                        text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                    )
                    st.altair_chart((barras_dia + rotulos_dia).properties(height=280), use_container_width=True)
        
        with c_top_efi:
            with st.container(border=True):
                st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:15px;'>📉 Top 5 - Piores Eficiências</div>", unsafe_allow_html=True)
                
                df_paradas_maq = df_paradas_filtro.groupby('Maquina_TAG')['Tempo_Parado_Horas'].sum().to_dict()
                efi_list = []
                for maq in maquinas_alvo_totais:
                    hp = df_paradas_maq.get(maq, 0.0)
                    hd = 24 * dias_calculo
                    efi = max(0.0, ((hd - hp) / hd) * 100) if hd > 0 else 0.0
                    efi_list.append({"Maquina": maq, "Eficiencia": efi, "Horas_Paradas": hp})
                
                df_top = pd.DataFrame(efi_list).sort_values(by=["Eficiencia", "Horas_Paradas"], ascending=[True, False]).head(5)

                if not df_top.empty and dias_calculo > 0:
                    for _, row in df_top.iterrows():
                        maq_t = row['Maquina']
                        efi_t = row['Eficiencia']
                        hp_t = row['Horas_Paradas']
                        cor_barra = "#ef4444" if efi_t < 90 else ("#f59e0b" if efi_t < 98 else "#10b981")
                        st.markdown(f"""
                            <div style="margin-bottom:14px;">
                                <div style="display:flex; justify-content:space-between; font-size:0.85rem; font-weight:800; color:#1e293b; margin-bottom:4px;">
                                    <span>⚙️ {maq_t} <span style="font-size:0.75rem; color:#64748b; font-weight:600;">({hp_t:.1f}h)</span></span>
                                    <span>{efi_t:.1f}%</span>
                                </div>
                                <div style="width:100%; background-color:#e2e8f0; border-radius:4px; height:8px;">
                                    <div style="width:{efi_t}%; background-color:{cor_barra}; height:8px; border-radius:4px;"></div>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                else:
                    st.info("Sem dados operacionais.")

    if setor_selecionado_exec != "Todos os Setores":
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(f"<div style='font-size:1.2rem; font-weight:900; margin-bottom:15px; color:#0f172a;'>📋 Serviços em Andamento — {setor_selecionado_exec}</div>", unsafe_allow_html=True)

            df_pend_setor = df_pendencias[df_pendencias["Setor"] == setor_selecionado_exec].copy()
            
            servicos_ativos = []
            if not df_pend_setor.empty:
                df_pend_setor["Nome_Servico"] = df_pend_setor["Nome_Servico"].fillna("Serviço sem título")
                df_pend_setor["Descricao_Pendencia"] = df_pend_setor["Descricao_Pendencia"].fillna("")
                for (nome_serv, desc), group in df_pend_setor.groupby(['Nome_Servico', 'Descricao_Pendencia']):
                    pendentes = group[~group['Status'].astype(str).str.lower().str.contains('conclu')]['Maquina_TAG'].tolist()
                    concluidas = group[group['Status'].astype(str).str.lower().str.contains('conclu')]['Maquina_TAG'].tolist()
                    if pendentes or concluidas:
                        servicos_ativos.append({"nome": nome_serv, "desc": desc, "pendentes": sorted(pendentes), "concluidas": sorted(concluidas)})

            if servicos_ativos:
                for item in servicos_ativos:
                    p_str = ", ".join(item["pendentes"]) if item["pendentes"] else "Nenhuma"
                    c_str = ", ".join(item["concluidas"]) if item["concluidas"] else "Nenhuma"

                    st.markdown(f"""
                        <div style="background: linear-gradient(to right, #ffffff, #f8fafc); border: 1px solid #cbd5e1; border-left: 8px solid #f59e0b; border-radius: 12px; padding: 20px; margin-bottom: 16px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);">
                            <div style="font-weight: 900; color: #0f172a; font-size: 1.3rem; margin-bottom: 6px; display: flex; align-items: center; gap: 8px;">
                                🛠️ {item['nome']}
                            </div>
                            <div style="font-size: 1rem; color: #475569; font-style: italic; margin-bottom: 16px; border-bottom: 1px dashed #cbd5e1; padding-bottom: 12px;">
                                {item['desc']}
                            </div>
                            <div style="display: flex; flex-direction: column; gap: 10px;">
                                <div style="font-size: 1rem; color: #334155; display: flex; align-items: center; flex-wrap: wrap; gap: 8px;">
                                    <span style="background: #fee2e2; color: #b91c1c; padding: 4px 12px; border-radius: 8px; font-weight: 800; font-size: 0.95rem;">
                                        ⏳ Pendentes ({len(item['pendentes'])})
                                    </span> 
                                    <span style="font-weight: 600;">{p_str}</span>
                                </div>
                                <div style="font-size: 1rem; color: #334155; display: flex; align-items: center; flex-wrap: wrap; gap: 8px;">
                                    <span style="background: #d1fae5; color: #047857; padding: 4px 12px; border-radius: 8px; font-weight: 800; font-size: 0.95rem;">
                                        ✅ Prontas ({len(item['concluidas'])})
                                    </span> 
                                    <span style="font-weight: 600;">{c_str}</span>
                                </div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("✅ Nenhum serviço pendente ou em andamento neste setor.")

# ------------------------------------------
# 4. PAINEL DE MÁQUINAS
# ------------------------------------------
elif tela == "Painel Maquinas":
    c_m_t, c_m_prt = st.columns([85, 15])
    with c_m_t:
        st.markdown("<h2 style='margin:0; font-weight:900;'>⚙️ Prontuário Individual da Máquina</h2>", unsafe_allow_html=True)
        st.caption("Consulte o histórico detalhado, manutenções, pendências e quebras de fusos por TAG.")
    with c_m_prt:
        modo_limpo = st.toggle("🖨️ Tela Limpa", key="tgl_maq")
        if modo_limpo: st.markdown("<style>[data-testid='stSidebar'] {display: none !important;} header[data-testid='stHeader'] {display: none !important;} .block-container {padding-top: 1rem !important; max-width: 100% !important;}</style>", unsafe_allow_html=True)

    col_sm, col_mq, col_mes = st.columns([15, 20, 15])
    with col_sm: setor_selecionado_maq = st.selectbox("Filtrar por Setor:", list(DICIONARIO_SETORES.keys()), key="sel_maq_painel_set")
    maqs_disponiveis = obter_maquinas_setor(setor_selecionado_maq, df_correias, df_fusos)
    
    if maqs_disponiveis:
        with col_mq: tag_selecionada = st.selectbox("Selecione a TAG da Máquina:", maqs_disponiveis, key="sel_maq_painel_tag")
        with col_mes: mes_filtro_maq = st.selectbox("Período:", ["Acumulado do Ano"] + LISTA_MESES_PUROS, key="sel_mes_maq")

        if tag_selecionada:
            info_cor_maq = dados_maquinas.get(tag_selecionada, {})
            
            ano_ref = date.today().year
            mes_ref_num = date.today().month
            dia_ref = date.today().day
            
            sub_fusos_maq = df_fusos[(df_fusos["Maquina_TAG"] == tag_selecionada) & (df_fusos["Ano"] == ano_ref)]
            
            sub_paradas_maq = df_paradas[df_paradas["Maquina_TAG"] == tag_selecionada].copy()
            sub_paradas_maq['Data_Parsed'] = pd.to_datetime(sub_paradas_maq['Data'], errors='coerce')
            
            if mes_filtro_maq == "Acumulado do Ano":
                dias_calculo = (date.today() - date(ano_ref, 1, 1)).days + 1
                if dias_calculo < 1: dias_calculo = 1
                df_paradas_filtro = sub_paradas_maq[sub_paradas_maq['Data_Parsed'].dt.year == ano_ref]
                df_fusos_filtro = sub_fusos_maq
                lbl_horas = "Horas Paradas (Ano)"
            else:
                num_mes_sel = LISTA_MESES_PUROS.index(mes_filtro_maq) + 1
                if num_mes_sel == mes_ref_num:
                    dias_calculo = dia_ref
                elif num_mes_sel > mes_ref_num:
                    dias_calculo = 0
                else:
                    dias_calculo = calendar.monthrange(ano_ref, num_mes_sel)[1]
                
                df_paradas_filtro = sub_paradas_maq[(sub_paradas_maq['Data_Parsed'].dt.year == ano_ref) & (sub_paradas_maq['Data_Parsed'].dt.month == num_mes_sel)]
                df_fusos_filtro = sub_fusos_maq[sub_fusos_maq["Mes"] == mes_filtro_maq]
                lbl_horas = f"Horas Paradas ({MAPA_MES_ABREV[mes_filtro_maq]})"
            
            tot_horas_paradas = float(df_paradas_filtro["Tempo_Parado_Horas"].sum()) if not df_paradas_filtro.empty else 0.0

            horas_totais_disp_maq = 24 * dias_calculo
            efi_maq_perc = max(0.0, ((horas_totais_disp_maq - tot_horas_paradas) / horas_totais_disp_maq) * 100) if horas_totais_disp_maq > 0 else 0.0

            mod_sup = info_cor_maq.get('t1', '')
            mod_inf = info_cor_maq.get('t2', '')
            if mod_sup and mod_inf and mod_sup != "Não informada" and mod_inf != "Não informada":
                modelos_str = mod_sup if mod_sup == mod_inf else f"{mod_sup} | {mod_inf}"
            elif mod_sup and mod_sup != "Não informada": modelos_str = mod_sup
            elif mod_inf and mod_inf != "Não informada": modelos_str = mod_inf
            else: modelos_str = "S/ Modelo"
            
            condicao_cor = info_cor_maq.get('status_label', 'S/ Dados')
            dot_cor = info_cor_maq.get('dot', '⚪')

            df_prev_maq = sub_paradas_maq[sub_paradas_maq['Tipo_Manutencao'].astype(str).str.contains('Preventiva|Preventivo|Prev', case=False, na=False)]
            if not df_prev_maq.empty:
                ultima_prev_date = df_prev_maq['Data_Parsed'].max()
                str_ultima_prev = ultima_prev_date.strftime("%d/%m/%Y") if pd.notnull(ultima_prev_date) else "Sem registro"
            else:
                str_ultima_prev = "Sem registro"

            cm1, cm2, cm3, cm4 = st.columns(4)
            lbl_efi = f"Eficiência ({MAPA_MES_ABREV[mes_filtro_maq]})" if mes_filtro_maq != "Acumulado do Ano" else "Eficiência (Ano)"
            
            cm1.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>{lbl_efi}</div><div class='kpi-val' style='color:#059669;'>{efi_maq_perc:.1f}%</div></div><div style='font-size:1.8rem;'>⏱️</div></div>", unsafe_allow_html=True)
            cm2.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Correia: {modelos_str}</div><div class='kpi-val' style='font-size:1.1rem;'>{dot_cor} {condicao_cor}</div></div><div style='font-size:1.8rem;'>🔄</div></div>", unsafe_allow_html=True)
            cm3.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>{lbl_horas}</div><div class='kpi-val' style='color:#dc2626;'>{round(tot_horas_paradas, 1)}h</div></div><div style='font-size:1.8rem;'>🛑</div></div>", unsafe_allow_html=True)
            cm4.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Última Preventiva</div><div class='kpi-val' style='font-size:1.1rem; color:#3b82f6;'>{str_ultima_prev}</div></div><div style='font-size:1.8rem;'>🛠️</div></div>", unsafe_allow_html=True)

            st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

            sub_pend_maq = df_pendencias[(df_pendencias["Maquina_TAG"] == tag_selecionada) & (df_pendencias["Status"].astype(str).str.lower() != "concluído")]
            with st.container(border=True):
                st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>📋 Manutenções Pendentes — {tag_selecionada}</div>", unsafe_allow_html=True)
                if not sub_pend_maq.empty: st.dataframe(sub_pend_maq[["Nome_Servico", "Descricao_Pendencia", "Prioridade", "Status"]], use_container_width=True, hide_index=True)
                else: st.info("Nenhuma manutenção pendente cadastrada.")

            st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

            with st.container(border=True):
                st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>🛠️ Corretivas Executadas ({mes_filtro_maq}) — {tag_selecionada}</div>", unsafe_allow_html=True)
                if not df_paradas_filtro.empty: st.dataframe(df_paradas_filtro[["Data", "Tipo_Manutencao", "Descricao_Servico", "Tempo_Parado_Horas"]].sort_values("Data", ascending=False), use_container_width=True, hide_index=True)
                else: st.info("Nenhuma manutenção corretiva registrada no período.")

            st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

            with st.container(border=True):
                if mes_filtro_maq == "Acumulado do Ano":
                    st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>📊 Evolução de Quebras de Fusos — {tag_selecionada}</div>", unsafe_allow_html=True)
                    if not df_fusos_filtro.empty:
                        df_f_maq = df_fusos_filtro.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                        df_f_maq["Mes_Abrev"] = df_f_maq["Mes"].map(MAPA_MES_ABREV)
                        bar_maq = alt.Chart(df_f_maq).mark_bar(color="#3b82f6", cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title="Mês"), y=alt.Y("Quantidade_Quebras:Q", title="Quebras"))
                        txt_maq = bar_maq.mark_text(dy=-6, fontSize=12, fontWeight=800).encode(text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")))
                        st.altair_chart((bar_maq + txt_maq).properties(height=260), use_container_width=True)
                    else:
                        st.info("Sem registros de quebras de fusos para esta máquina no ano.")
                else:
                    st.markdown(f"<div style='font-size:1rem; font-weight:800; margin-bottom:10px;'>📊 Evolução Diária de Fusos ({mes_filtro_maq}) — {tag_selecionada}</div>", unsafe_allow_html=True)
                    if not df_fusos_filtro.empty:
                        num_mes_sel = LISTA_MESES_PUROS.index(mes_filtro_maq) + 1
                        lista_dias = list(range(1, calendar.monthrange(ano_ref, num_mes_sel)[1] + 1))
                        
                        df_fusos_filtro_grouped = df_fusos_filtro.groupby("Dia")["Quantidade_Quebras"].sum()
                        if df_fusos_filtro_grouped.empty:
                            df_f_maq_dia = pd.DataFrame({"Dia": lista_dias, "Quantidade_Quebras": 0})
                        else:
                            df_f_maq_dia = df_fusos_filtro_grouped.reindex(lista_dias, fill_value=0).reset_index()
                        
                        df_f_maq_dia["Dia_Str"] = df_f_maq_dia["Dia"].astype(str)
                        bar_maq_dia = alt.Chart(df_f_maq_dia).mark_bar(color="#3b82f6", cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
                            x=alt.X("Dia_Str:N", sort=[str(d) for d in lista_dias], title="Dia do Mês", axis=alt.Axis(labelAngle=0)),
                            y=alt.Y("Quantidade_Quebras:Q", title="Quebras")
                        )
                        txt_maq_dia = bar_maq_dia.mark_text(dy=-6, fontSize=11, fontWeight=800).encode(text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")))
                        st.altair_chart((bar_maq_dia + txt_maq_dia).properties(height=260), use_container_width=True)
                    else:
                        st.info("Sem registros de quebras de fusos para esta máquina no período selecionado.")
    else:
        st.info("Nenhuma máquina encontrada neste setor.")

# ------------------------------------------
# 5. BANCO DE DADOS & GESTÃO
# ------------------------------------------
elif tela == "Banco de Dados":
    st.title("🗄️ Banco de Dados & Sistema")
    st.caption("Central de exportação/importação de dados e gestão integral do parque de máquinas da fábrica.")

    tab_fusos_db, tab_correias_db, tab_paradas_db, tab_pendencias_db, tab_maquinas_db, tab_backups_db = st.tabs([
        "🔩 Base de Fusos", "🔄 Base de Correias", "🛠️ Corretivas & Paradas", "📋 Manutenções Pendentes", "🏭 Ativos & Máquinas", "🛡️ Backups"
    ])

    with tab_fusos_db:
        st.markdown("### 📅 Gestão de Fusos Anual Consolidada")
        c_ano_db, c_set_db = st.columns([15, 25])
        with c_ano_db: ano_db_fuso = st.selectbox("Ano de Trabalho:", [2024, 2025, 2026, 2027], index=2, key="sel_ano_db_fuso")
        with c_set_db: setor_db_fuso = st.selectbox("Setor:", list(DICIONARIO_SETORES.keys()), key="sel_setor_db_fuso")

        maquinas_set_db = obter_maquinas_setor(setor_db_fuso, df_correias, df_fusos)

        buffer_excel_ano = io.BytesIO()
        with pd.ExcelWriter(buffer_excel_ano, engine="openpyxl") as writer:
            for idx_m, nome_mes_aba in enumerate(LISTA_MESES_PUROS):
                num_mes = idx_m + 1
                _, dias_no_mes = calendar.monthrange(int(ano_db_fuso), num_mes)
                cols_dias_aba = [str(d) for d in range(1, dias_no_mes + 1)]
                df_mes_fuso = df_fusos[(df_fusos["Ano"] == int(ano_db_fuso)) & (df_fusos["Mes"] == nome_mes_aba) & (df_fusos["Setor"] == setor_db_fuso)]
                grade_aba = []
                for maq in maquinas_set_db:
                    sub_maq = df_mes_fuso[df_mes_fuso["Maquina_TAG"] == maq]
                    linha_aba = {"MAQUINA": maq}
                    for d in range(1, dias_no_mes + 1):
                        sub_d = sub_maq[sub_maq["Dia"] == d]
                        linha_aba[str(d)] = int(sub_d.iloc[0]["Quantidade_Quebras"]) if not sub_d.empty else 0
                    grade_aba.append(linha_aba)
                pd.DataFrame(grade_aba)[["MAQUINA"] + cols_dias_aba].to_excel(writer, index=False, sheet_name=nome_mes_aba[:31])
        buffer_excel_ano.seek(0)

        st.markdown("---")
        col_down_ano, col_up_ano = st.columns([15, 25])

        with col_down_ano:
            st.markdown("#### 📥 Descarregar Livro")
            st.download_button(f"📥 Baixar Ano {ano_db_fuso}", data=buffer_excel_ano, file_name=f"fusos_{setor_db_fuso.replace(' ', '_')}_{ano_db_fuso}.xlsx", use_container_width=True)

        with col_up_ano:
            st.markdown("#### 📤 Importar Livro")
            upload_ano = st.file_uploader(f"Enviar ficheiro {ano_db_fuso} (.xlsx)", type=["xlsx"], key=f"upload_ano_db_{ano_db_fuso}_{setor_db_fuso}")
            if upload_ano is not None:
                if st.button(f"Atualizar Ano {ano_db_fuso}", type="primary"):
                    try:
                        excel_importado = pd.ExcelFile(upload_ano)
                        abas_encontradas = excel_importado.sheet_names
                        novos_registros_ano, meses_atualizados = [], []

                        for nome_mes_oficial in LISTA_MESES_PUROS:
                            aba_alvo = next((sh for sh in abas_encontradas if sh.strip().lower() == nome_mes_oficial.lower()), None)
                            if aba_alvo:
                                df_aba = pd.read_excel(excel_importado, sheet_name=aba_alvo)
                                col_maq = next((cand for cand in ["MAQUINA", "Máquina", "Maquina", "Maquina_TAG"] if cand in df_aba.columns), None)
                                if col_maq:
                                    _, dias_no_mes = calendar.monthrange(int(ano_db_fuso), LISTA_MESES_PUROS.index(nome_mes_oficial) + 1)
                                    meses_atualizados.append(nome_mes_oficial)
                                    for _, r_up in df_aba.iterrows():
                                        m_val = str(r_up[col_maq]).strip().upper()
                                        fuso_padrao = obter_fuso_padrao(m_val, setor_db_fuso)
                                        for d in range(1, dias_no_mes + 1):
                                            qtd_val = 0
                                            if d in df_aba.columns: qtd_val = int(r_up[d]) if pd.notna(r_up[d]) else 0
                                            elif str(d) in df_aba.columns: qtd_val = int(r_up[str(d)]) if pd.notna(r_up[str(d)]) else 0
                                            novos_registros_ano.append({"Ano": int(ano_db_fuso), "Mes": nome_mes_oficial, "Dia": d, "Setor": setor_db_fuso, "Maquina_TAG": m_val, "Quantidade_Quebras": qtd_val, "Tipo_Fuso": fuso_padrao})

                        if novos_registros_ano:
                            df_limpo_fusos = df_fusos[~((df_fusos["Ano"] == int(ano_db_fuso)) & (df_fusos["Setor"] == setor_db_fuso) & (df_fusos["Mes"].isin(meses_atualizados)))]
                            df_atualizado_fusos = pd.concat([df_limpo_fusos, pd.DataFrame(novos_registros_ano)], ignore_index=True)
                            gerar_backup_seguro(ARQUIVO_FUSOS)
                            df_atualizado_fusos.to_excel(ARQUIVO_FUSOS, index=False)
                            invalidar_cache()
                            st.toast("✅ Base de fusos atualizada com sucesso!")
                            st.rerun()
                    except Exception as erro: st.error(f"Erro: {erro}")

    with tab_correias_db:
        st.markdown("### 🔄 Troca de Dados de Correias")
        col_d_cor, col_u_cor = st.columns([15, 25])
        with col_d_cor:
            st.markdown("#### 📥 Descarregar Planilha")
            filtro_export_setor = st.selectbox("Exportar Setor:", ["Todos os Setores"] + list(DICIONARIO_SETORES.keys()), key="sel_export_setor_cor")
            setores_alvo = list(DICIONARIO_SETORES.keys()) if filtro_export_setor == "Todos os Setores" else [filtro_export_setor]
            linhas_export = []
            for s_nome in setores_alvo:
                for maq in obter_maquinas_setor(s_nome, df_correias, df_fusos):
                    reg = df_correias[(df_correias["Setor"] == s_nome) & (df_correias["Maquina_TAG"] == maq)]
                    t1, d1, t2, d2 = "", "", "", ""
                    if not reg.empty:
                        ult = reg.iloc[-1]
                        t1, d1 = formatar_modelo(ult.get("Tipo_Correia_1", "")), str(ult.get("Data_Instalacao_1", "")).replace("nan", "").strip()
                        t2, d2 = formatar_modelo(ult.get("Tipo_Correia_2", "")), str(ult.get("Data_Instalacao_2", "")).replace("nan", "").strip()
                    linhas_export.append({"Setor": s_nome, "Maquina_TAG": maq, "Tipo_Correia_1": t1, "Data_Instalacao_1": d1, "Tipo_Correia_2": t2, "Data_Instalacao_2": d2})
            
            buf_down_cor = io.BytesIO()
            pd.DataFrame(linhas_export).to_excel(buf_down_cor, index=False, sheet_name="Correias")
            buf_down_cor.seek(0)
            st.download_button("📥 Baixar Correias (.xlsx)", data=buf_down_cor, file_name=f"correias_{date.today().strftime('%Y%m%d')}.xlsx", use_container_width=True)

        with col_u_cor:
            st.markdown("#### 📤 Enviar Dados Atualizados")
            up_arquivo_cor = st.file_uploader("Carregar nova planilha (.xlsx)", type=["xlsx"], key="uploader_novas_correias")
            modo_gravacao = st.radio("Modo de Atualização:", ["Mesclar e Atualizar", "Substituição Completa"], key="radio_modo_up_cor")
            if up_arquivo_cor is not None:
                if st.button("🚀 Atualizar Base de Correias", type="primary"):
                    try:
                        df_novo_cor = pd.read_excel(up_arquivo_cor)
                        mapa = {c: "Setor" if "setor" in c.lower() else "Maquina_TAG" if "maq" in c.lower() or "tag" in c.lower() else "Tipo_Correia_1" if "tipo" in c.lower() and "1" in c.lower() else "Data_Instalacao_1" if "data" in c.lower() and "1" in c.lower() else "Tipo_Correia_2" if "tipo" in c.lower() and "2" in c.lower() else "Data_Instalacao_2" if "data" in c.lower() and "2" in c.lower() else c for c in df_novo_cor.columns}
                        df_novo_cor.rename(columns=mapa, inplace=True)
                        for c in COLUNAS_CORREIAS: 
                            if c not in df_novo_cor.columns: df_novo_cor[c] = ""
                        
                        df_novo_cor["Maquina_TAG"] = df_novo_cor["Maquina_TAG"].astype(str).str.strip().str.upper()
                        gerar_backup_seguro(ARQUIVO_CORREIAS)
                        
                        if "Substituição Completa" in modo_gravacao: df_final_up_c = df_novo_cor[COLUNAS_CORREIAS]
                        else:
                            chaves = set(zip(df_novo_cor["Setor"], df_novo_cor["Maquina_TAG"]))
                            df_base_restante = df_correias[~df_correias.set_index(["Setor", "Maquina_TAG"]).index.isin(chaves)]
                            df_final_up_c = pd.concat([df_base_restante, df_novo_cor[COLUNAS_CORREIAS]], ignore_index=True)
                        
                        df_final_up_c.to_excel(ARQUIVO_CORREIAS, index=False)
                        invalidar_cache()
                        st.toast("✅ Base de Correias atualizada!")
                        st.rerun()
                    except Exception as e: st.error(f"Erro: {e}")

    with tab_paradas_db:
        st.markdown("### 🛠️ Gestão de Manutenções Corretivas")
        col_d_par, col_u_par = st.columns([15, 25])
        with col_d_par:
            st.markdown("#### 📥 Descarregar Registos")
            filtro_setor_par = st.selectbox("Filtrar Setor:", ["Todos os Setores"] + list(DICIONARIO_SETORES.keys()), key="sel_export_setor_parada")
            df_export_par = df_paradas if filtro_setor_par == "Todos os Setores" else df_paradas[df_paradas["Setor"] == filtro_setor_par]
            buf_down_p = io.BytesIO()
            df_export_par[COLUNAS_PARADAS].to_excel(buf_down_p, index=False, sheet_name="Corretivas")
            buf_down_p.seek(0)
            st.download_button("📥 Baixar Corretivas (.xlsx)", data=buf_down_p, file_name=f"corretivas_{date.today().strftime('%Y%m%d')}.xlsx", use_container_width=True)

        with col_u_par:
            st.markdown("#### 📤 Enviar Registos")
            up_arquivo_par = st.file_uploader("Carregar planilha (.xlsx)", type=["xlsx"], key="uploader_novas_paradas")
            if up_arquivo_par is not None:
                if st.button("🚀 Atualizar Corretivas", type="primary"):
                    try:
                        df_novo_p = pd.read_excel(up_arquivo_par)
                        mapa = {c: "Data" if "data" in c.lower() else "Setor" if "setor" in c.lower() else "Maquina_TAG" if "maq" in c.lower() else "Tipo_Manutencao" if "tipo" in c.lower() else "Descricao_Servico" if "desc" in c.lower() or "serv" in c.lower() else "Tempo_Parado_Horas" if "tempo" in c.lower() or "hora" in c.lower() else c for c in df_novo_p.columns}
                        df_novo_p.rename(columns=mapa, inplace=True)
                        for c in COLUNAS_PARADAS: 
                            if c not in df_novo_p.columns: df_novo_p[c] = 0.0 if c == "Tempo_Parado_Horas" else ""
                        
                        df_novo_p["Tempo_Parado_Horas"] = pd.to_numeric(df_novo_p["Tempo_Parado_Horas"], errors="coerce").fillna(0.0)
                        df_novo_p = df_novo_p[(df_novo_p["Tempo_Parado_Horas"] > 0) | (df_novo_p["Descricao_Servico"].astype(str).str.strip() != "")]
                        gerar_backup_seguro(ARQUIVO_PARADAS)
                        pd.concat([df_paradas, df_novo_p[COLUNAS_PARADAS]], ignore_index=True).to_excel(ARQUIVO_PARADAS, index=False)
                        invalidar_cache()
                        st.toast("✅ Corretivas registradas!")
                        st.rerun()
                    except Exception as e: st.error(f"Erro: {e}")

    with tab_pendencias_db:
        st.markdown("### 📋 Gestão de Pendências")
        col_d_pend, col_u_pend = st.columns([15, 25])
        with col_d_pend:
            st.markdown("#### 📥 Descarregar Pendências")
            buf_down_pend = io.BytesIO()
            df_pendencias[COLUNAS_PENDENCIAS].to_excel(buf_down_pend, index=False)
            buf_down_pend.seek(0)
            st.download_button("📥 Baixar Pendências (.xlsx)", data=buf_down_pend, file_name=f"pendencias_{date.today().strftime('%Y%m%d')}.xlsx", use_container_width=True)
        with col_u_pend:
            st.markdown("#### 📤 Enviar Backlog")
            up_arquivo_pend = st.file_uploader("Carregar planilha (.xlsx)", type=["xlsx"], key="up_pend")
            if up_arquivo_pend and st.button("🚀 Atualizar Pendências", type="primary"):
                try:
                    df_novo = pd.read_excel(up_arquivo_pend)
                    mapa = {c: "Setor" if "setor" in c.lower() else "Maquina_TAG" if "maq" in c.lower() else "Nome_Servico" if "nome" in c.lower() or "servi" in c.lower() else "Descricao_Pendencia" if "desc" in c.lower() else "Prioridade" if "prio" in c.lower() else "Status" if "stat" in c.lower() else c for c in df_novo.columns}
                    df_novo.rename(columns=mapa, inplace=True)
                    for c in COLUNAS_PENDENCIAS:
                        if c not in df_novo.columns: df_novo[c] = "Pendente" if c == "Status" else "Média" if c == "Prioridade" else ""
                    df_novo = df_novo[df_novo["Descricao_Pendencia"].astype(str).str.strip() != ""]
                    gerar_backup_seguro(ARQUIVO_PENDENCIAS)
                    pd.concat([df_pendencias, df_novo[COLUNAS_PENDENCIAS]], ignore_index=True).to_excel(ARQUIVO_PENDENCIAS, index=False)
                    invalidar_cache()
                    st.toast("✅ Pendências atualizadas!")
                    st.rerun()
                except Exception as e: st.error(f"Erro: {e}")

    with tab_maquinas_db:
        st.markdown("### 🏭 Cadastro e Gestão de Ativos")
        st.caption("Crie novas máquinas no sistema, remova máquinas antigas e parametrize os ativos operacionais.")

        c_add, c_del = st.columns(2)
        with c_add:
            with st.container(border=True):
                st.markdown("#### ➕ Inserir Nova Máquina")
                setor_add = st.selectbox("Setor de Destino:", list(DICIONARIO_SETORES.keys()), key="setor_add_maq")
                tag_add = st.text_input("Nova TAG:", placeholder="Ex: L-99").strip().upper()
                
                if st.button("Adicionar Máquina no Sistema", type="primary", use_container_width=True):
                    if tag_add:
                        if tag_add in obter_maquinas_setor(setor_add, df_correias, df_fusos):
                            st.warning(f"A máquina {tag_add} já está cadastrada no {setor_add}.")
                        else:
                            novo_registro_cor = {"Setor": setor_add, "Maquina_TAG": tag_add, "Tipo_Correia_1": "", "Data_Instalacao_1": "", "Tipo_Correia_2": "", "Data_Instalacao_2": ""}
                            df_correias = pd.concat([df_correias, pd.DataFrame([novo_registro_cor])], ignore_index=True)
                            gerar_backup_seguro(ARQUIVO_CORREIAS)
                            df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
                            
                            ano_corrente = date.today().year
                            novos_fusos_ano = [{"Ano": ano_corrente, "Mes": m_n, "Dia": 1, "Setor": setor_add, "Maquina_TAG": tag_add, "Quantidade_Quebras": 0, "Tipo_Fuso": obter_fuso_padrao(tag_add, setor_add)} for m_n in LISTA_MESES_PUROS]
                            df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_fusos_ano)], ignore_index=True)
                            gerar_backup_seguro(ARQUIVO_FUSOS)
                            df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
                            
                            invalidar_cache()
                            st.toast(f"✅ Máquina {tag_add} inserida com sucesso!")
                            st.rerun()

        with c_del:
            with st.container(border=True):
                st.markdown("#### 🗑️ Excluir Máquina")
                setor_del = st.selectbox("Setor de Origem:", list(DICIONARIO_SETORES.keys()), key="setor_del_maq")
                maqs_del_opcoes = obter_maquinas_setor(setor_del, df_correias, df_fusos)
                tag_del = st.selectbox("Máquina a Excluir:", ["-- Selecione --"] + maqs_del_opcoes, key="tag_del_maq")
                
                confirm_del = st.checkbox("Confirmo a exclusão definitiva desta máquina e todos os seus históricos", key="chk_del_maq")
                
                if st.button("Excluir Máquina e Dados", type="primary", use_container_width=True):
                    if tag_del != "-- Selecione --" and confirm_del:
                        
                        df_fusos = df_fusos[~((df_fusos["Setor"] == setor_del) & (df_fusos["Maquina_TAG"] == tag_del))]
                        df_correias = df_correias[~((df_correias["Setor"] == setor_del) & (df_correias["Maquina_TAG"] == tag_del))]
                        df_paradas = df_paradas[~((df_paradas["Setor"] == setor_del) & (df_paradas["Maquina_TAG"] == tag_del))]
                        df_pendencias = df_pendencias[~((df_pendencias["Setor"] == setor_del) & (df_pendencias["Maquina_TAG"] == tag_del))]
                        
                        gerar_backup_seguro(ARQUIVO_FUSOS)
                        gerar_backup_seguro(ARQUIVO_CORREIAS)
                        gerar_backup_seguro(ARQUIVO_PARADAS)
                        gerar_backup_seguro(ARQUIVO_PENDENCIAS)
                        
                        df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
                        df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
                        df_paradas.to_excel(ARQUIVO_PARADAS, index=False)
                        df_pendencias.to_excel(ARQUIVO_PENDENCIAS, index=False)
                        
                        invalidar_cache()
                        st.toast(f"✅ Máquina {tag_del} excluída de todo o sistema!")
                        st.rerun()
                    elif tag_del != "-- Selecione --" and not confirm_del:
                        st.warning("Marque a caixa de confirmação para poder excluir.")

        st.markdown("---")
        
        c_f_set, c_f_maq = st.columns([15, 20])
        with c_f_set:
            setores_disponiveis = list(DICIONARIO_SETORES.keys())
            setor_selecionado = st.selectbox("Configurar Setor Operacional:", setores_disponiveis, key="sel_setor_gestao_maq")

        maquinas_do_setor = obter_maquinas_setor(setor_selecionado, df_correias, df_fusos)

        with c_f_maq:
            maq_selecionada = st.selectbox("Máquina (TAG):", maquinas_do_setor, key="sel_maq_gestao_maq")

        fuso_atual = "FAG"
        sub_fuso = df_fusos[(df_fusos["Setor"] == setor_selecionado) & (df_fusos["Maquina_TAG"] == maq_selecionada)]
        if not sub_fuso.empty:
            val_fuso = str(sub_fuso.iloc[-1].get("Tipo_Fuso", "")).strip()
            if val_fuso in OPCOES_TIPO_FUSO: fuso_atual = val_fuso
            else: fuso_atual = obter_fuso_padrao(maq_selecionada, setor_selecionado)
        else:
            fuso_atual = obter_fuso_padrao(maq_selecionada, setor_selecionado)

        sub_cor = df_correias[(df_correias["Setor"] == setor_selecionado) & (df_correias["Maquina_TAG"] == maq_selecionada)]
        mod1_atual, dt1_atual, mod2_atual, dt2_atual = "", None, "", None

        if not sub_cor.empty:
            ult_c = sub_cor.iloc[-1]
            mod1_atual = formatar_modelo(ult_c.get("Tipo_Correia_1", ""))
            try: dt1_atual = pd.to_datetime(ult_c.get("Data_Instalacao_1", "")).date()
            except: pass
            mod2_atual = formatar_modelo(ult_c.get("Tipo_Correia_2", ""))
            try: dt2_atual = pd.to_datetime(ult_c.get("Data_Instalacao_2", "")).date()
            except: pass

        tem_duas_inicial = bool(mod2_atual or dt2_atual)

        with st.container(border=True):
            st.markdown(f"#### ⚙️ Parâmetros do Ativo: **{maq_selecionada if maq_selecionada else 'Nenhuma máquina selecionada'}**")
            
            if maq_selecionada:
                c_fuso, c_qtd_cor = st.columns([15, 20])
                
                with c_fuso:
                    idx_fuso = OPCOES_TIPO_FUSO.index(fuso_atual) if fuso_atual in OPCOES_TIPO_FUSO else 0
                    novo_fuso = st.selectbox("Tipo de Fuso:", OPCOES_TIPO_FUSO, index=idx_fuso, key=f"fuso_edit_{maq_selecionada}")

                with c_qtd_cor:
                    qtd_correias_opc = st.radio("Quantidade de Correias:", ["1 Correia (Única)", "2 Correias (Superior e Inferior)"], index=1 if tem_duas_inicial else 0, horizontal=True)

                st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

                if qtd_correias_opc == "1 Correia (Única)":
                    c_c1_mod, c_c1_dt = st.columns(2)
                    with c_c1_mod: novo_mod1 = st.text_input("Modelo da Correia:", value=mod1_atual)
                    with c_c1_dt: nova_dt1 = st.date_input("Data de Instalação:", value=dt1_atual, format="DD/MM/YYYY")
                    novo_mod2, nova_dt2 = "", None
                else:
                    c_sup1, c_sup2 = st.columns(2)
                    with c_sup1:
                        st.markdown("<p style='font-size:0.85rem; font-weight:800; color:#0f172a; margin-bottom:2px;'>🔼 Superior / Cabeceira</p>", unsafe_allow_html=True)
                        novo_mod1 = st.text_input("Modelo:", value=mod1_atual, key="m1")
                        nova_dt1 = st.date_input("Data de Instalação:", value=dt1_atual, format="DD/MM/YYYY", key="d1")
                    with c_sup2:
                        st.markdown("<p style='font-size:0.85rem; font-weight:800; color:#0f172a; margin-bottom:2px;'>🔽 Inferior / Traseira</p>", unsafe_allow_html=True)
                        novo_mod2 = st.text_input("Modelo:", value=mod2_atual, key="m2")
                        nova_dt2 = st.date_input("Data de Instalação:", value=dt2_atual, format="DD/MM/YYYY", key="d2")

                st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

                if st.button("💾 Salvar Parâmetros da Máquina", type="primary", use_container_width=True):
                    mask_fusos_maq = (df_fusos["Setor"] == setor_selecionado) & (df_fusos["Maquina_TAG"] == maq_selecionada)
                    if mask_fusos_maq.any(): df_fusos.loc[mask_fusos_maq, "Tipo_Fuso"] = novo_fuso
                    else:
                        novos_ano = [{"Ano": 2026, "Mes": m_n, "Dia": 1, "Setor": setor_selecionado, "Maquina_TAG": maq_selecionada, "Quantidade_Quebras": 0, "Tipo_Fuso": novo_fuso} for m_n in LISTA_MESES_PUROS]
                        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_ano)], ignore_index=True)
                    gerar_backup_seguro(ARQUIVO_FUSOS)
                    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

                    mask_cor_maq = (df_correias["Setor"] == setor_selecionado) & (df_correias["Maquina_TAG"] == maq_selecionada)
                    if mask_cor_maq.any():
                        idx_c = df_correias[mask_cor_maq].index[0]
                        df_correias.loc[idx_c, "Tipo_Correia_1"] = formatar_modelo(novo_mod1)
                        df_correias.loc[idx_c, "Data_Instalacao_1"] = str(nova_dt1) if nova_dt1 else ""
                        df_correias.loc[idx_c, "Tipo_Correia_2"] = formatar_modelo(novo_mod2)
                        df_correias.loc[idx_c, "Data_Instalacao_2"] = str(nova_dt2) if nova_dt2 else ""
                    else:
                        novo_registro_cor = {"Setor": setor_selecionado, "Maquina_TAG": maq_selecionada, "Tipo_Correia_1": formatar_modelo(novo_mod1), "Data_Instalacao_1": str(nova_dt1) if nova_dt1 else "", "Tipo_Correia_2": formatar_modelo(novo_mod2), "Data_Instalacao_2": str(nova_dt2) if nova_dt2 else ""}
                        df_correias = pd.concat([df_correias, pd.DataFrame([novo_registro_cor])], ignore_index=True)
                    
                    gerar_backup_seguro(ARQUIVO_CORREIAS)
                    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
                    invalidar_cache()
                    st.toast(f"✅ Parâmetros de {maq_selecionada} atualizados!")
                    st.rerun()

    with tab_backups_db:
        st.markdown("#### Cópias de Segurança")
        if os.path.exists("backups"):
            arquivos_bkp = sorted(os.listdir("backups"), reverse=True)
            if arquivos_bkp: st.dataframe([{"Backup": b, "Tamanho KB": round(os.path.getsize(os.path.join('backups', b))/1024, 1)} for b in arquivos_bkp], use_container_width=True)
            else: st.info("Nenhum backup.")
        else: st.info("Pasta não inicializada.")
