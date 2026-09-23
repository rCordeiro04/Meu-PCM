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

# Estilização CSS global para cartões, KPIs e mosaico
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 4.4rem !important;
            padding-bottom: 1rem !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }
        div[data-testid="column"] {
            padding: 1px !important;
            margin: 0px !important;
        }
        div[data-testid="stHorizontalBlock"] {
            gap: 4px !important;
            margin-bottom: 4px !important;
        }
        div[data-testid="stVegaLiteChart"] summary,
        div[data-testid="stVegaLiteChart"] .vega-actions {
            display: none !important;
        }
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
        .pill-legenda {
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
        }
        .dot-legenda {
            width: 9px;
            height: 9px;
            border-radius: 50%;
            display: inline-block;
        }
        .hud-detalhe {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            padding: 14px 18px;
            margin: 6px 0 12px 0;
            box-shadow: 0 6px 18px rgba(0,0,0,0.06);
            border-left: 6px solid #64748b;
        }
        .hud-detalhe.status-verde { border-left-color: #10b981; }
        .hud-detalhe.status-amarelo { border-left-color: #f59e0b; }
        .hud-detalhe.status-vermelho { border-left-color: #ef4444; }
        .hud-detalhe.status-cinza { border-left-color: #94a3b8; }
        .tag-pill {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            padding: 5px 12px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 700;
            color: #334155;
            display: inline-block;
            margin-right: 8px;
        }
        .badge-status {
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 800;
            text-transform: uppercase;
        }
        .badge-verde { background: #d1fae5; color: #065f46; }
        .badge-amarelo { background: #fef3c7; color: #92400e; }
        .badge-vermelho { background: #fee2e2; color: #991b1b; }
        .badge-cinza { background: #e2e8f0; color: #475569; }
        .chart-header-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-bottom: 6px;
            margin-bottom: 6px;
            border-bottom: 1px solid #f1f5f9;
        }
        .chart-header-title {
            font-size: 0.96rem;
            font-weight: 800;
            color: #0f172a;
        }
        .chart-header-badge {
            font-size: 0.82rem;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 6px;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

### Ficheiros de dados blindados e separados
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

# ==========================================
# LEITURA E BLINDAGEM DAS BASES
# ==========================================
df_fusos = None
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

# ==========================================
# DADOS HISTÓRICOS OFICIAIS (INJEÇÃO 2026)
# ==========================================
mapa_cargas_setor_b = [
    ("Janeiro", [("L-29", 1), ("L-30", 13), ("L-31", 10), ("L-32", 7), ("L-33", 5), ("L-34", 2), ("L-35", 2), ("L-36", 8), ("L-37", 2), ("L-38", 1), ("L-39", 1), ("L-40", 4), ("L-50", 7), ("L-51", 0), ("L-41", 1), ("L-42", 2), ("L-43", 0), ("L-44", 3), ("L-45", 0), ("L-46", 2), ("L-52", 3), ("L-53", 6)]),
    ("Fevereiro", [("L-29", 4), ("L-30", 23), ("L-31", 9), ("L-32", 34), ("L-33", 22), ("L-34", 19), ("L-35", 42), ("L-36", 22), ("L-37", 7), ("L-38", 4), ("L-39", 5), ("L-40", 11), ("L-50", 45), ("L-51", 2), ("L-41", 1), ("L-42", 2), ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 1), ("L-52", 2), ("L-53", 5)]),
    ("Março", [("L-29", 1), ("L-30", 15), ("L-31", 4), ("L-32", 8), ("L-33", 14), ("L-34", 6), ("L-35", 8), ("L-36", 9), ("L-37", 18), ("L-38", 2), ("L-39", 5), ("L-40", 6), ("L-50", 0), ("L-51", 4), ("L-41", 1), ("L-42", 1), ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 0), ("L-52", 3), ("L-53", 6)]),
    ("Abril", [("L-29", 2), ("L-30", 3), ("L-31", 1), ("L-32", 2), ("L-33", 17), ("L-34", 8), ("L-35", 7), ("L-36", 15), ("L-37", 17), ("L-38", 5), ("L-39", 12), ("L-40", 3), ("L-50", 2), ("L-51", 3), ("L-41", 3), ("L-42", 4), ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 1), ("L-52", 1), ("L-53", 4)]),
    ("Maio", [("L-29", 0), ("L-30", 5), ("L-31", 8), ("L-32", 0), ("L-33", 6), ("L-34", 2), ("L-35", 7), ("L-36", 10), ("L-37", 12), ("L-38", 2), ("L-39", 5), ("L-40", 3), ("L-50", 1), ("L-51", 0), ("L-41", 2), ("L-42", 1), ("L-43", 8), ("L-44", 1), ("L-45", 0), ("L-46", 2), ("L-52", 5), ("L-53", 5)]),
    ("Junho", [("L-29", 6), ("L-30", 16), ("L-31", 35), ("L-32", 4), ("L-33", 10), ("L-34", 4), ("L-35", 8), ("L-36", 11), ("L-37", 6), ("L-38", 1), ("L-39", 9), ("L-40", 11), ("L-50", 1), ("L-51", 1), ("L-41", 6), ("L-42", 1), ("L-43", 2), ("L-44", 0), ("L-45", 0), ("L-46", 0), ("L-52", 1), ("L-53", 2)]),
    ("Julho", [("L-29", 0), ("L-30", 31), ("L-31", 4), ("L-32", 13), ("L-33", 26), ("L-34", 7), ("L-35", 14), ("L-36", 2), ("L-37", 14), ("L-38", 0), ("L-39", 2), ("L-40", 11), ("L-50", 2), ("L-51", 6), ("L-41", 3), ("L-42", 2), ("L-43", 3), ("L-44", 5), ("L-45", 0), ("L-46", 0), ("L-52", 6), ("L-53", 4)]),
]

mapa_cargas_setor_a = [
    ("Janeiro", [("L-01", 8), ("L-02", 9), ("L-03", 2), ("L-04", 3), ("L-05", 3), ("L-06", 13), ("L-07", 6), ("L-08", 4), ("L-09", 8), ("L-10", 9), ("L-11", 3), ("L-12", 6), ("L-13", 5), ("L-14", 8), ("L-15", 9), ("L-16", 4), ("L-17", 2), ("L-18", 3), ("L-19", 10), ("L-20", 5), ("L-21", 4), ("L-22", 4), ("L-23", 6), ("L-24", 10), ("L-25", 4), ("L-26", 3), ("L-27", 4), ("L-28", 3)]),
    ("Fevereiro", [("L-01", 3), ("L-02", 6), ("L-03", 4), ("L-04", 4), ("L-05", 2), ("L-06", 5), ("L-07", 7), ("L-08", 3), ("L-09", 4), ("L-10", 2), ("L-11", 4), ("L-12", 14), ("L-13", 14), ("L-14", 13), ("L-15", 4), ("L-16", 10), ("L-17", 1), ("L-18", 4), ("L-19", 0), ("L-20", 6), ("L-21", 7), ("L-22", 4), ("L-23", 3), ("L-24", 4), ("L-25", 4), ("L-26", 0), ("L-27", 0), ("L-28", 8)]),
    ("Março", [("L-01", 4), ("L-02", 7), ("L-03", 6), ("L-04", 3), ("L-05", 1), ("L-06", 5), ("L-07", 10), ("L-08", 2), ("L-09", 12), ("L-10", 8), ("L-11", 10), ("L-12", 9), ("L-13", 9), ("L-14", 9), ("L-15", 12), ("L-16", 3), ("L-17", 8), ("L-18", 20), ("L-19", 10), ("L-20", 3), ("L-21", 4), ("L-22", 5), ("L-23", 3), ("L-24", 8), ("L-25", 4), ("L-26", 9), ("L-27", 2), ("L-28", 2)]),
    ("Abril", [("L-01", 1), ("L-02", 8), ("L-03", 3), ("L-04", 6), ("L-05", 2), ("L-06", 7), ("L-07", 8), ("L-08", 0), ("L-09", 2), ("L-10", 6), ("L-11", 9), ("L-12", 6), ("L-13", 5), ("L-14", 2), ("L-15", 2), ("L-16", 7), ("L-17", 5), ("L-18", 8), ("L-19", 2), ("L-20", 5), ("L-21", 3), ("L-22", 12), ("L-23", 5), ("L-24", 9), ("L-25", 2), ("L-26", 7), ("L-27", 5), ("L-28", 2)]),
    ("Maio", [("L-01", 8), ("L-02", 5), ("L-03", 11), ("L-04", 1), ("L-05", 5), ("L-06", 7), ("L-07", 6), ("L-08", 2), ("L-09", 8), ("L-10", 1), ("L-11", 4), ("L-12", 1), ("L-13", 2), ("L-14", 6), ("L-15", 6), ("L-16", 10), ("L-17", 4), ("L-18", 5), ("L-19", 7), ("L-20", 13), ("L-21", 3), ("L-22", 4), ("L-23", 9), ("L-24", 7), ("L-25", 1), ("L-26", 2), ("L-27", 4), ("L-28", 4)]),
    ("Junho", [("L-01", 5), ("L-02", 2), ("L-03", 4), ("L-04", 4), ("L-05", 5), ("L-06", 3), ("L-07", 8), ("L-08", 1), ("L-09", 6), ("L-10", 5), ("L-11", 9), ("L-12", 5), ("L-13", 8), ("L-14", 3), ("L-15", 2), ("L-16", 5), ("L-17", 2), ("L-18", 4), ("L-19", 3), ("L-20", 5), ("L-21", 5), ("L-22", 3), ("L-23", 4), ("L-24", 4), ("L-25", 3), ("L-26", 9), ("L-27", 3), ("L-28", 1)]),
    ("Julho", [("L-01", 5), ("L-02", 3), ("L-03", 3), ("L-04", 10), ("L-05", 2), ("L-06", 6), ("L-07", 11), ("L-08", 8), ("L-09", 11), ("L-10", 7), ("L-11", 5), ("L-12", 3), ("L-13", 2), ("L-14", 7), ("L-15", 6), ("L-16", 9), ("L-17", 4), ("L-18", 8), ("L-19", 8), ("L-20", 5), ("L-21", 11), ("L-22", 18), ("L-23", 9), ("L-24", 10), ("L-25", 7), ("L-26", 9), ("L-27", 3), ("L-28", 9)]),
    ("Agosto", [("L-01", 5), ("L-02", 8), ("L-03", 3), ("L-04", 5), ("L-05", 11), ("L-06", 8), ("L-07", 15), ("L-08", 5), ("L-09", 14), ("L-10", 6), ("L-11", 9), ("L-12", 5), ("L-13", 7), ("L-14", 9), ("L-15", 11), ("L-16", 5), ("L-17", 6), ("L-18", 12), ("L-19", 10), ("L-20", 10), ("L-21", 9), ("L-22", 16), ("L-23", 11), ("L-24", 8), ("L-25", 4), ("L-26", 2), ("L-27", 5), ("L-28", 10)]),
    ("Setembro", [("L-01", 1), ("L-02", 2), ("L-03", 2), ("L-04", 4), ("L-05", 1), ("L-06", 0), ("L-07", 1), ("L-08", 0), ("L-09", 3), ("L-10", 4), ("L-11", 7), ("L-12", 2), ("L-13", 4), ("L-14", 3), ("L-15", 1), ("L-16", 0), ("L-17", 1), ("L-18", 1), ("L-19", 4), ("L-20", 4), ("L-21", 2), ("L-22", 3), ("L-23", 1), ("L-24", 5), ("L-25", 1), ("L-26", 0), ("L-27", 2), ("L-28", 4)]),
]

mapa_cargas_setor_latex = [
    ("Janeiro", [("B-71", 7), ("B-72", 10), ("B-73", 13), ("B-74", 18), ("B-75", 4), ("B-76", 9), ("B-77", 10), ("B-78", 5), ("B-79", 1), ("B-80", 0), ("B-83", 6), ("B-84", 7), ("B-85", 9), ("B-86", 5), ("B-87", 5), ("B-88", 10), ("B-89", 1), ("B-102", 0), ("B-103", 0), ("B-104", 0)]),
    ("Fevereiro", [("B-71", 1), ("B-72", 2), ("B-73", 5), ("B-74", 0), ("B-75", 0), ("B-76", 1), ("B-77", 2), ("B-78", 3), ("B-79", 4), ("B-80", 0), ("B-83", 1), ("B-84", 2), ("B-85", 4), ("B-86", 4), ("B-87", 2), ("B-88", 0), ("B-89", 0), ("B-102", 0), ("B-103", 0), ("B-104", 0)]),
    ("Março", [("B-71", 5), ("B-72", 7), ("B-73", 2), ("B-74", 1), ("B-75", 0), ("B-76", 3), ("B-77", 4), ("B-78", 4), ("B-79", 0), ("B-80", 1), ("B-83", 9), ("B-84", 18), ("B-85", 9), ("B-86", 9), ("B-87", 14), ("B-88", 0), ("B-89", 8), ("B-102", 0), ("B-103", 0), ("B-104", 0)]),
    ("Abril", [("B-71", 5), ("B-72", 9), ("B-73", 3), ("B-74", 5), ("B-75", 2), ("B-76", 4), ("B-77", 5), ("B-78", 3), ("B-79", 0), ("B-80", 0), ("B-83", 10), ("B-84", 18), ("B-85", 8), ("B-86", 12), ("B-87", 9), ("B-88", 0), ("B-89", 10), ("B-102", 0), ("B-103", 0), ("B-104", 0)]),
    ("Maio", [("B-71", 1), ("B-72", 4), ("B-73", 7), ("B-74", 5), ("B-75", 2), ("B-76", 4), ("B-77", 1), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 15), ("B-84", 18), ("B-85", 6), ("B-86", 12), ("B-87", 7), ("B-88", 0), ("B-89", 17), ("B-102", 0), ("B-103", 0), ("B-104", 1)]),
    ("Junho", [("B-71", 1), ("B-72", 1), ("B-73", 7), ("B-74", 0), ("B-75", 7), ("B-76", 1), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 9), ("B-84", 7), ("B-85", 10), ("B-86", 13), ("B-87", 5), ("B-88", 1), ("B-89", 8), ("B-102", 0), ("B-103", 0), ("B-104", 0)]),
    ("Julho", [("B-71", 3), ("B-72", 1), ("B-73", 3), ("B-74", 0), ("B-75", 0), ("B-76", 0), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 8), ("B-84", 10), ("B-85", 7), ("B-86", 9), ("B-87", 10), ("B-88", 0), ("B-89", 7), ("B-102", 0), ("B-103", 0), ("B-104", 0)]),
    ("Agosto", [("B-71", 5), ("B-72", 1), ("B-73", 0), ("B-74", 0), ("B-75", 0), ("B-76", 0), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 1), ("B-84", 3), ("B-85", 6), ("B-86", 1), ("B-87", 1), ("B-88", 0), ("B-89", 1), ("B-102", 0), ("B-103", 0), ("B-104", 0)]),
]

mapa_cargas_setor_menegatto = [
    ("Janeiro", [("B-47", 4), ("B-48", 3), ("B-49", 5), ("B-81", 3), ("B-82", 3), ("B-93", 10), ("B-94", 9), ("B-95", 4), ("B-96", 6), ("B-97", 4), ("B-98", 1), ("B-99", 6), ("B-100", 6), ("B-101", 5), ("B-107", 0), ("B-108", 0)]),
    ("Fevereiro", [("B-47", 4), ("B-48", 3), ("B-49", 2), ("B-81", 2), ("B-82", 0), ("B-93", 1), ("B-94", 0), ("B-95", 3), ("B-96", 4), ("B-97", 6), ("B-98", 1), ("B-99", 6), ("B-100", 16), ("B-101", 1), ("B-107", 0), ("B-108", 0)]),
    ("Março", [("B-47", 0), ("B-48", 5), ("B-49", 2), ("B-81", 3), ("B-82", 3), ("B-93", 3), ("B-94", 7), ("B-95", 12), ("B-96", 2), ("B-97", 5), ("B-98", 1), ("B-99", 9), ("B-100", 0), ("B-101", 3), ("B-107", 0), ("B-108", 0)]),
    ("Abril", [("B-47", 9), ("B-48", 8), ("B-49", 3), ("B-81", 8), ("B-82", 5), ("B-93", 3), ("B-94", 2), ("B-95", 25), ("B-96", 5), ("B-97", 3), ("B-98", 2), ("B-99", 14), ("B-100", 2), ("B-101", 4), ("B-107", 0), ("B-108", 0)]),
    ("Maio", [("B-47", 5), ("B-48", 9), ("B-49", 1), ("B-81", 7), ("B-82", 1), ("B-93", 1), ("B-94", 0), ("B-95", 4), ("B-96", 0), ("B-97", 9), ("B-98", 0), ("B-99", 1), ("B-100", 4), ("B-101", 3), ("B-107", 0), ("B-108", 0)]),
    ("Junho", [("B-47", 4), ("B-48", 6), ("B-49", 1), ("B-81", 1), ("B-82", 3), ("B-93", 5), ("B-94", 7), ("B-95", 15), ("B-96", 5), ("B-97", 9), ("B-98", 0), ("B-99", 6), ("B-100", 3), ("B-101", 1), ("B-107", 0), ("B-108", 0)]),
    ("Julho", [("B-47", 11), ("B-48", 5), ("B-49", 7), ("B-81", 6), ("B-82", 5), ("B-93", 1), ("B-94", 0), ("B-95", 1), ("B-96", 3), ("B-97", 1), ("B-98", 1), ("B-99", 3), ("B-100", 2), ("B-101", 5), ("B-107", 1), ("B-108", 0)]),
    ("Agosto", [("B-47", 7), ("B-48", 1), ("B-49", 4), ("B-81", 0), ("B-82", 3), ("B-93", 2), ("B-94", 0), ("B-95", 1), ("B-96", 2), ("B-97", 3), ("B-98", 0), ("B-99", 0), ("B-100", 1), ("B-101", 2), ("B-107", 0), ("B-108", 0)]),
]

precisa_salvar_fusos = False

for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_b:
    linhas_mes = df_fusos[(df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor B")]
    if linhas_mes.empty or linhas_mes["Quantidade_Quebras"].sum() == 0:
        df_fusos = df_fusos[~((df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor B"))]
        novos_reg = [{"Ano": 2026, "Mes": nome_mes_carga, "Setor": "Setor B", "Maquina_TAG": maq, "Quantidade_Quebras": int(qtd), "Tipo_Fuso": "TEP"} for maq, qtd in lista_dados_carga]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_a:
    linhas_mes = df_fusos[(df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor A")]
    if linhas_mes.empty or linhas_mes["Quantidade_Quebras"].sum() == 0:
        df_fusos = df_fusos[~((df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor A"))]
        novos_reg = [{"Ano": 2026, "Mes": nome_mes_carga, "Setor": "Setor A", "Maquina_TAG": maq, "Quantidade_Quebras": int(qtd), "Tipo_Fuso": "FAG"} for maq, qtd in lista_dados_carga]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_latex:
    linhas_mes = df_fusos[(df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor Látex")]
    if linhas_mes.empty or len(linhas_mes) < 20:
        df_fusos = df_fusos[~((df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor Látex"))]
        novos_reg = [{"Ano": 2026, "Mes": nome_mes_carga, "Setor": "Setor Látex", "Maquina_TAG": maq, "Quantidade_Quebras": int(qtd), "Tipo_Fuso": "M4BA"} for maq, qtd in lista_dados_carga]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_menegatto:
    linhas_mes = df_fusos[(df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor Menegatto")]
    if linhas_mes.empty or linhas_mes["Quantidade_Quebras"].sum() == 0:
        df_fusos = df_fusos[~((df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes_carga) & (df_fusos["Setor"] == "Setor Menegatto"))]
        novos_reg = [{"Ano": 2026, "Mes": nome_mes_carga, "Setor": "Setor Menegatto", "Maquina_TAG": maq, "Quantidade_Quebras": int(qtd), "Tipo_Fuso": "MENEGATTO"} for maq, qtd in lista_dados_carga]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

if precisa_salvar_fusos:
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Leitura blindada de Correias
df_correias = None
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

df_correias["Tipo_Correia_1"] = df_correias["Tipo_Correia_1"].apply(formatar_modelo)
df_correias["Data_Instalacao_1"] = df_correias["Data_Instalacao_1"].astype(str)
df_correias["Tipo_Correia_2"] = df_correias["Tipo_Correia_2"].apply(formatar_modelo)
df_correias["Data_Instalacao_2"] = df_correias["Data_Instalacao_2"].astype(str)

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

def obter_maquinas_setor(setor_nome, df_ref_correias=None, df_ref_fusos=None):
    lista_base = list(DICIONARIO_SETORES.get(setor_nome, []))
    extras = set()
    if df_ref_correias is not None and not df_ref_correias.empty:
        extras.update(df_ref_correias[df_ref_correias["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique())
    if df_ref_fusos is not None and not df_ref_fusos.empty:
        extras.update(df_ref_fusos[df_ref_fusos["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique())
    return sorted(list(set(lista_base).union(extras)))

mapa_setor_maquina = {}
for setor_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m:
        mapa_setor_maquina[m] = setor_nome
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
if "card_selecionado_kpi" not in st.session_state:
    st.session_state.card_selecionado_kpi = None
if "aba_setor_fuso" not in st.session_state:
    st.session_state.aba_setor_fuso = "Geral"

def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina

# ==========================================
# CÁLCULOS DO MOSAICO DE CORREIAS
# ==========================================
todas_as_maquinas = []
for setor_nome in DICIONARIO_SETORES.keys():
    for m in obter_maquinas_setor(setor_nome, df_correias, df_fusos):
        if m not in todas_as_maquinas:
            todas_as_maquinas.append(m)

data_hoje = date.today()
dados_maquinas = {}
css_botoes = []
lista_correias_todas = []
lista_correias_novas = []
lista_correias_meia = []
lista_correias_criticas = []

def avaliar_correia(dt_val):
    if not dt_val or dt_val == "nan" or str(dt_val).strip() == "":
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
        return None, str(dt_val), "Data inválida"

for maq_tag in todas_as_maquinas:
    setor_m = mapa_setor_maquina.get(maq_tag, "Setor A")
    reg_maq = df_correias[(df_correias["Setor"] == setor_m) & (df_correias["Maquina_TAG"] == maq_tag)]
    t1, d1_str, t1_uso, c1_score = "Não informada", "Sem registro", "Sem histórico", None
    t2, d2_str, t2_uso, c2_score = "Não informada", "Sem registro", "Sem histórico", None
    tem_c1, tem_c2 = False, False

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
            if c1_score == 1: lista_correias_novas.append(reg_c1)
            elif c1_score == 2: lista_correias_meia.append(reg_c1)
            elif c1_score == 3: lista_correias_criticas.append(reg_c1)

        v2 = formatar_modelo(ultimo.get("Tipo_Correia_2", ""))
        dt2_raw = str(ultimo.get("Data_Instalacao_2", "")).strip()
        if (v2 and v2 != "nan") or (dt2_raw and dt2_raw != "nan"):
            t2 = v2 if (v2 and v2 != "nan") else "Não informada"
            c2_score, d2_str, t2_uso = avaliar_correia(dt2_raw)
            tem_c2 = True
            reg_c2 = {"tag": maq_tag, "pos": "Inferior", "modelo": t2, "data": d2_str, "uso": t2_uso, "setor": setor_m}
            lista_correias_todas.append(reg_c2)
            if c2_score == 1: lista_correias_novas.append(reg_c2)
            elif c2_score == 2: lista_correias_meia.append(reg_c2)
            elif c2_score == 3: lista_correias_criticas.append(reg_c2)

    scores = [s for s in [c1_score, c2_score] if s is not None]
    if scores:
        pior = max(scores)
        if pior == 1:
            classe_card, cor_grad, cor_borda, status_label = "status-verde", "linear-gradient(135deg, #10b981 0%, #059669 100%)", "#047857", "Nova"
        elif pior == 2:
            classe_card, cor_grad, cor_borda, status_label = "status-amarelo", "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)", "#b45309", "Meia-Vida"
        else:
            classe_card, cor_grad, cor_borda, status_label = "status-vermelho", "linear-gradient(135deg, #ef4444 0%, #dc2626 100%)", "#b91c1c", "Troca Necessária"
    else:
        classe_card, cor_grad, cor_borda, status_label = "status-cinza", "linear-gradient(135deg, #64748b 0%, #475569 100%)", "#334155", "Sem Dados"

    dados_maquinas[maq_tag] = {
        "setor": setor_m, "t1": t1, "d1": d1_str, "uso1": t1_uso, "tem_c1": tem_c1,
        "t2": t2, "d2": d2_str, "uso2": t2_uso, "tem_c2": tem_c2, "status_label": status_label, "classe_card": classe_card
    }
    chave_btn = f"btn_q_{maq_tag.replace('-', '_')}"
    css_botoes.append(
        f"""
        button[key="{chave_btn}"], div.st-key-{chave_btn} button {{
            background: {cor_grad} !important; color: #ffffff !important; border: 1px solid {cor_borda} !important;
        }}
        """
    )

st.markdown(f"<style>{''.join(css_botoes)}</style>", unsafe_allow_html=True)

# ==========================================
# BARRA LATERAL (SIDEBAR COM BOTÕES LADO A LADO E KEYS ÚNICAS)
# ==========================================
with st.sidebar:
    st.title("⚙️ Portal PCM")
    st.caption("Planejamento e Controle de Manutenção")
    st.markdown("---")
    
    st.markdown("<div style='font-size:0.85rem; font-weight:700; color:#64748b; margin-bottom:6px;'>📊 PAINÉIS</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        tipo_pf = "primary" if st.session_state.pagina_atual == "Painel Fusos" else "secondary"
        if st.button("🔩 Fusos", key="btn_nav_painel_fusos_menu", use_container_width=True, type=tipo_pf):
            navegar("Painel Fusos")
            st.rerun()
    with c2:
        tipo_pc = "primary" if st.session_state.pagina_atual == "Painel Correias" else "secondary"
        if st.button("🔄 Correias", key="btn_nav_painel_correias_menu", use_container_width=True, type=tipo_pc):
            navegar("Painel Correias")
            st.rerun()

    st.markdown("<div style='font-size:0.85rem; font-weight:700; color:#64748b; margin-top:10px; margin-bottom:6px;'>📝 LANÇAMENTOS</div>", unsafe_allow_html=True)
    c3, c4 = st.columns(2)
    with c3:
        tipo_lf = "primary" if st.session_state.pagina_atual == "Lançamento Fusos" else "secondary"
        if st.button("🔩 Fusos", key="btn_nav_lancto_fusos_menu", use_container_width=True, type=tipo_lf):
            navegar("Lançamento Fusos")
            st.rerun()
    with c4:
        tipo_lc = "primary" if st.session_state.pagina_atual == "Correias" else "secondary"
        if st.button("🔄 Correias", key="btn_nav_lancto_correias_menu", use_container_width=True, type=tipo_lc):
            navegar("Correias")
            st.rerun()

    c5, c6 = st.columns(2)
    with c5:
        tipo_prev = "primary" if st.session_state.pagina_atual == "Preventiva" else "secondary"
        if st.button("🛠️ Preventiva", key="btn_nav_lancto_prev_menu", use_container_width=True, type=tipo_prev):
            navegar("Preventiva")
            st.rerun()
    with c6:
        tipo_maq = "primary" if st.session_state.pagina_atual == "Máquinas" else "secondary"
        if st.button("🏭 Máquinas", key="btn_nav_lancto_maq_menu", use_container_width=True, type=tipo_maq):
            navegar("Máquinas")
            st.rerun()

    st.markdown("---")
    st.caption("PCM • Versão Gerencial")

tela = st.session_state.pagina_atual
lista_meses_puros = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]
MAPA_MES_ABREV = {
    "Janeiro": "JAN", "Fevereiro": "FEV", "Março": "MAR", "Abril": "ABR",
    "Maio": "MAI", "Junho": "JUN", "Julho": "JUL", "Agosto": "AGO",
    "Setembro": "SET", "Outubro": "OUT", "Novembro": "NOV", "Dezembro": "DEZ"
}
ORDEM_MESES_ABREV = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
colunas_dias = [f"{i:02d}" for i in range(1, 32)]

# =========================================================================
# 1. PAINEL GERENCIAL DE CORREIAS (DASHBOARD)
# =========================================================================
if tela == "Painel Correias":
    col_t, col_f1, col_f2, col_leg = st.columns([3.2, 1.4, 1.4, 5.0])
    with col_t:
        st.markdown("<div style='font-size: 2rem; font-weight: 900; color: #0f172a;'>Dashboard Correias</div>", unsafe_allow_html=True)
    with col_f1:
        filtro_setor = st.selectbox("Setor", ["Todos os Setores"] + list(DICIONARIO_SETORES.keys()), key="f_set_cor_painel", label_visibility="collapsed")
    with col_f2:
        modelos_unicos = {r["modelo"] for r in lista_correias_todas if r["modelo"] != "Não informada"}
        filtro_modelo = st.selectbox("Tipo", ["Todos os Tipos"] + sorted(list(modelos_unicos)), key="f_mod_cor_painel", label_visibility="collapsed")
    with col_leg:
        st.markdown(
            """
            <div style='display: flex; justify-content: flex-end; align-items: center; gap: 8px;'>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#10b981;'></span> Nova</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#f59e0b;'></span> Meia-Vida</span>
                <span class='pill-legenda'><span class='dot-legenda' style='background:#ef4444;'></span> Troca Necessária</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Total Instaladas</div><div class='kpi-val'>{len(lista_correias_todas)}</div></div><div>📦</div></div>", unsafe_allow_html=True)
    with k2:
        st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Novas</div><div class='kpi-val' style='color:#059669;'>{len(lista_correias_novas)}</div></div><div>🟢</div></div>", unsafe_allow_html=True)
    with k3:
        st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Meia Vida</div><div class='kpi-val' style='color:#d97706;'>{len(lista_correias_meia)}</div></div><div>🟡</div></div>", unsafe_allow_html=True)
    with k4:
        st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Críticas</div><div class='kpi-val' style='color:#dc2626;'>{len(lista_correias_criticas)}</div></div><div>🔴</div></div>", unsafe_allow_html=True)

    if st.session_state.maq_clicada_cor is not None:
        maq_sel = st.session_state.maq_clicada_cor
        c_box, c_close = st.columns([6.2, 0.8])
        with c_box:
            st.markdown(
                f"""
                <div class="hud-detalhe {maq_sel['classe_card']}">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                        <span class="tag-pill" style="background:#0f172a; color:#fff;">⚙️ {maq_sel['tag']} ({maq_sel['setor']})</span>
                        <span class="badge-status">{maq_sel['status_label']}</span>
                    </div>
                    <div>
                        <span class="tag-pill">🔼 <b>Superior:</b> {maq_sel['t1']} | 📅 {maq_sel['d1']} | ⏱️ {maq_sel['uso1']}</span>
                        <span class="tag-pill">🔽 <b>Inferior:</b> {maq_sel['t2']} | 📅 {maq_sel['d2']} | ⏱️ {maq_sel['uso2']}</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with c_close:
            if st.button("✖ Fechar", key="btn_fechar_detalhe_cor"):
                st.session_state.maq_clicada_cor = None
                st.rerun()

    COLS_GRELHA = 12
    linhas_grid = [todas_as_maquinas[i:i + COLS_GRELHA] for i in range(0, len(todas_as_maquinas), COLS_GRELHA)]
    for linha in linhas_grid:
        cols = st.columns(COLS_GRELHA)
        for idx_col, maq_tag in enumerate(linha):
            with cols[idx_col]:
                if st.button(maq_tag, key=f"btn_q_{maq_tag.replace('-', '_')}", use_container_width=True):
                    st.session_state.maq_clicada_cor = {"tag": maq_tag, **dados_maquinas[maq_tag]}
                    st.rerun()

# =========================================================================
# 2. LANÇAMENTO DE CORREIAS (SUPERIOR / INFERIOR)
# =========================================================================
elif tela == "Correias":
    st.title("🔄 Lançamento: Gestão de Correias")
    with st.container(border=True):
        col_setor, col_add, col_del = st.columns([1.6, 2.2, 2.2])
        with col_setor:
            setor_selecionado = st.selectbox("🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_set_cor_lancto")
        df_atual_cor = pd.read_excel(ARQUIVO_CORREIAS)
        maquinas_do_setor = obter_maquinas_setor(setor_selecionado, df_atual_cor, df_fusos)

        with col_add:
            nova_maq_cor = st.text_input("Nova Máquina", placeholder="Ex: L-99 ou B-105", key=f"inp_add_cor_{setor_selecionado}").strip().upper()
            if st.button("➕ Adicionar", key=f"btn_add_cor_{setor_selecionado}"):
                if nova_maq_cor and nova_maq_cor not in maquinas_do_setor:
                    df_atual_cor = pd.concat([df_atual_cor, pd.DataFrame([{"Setor": setor_selecionado, "Maquina_TAG": nova_maq_cor, "Tipo_Correia_1": "", "Data_Instalacao_1": "", "Tipo_Correia_2": "", "Data_Instalacao_2": ""}])], ignore_index=True)
                    df_atual_cor.to_excel(ARQUIVO_CORREIAS, index=False)
                    st.success(f"Máquina {nova_maq_cor} adicionada!")
                    st.rerun()

        with col_del:
            maq_del_cor = st.selectbox("Excluir", ["-- Selecione --"] + maquinas_do_setor, key=f"sel_del_cor_{setor_selecionado}")
            if st.button("🗑️ Excluir", key=f"btn_del_cor_{setor_selecionado}"):
                if maq_del_cor != "-- Selecione --":
                    df_atual_cor = df_atual_cor[~((df_atual_cor["Setor"] == setor_selecionado) & (df_atual_cor["Maquina_TAG"] == maq_del_cor))]
                    df_atual_cor.to_excel(ARQUIVO_CORREIAS, index=False)
                    st.success(f"Máquina {maq_del_cor} removida!")
                    st.rerun()

    dados_grade_cor = []
    df_filtrado_cor = df_atual_cor[df_atual_cor["Setor"] == setor_selecionado]
    for maq in maquinas_do_setor:
        reg = df_filtrado_cor[df_filtrado_cor["Maquina_TAG"] == maq]
        t1, dt1, t2, dt2 = "", None, "", None
        if not reg.empty:
            ult = reg.iloc[-1]
            t1 = formatar_modelo(ult.get("Tipo_Correia_1", ""))
            t2 = formatar_modelo(ult.get("Tipo_Correia_2", ""))
            try: dt1 = pd.to_datetime(ult.get("Data_Instalacao_1")).date() if pd.notna(ult.get("Data_Instalacao_1")) else None
            except Exception: dt1 = None
            try: dt2 = pd.to_datetime(ult.get("Data_Instalacao_2")).date() if pd.notna(ult.get("Data_Instalacao_2")) else None
            except Exception: dt2 = None
        dados_grade_cor.append({"Máquina": maq, "Modelo (Superior / Cabeceira)": t1, "Data (Superior / Cabeceira)": dt1, "Modelo (Inferior / Traseira)": t2, "Data (Inferior / Traseira)": dt2})

    tabela_editada_cor = st.data_editor(
        pd.DataFrame(dados_grade_cor),
        column_config={
            "Máquina": st.column_config.TextColumn("Máquina", disabled=True),
            "Data (Superior / Cabeceira)": st.column_config.DateColumn("Data (Superior)", format="DD/MM/YYYY"),
            "Data (Inferior / Traseira)": st.column_config.DateColumn("Data (Inferior)", format="DD/MM/YYYY"),
        },
        hide_index=True,
        use_container_width=True,
        height=440,
        key=f"editor_grade_cor_{setor_selecionado}",
    )

    if st.button("💾 Salvar Correias", key=f"btn_salvar_cor_{setor_selecionado}", type="primary"):
        df_limpo_cor = df_atual_cor[df_atual_cor["Setor"] != setor_selecionado]
        novos_regs = []
        for _, l in tabela_editada_cor.iterrows():
            d1_val = str(l["Data (Superior / Cabeceira)"]) if pd.notna(l["Data (Superior / Cabeceira)"]) else ""
            d2_val = str(l["Data (Inferior / Traseira)"]) if pd.notna(l["Data (Inferior / Traseira)"]) else ""
            novos_regs.append({"Setor": setor_selecionado, "Maquina_TAG": l["Máquina"], "Tipo_Correia_1": formatar_modelo(l["Modelo (Superior / Cabeceira)"]), "Data_Instalacao_1": d1_val, "Tipo_Correia_2": formatar_modelo(l["Modelo (Inferior / Traseira)"]), "Data_Instalacao_2": d2_val})
        pd.concat([df_limpo_cor, pd.DataFrame(novos_regs)], ignore_index=True).to_excel(ARQUIVO_CORREIAS, index=False)
        st.success("✅ Apontamentos de correias salvos com sucesso!")
        st.rerun()

# =========================================================================
# 3. PAINEL GERENCIAL DE FUSOS (DASHBOARD COMPLETO)
# =========================================================================
elif tela == "Painel Fusos":
    col_tf, col_ano_f = st.columns([3.8, 1.4])
    with col_tf:
        st.markdown("<div style='font-size: 2rem; font-weight: 900; color: #0f172a;'>Dashboard Fusos</div>", unsafe_allow_html=True)
    with col_ano_f:
        ano_painel = st.selectbox("Ano", [2024, 2025, 2026, 2027, 2028], index=2, key="sel_ano_dash_fuso", label_visibility="collapsed")

    meses_divisor = max(1, date.today().month) if int(ano_painel) == date.today().year else (12 if int(ano_painel) < date.today().year else 1)
    desc_meses_divisor = f"Jan a {ORDEM_MESES_ABREV[meses_divisor - 1]}" if int(ano_painel) == date.today().year else "12 meses"

    col_bg, col_ba, col_bb, col_bl, col_bm = st.columns(5)
    with col_bg:
        if st.button("🌐 Geral (Fábrica)", key="btn_fuso_aba_geral", use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == "Geral" else "secondary"):
            st.session_state.aba_setor_fuso = "Geral"
            st.rerun()
    with col_ba:
        if st.button("🏭 Setor A", key="btn_fuso_aba_sa", use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == "Setor A" else "secondary"):
            st.session_state.aba_setor_fuso = "Setor A"
            st.rerun()
    with col_bb:
        if st.button("🏭 Setor B", key="btn_fuso_aba_sb", use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == "Setor B" else "secondary"):
            st.session_state.aba_setor_fuso = "Setor B"
            st.rerun()
    with col_bl:
        if st.button("🌿 Setor Látex", key="btn_fuso_aba_latex", use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == "Setor Látex" else "secondary"):
            st.session_state.aba_setor_fuso = "Setor Látex"
            st.rerun()
    with col_bm:
        if st.button("⚙️ Setor Menegatto", key="btn_fuso_aba_men", use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == "Setor Menegatto" else "secondary"):
            st.session_state.aba_setor_fuso = "Setor Menegatto"
            st.rerun()

    df_dados_fusos = pd.read_excel(ARQUIVO_FUSOS)
    df_fuso_ano = df_dados_fusos[df_dados_fusos["Ano"] == int(ano_painel)].copy()

    if st.session_state.aba_setor_fuso == "Geral":
        tot_geral = int(df_fuso_ano["Quantidade_Quebras"].sum()) if not df_fuso_ano.empty else 0
        media_geral = round(tot_geral / meses_divisor, 1)

        ultimo_mes_fab = "Nenhum"
        tot_mes_atual = 0
        setor_ofensor, qtd_setor_ofensor = "Nenhum", 0
        if not df_fuso_ano.empty and tot_geral > 0:
            df_reais = df_fuso_ano[df_fuso_ano["Quantidade_Quebras"] > 0]
            for m in reversed(lista_meses_puros):
                df_sub = df_reais[df_reais["Mes"] == m]
                if not df_sub.empty and df_sub["Quantidade_Quebras"].sum() > 0:
                    ultimo_mes_fab = m
                    tot_mes_atual = int(df_sub["Quantidade_Quebras"].sum())
                    agrup = df_sub.groupby("Setor")["Quantidade_Quebras"].sum().sort_values(ascending=False)
                    setor_ofensor = agrup.index[0]
                    qtd_setor_ofensor = int(agrup.iloc[0])
                    break

        kf1, kf2, kf3, kf4 = st.columns(4)
        with kf1: st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Total Fábrica</div><div class='kpi-val'>{tot_geral}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
        with kf2: st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Média Mensal</div><div class='kpi-val' style='color:#059669;'>{media_geral}</div></div><div>📈</div></div>", unsafe_allow_html=True)
        with kf3: st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Setor Crítico ({ultimo_mes_fab})</div><div class='kpi-val' style='color:#d97706; font-size:1.1rem;'>{setor_ofensor} ({qtd_setor_ofensor})</div></div><div>🏭</div></div>", unsafe_allow_html=True)
        with kf4: st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Quebras ({ultimo_mes_fab})</div><div class='kpi-val' style='color:#dc2626;'>{tot_mes_atual}</div></div><div>🚨</div></div>", unsafe_allow_html=True)

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            with st.container(border=True):
                st.subheader("📊 Quebras Totais por Setor")
                if not df_fuso_ano.empty:
                    df_set = df_fuso_ano.groupby("Setor")["Quantidade_Quebras"].sum().reset_index()
                    st.altair_chart(alt.Chart(df_set).mark_bar(color="#3b82f6").encode(x="Setor:N", y="Quantidade_Quebras:Q"), use_container_width=True)
        with col_g2:
            with st.container(border=True):
                st.subheader("🍩 Distribuição por Tipo de Fuso")
                if not df_fuso_ano.empty:
                    df_tip = df_fuso_ano.groupby("Tipo_Fuso")["Quantidade_Quebras"].sum().reset_index()
                    st.altair_chart(alt.Chart(df_tip).mark_arc(innerRadius=60).encode(theta="Quantidade_Quebras:Q", color="Tipo_Fuso:N"), use_container_width=True)
    else:
        setor_ativo = st.session_state.aba_setor_fuso
        df_setor = df_fuso_ano[df_fuso_ano["Setor"] == setor_ativo].copy()
        tot_setor = int(df_setor["Quantidade_Quebras"].sum()) if not df_setor.empty else 0
        media_set = round(tot_setor / meses_divisor, 1)

        sk1, sk2 = st.columns(2)
        with sk1: st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Total ({setor_ativo})</div><div class='kpi-val'>{tot_setor}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
        with sk2: st.markdown(f"<div class='card-kpi-bonito'><div><div class='kpi-lbl'>Média Mensal</div><div class='kpi-val' style='color:#059669;'>{media_set}</div></div><div>📈</div></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.subheader(f"🗺️ Mapa Térmico Operacional — {setor_ativo} (Máquinas vs. 12 Meses)")
            maquinas_lista = DICIONARIO_SETORES.get(setor_ativo, [])
            if not df_setor.empty and len(maquinas_lista) > 0:
                grid_c = []
                for maq in maquinas_lista:
                    for mes in lista_meses_puros:
                        reg_c = df_setor[(df_setor["Maquina_TAG"] == maq) & (df_setor["Mes"] == mes)]
                        grid_c.append({"Maquina_TAG": maq, "Mes": mes, "Mes_Abrev": MAPA_MES_ABREV[mes], "Quantidade": int(reg_c.iloc[0]["Quantidade_Quebras"]) if not reg_c.empty else 0})
                df_heat = pd.DataFrame(grid_c)
                base = alt.Chart(df_heat).encode(x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV), y=alt.Y("Maquina_TAG:N", sort=maquinas_lista))
                rect = base.mark_rect().encode(color=alt.Color("Quantidade:Q", scale=alt.Scale(scheme="reds")))
                txt = base.mark_text().encode(text=alt.condition(alt.datum.Quantidade > 0, "Quantidade:Q", alt.value("")))
                st.altair_chart((rect + txt).properties(height=max(320, len(maquinas_lista) * 18)), use_container_width=True)

# =========================================================================
# 4. LANÇAMENTOS: FUSOS (MODELO DA PLANILHA ANTIGA: DIAS 01 A 31 E TOTAL)
# =========================================================================
elif tela == "Lançamento Fusos":
    st.title("🔩 Lançamento: Quebras Diárias de Fusos")
    st.caption("Visualização idêntica à planilha: Máquinas na vertical e dias de 01 a 31 nas colunas.")

    with st.container(border=True):
        col_ano, col_setor, col_mes = st.columns(3)
        with col_ano:
            ano_selecionado = st.selectbox("📅 Ano:", [2024, 2025, 2026, 2027, 2028], index=2, key="sel_ano_grade_fusos")
        with col_setor:
            setor_selecionado = st.selectbox("🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_setor_grade_fusos")
        with col_mes:
            mes_selecionado = st.selectbox("🗓️ Mês de Referência:", lista_meses_puros, index=date.today().month - 1, key="sel_mes_grade_fusos")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(f"Apontamentos Diários — {setor_selecionado} ({mes_selecionado}/{ano_selecionado})")

    df_atual = pd.read_excel(ARQUIVO_FUSOS)
    df_filtrado = df_atual[(df_atual["Ano"] == ano_selecionado) & (df_atual["Mes"] == mes_selecionado) & (df_atual["Setor"] == setor_selecionado)]
    maquinas_do_setor = obter_maquinas_setor(setor_selecionado, df_correias, df_atual)

    dados_grade = []
    for maq in maquinas_do_setor:
        reg = df_filtrado[df_filtrado["Maquina_TAG"] == maq]
        tot_salvo = int(reg.iloc[0]["Quantidade_Quebras"]) if not reg.empty else 0
        tipo_salvo = str(reg.iloc[0]["Tipo_Fuso"]) if not reg.empty and str(reg.iloc[0]["Tipo_Fuso"]) in OPCOES_TIPO_FUSO else OPCOES_TIPO_FUSO[0]

        linha = {"Num Maq": maq, "Tipo Fuso": tipo_salvo}
        for dia_col in colunas_dias:
            linha[dia_col] = 0
        linha["01"] = tot_salvo
        linha["Total"] = tot_salvo
        dados_grade.append(linha)

    df_grade = pd.DataFrame(dados_grade)
    cfg_colunas = {
        "Num Maq": st.column_config.TextColumn("Num Maq", disabled=True),
        "Tipo Fuso": st.column_config.SelectboxColumn("Tipo Fuso", options=OPCOES_TIPO_FUSO, required=True),
        "Total": st.column_config.NumberColumn("Total", disabled=True, format="%d"),
    }
    for dia_col in colunas_dias:
        cfg_colunas[dia_col] = st.column_config.NumberColumn(dia_col, min_value=0, step=1, format="%d", width="small")

    tabela_editada = st.data_editor(
        df_grade,
        column_config=cfg_colunas,
        hide_index=True,
        use_container_width=True,
        height=540,
        key=f"editor_planilha_fusos_matriz_{setor_selecionado}_{ano_selecionado}_{mes_selecionado}",
    )

    if st.button(f"💾 Salvar Apontamentos ({mes_selecionado}/{ano_selecionado})", key="btn_salvar_matriz_fuso_direto", type="primary"):
        df_limpo = df_atual[~((df_atual["Ano"] == ano_selecionado) & (df_atual["Mes"] == mes_selecionado) & (df_atual["Setor"] == setor_selecionado))]
        novos_regs = []
        for _, linha in tabela_editada.iterrows():
            total_quebras = sum(int(linha[d]) for d in colunas_dias if pd.notna(linha[d]))
            novos_regs.append({
                "Ano": int(ano_selecionado),
                "Mes": mes_selecionado,
                "Setor": setor_selecionado,
                "Maquina_TAG": linha["Num Maq"],
                "Quantidade_Quebras": total_quebras,
                "Tipo_Fuso": str(linha["Tipo Fuso"]),
            })
        pd.concat([df_limpo, pd.DataFrame(novos_regs)], ignore_index=True).to_excel(ARQUIVO_FUSOS, index=False)
        st.success(f"✅ Fechamento de {mes_selecionado}/{ano_selecionado} para o {setor_selecionado} salvo com sucesso!")
        st.rerun()

elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
    st.info("Módulo em desenvolvimento.")

elif tela == "Máquinas":
    st.header("🏭 Cadastro de Máquinas")
    st.info("Módulo em desenvolvimento.")
