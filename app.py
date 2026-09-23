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
# 1. INICIALIZAÇÃO DA BASE DE FUSOS
# ==========================================
df_fusos = None
if os.path.exists(ARQUIVO_FUSOS):
    try:
        df_fusos = pd.read_excel(ARQUIVO_FUSOS)
        if "Dia" not in df_fusos.columns:
            df_fusos["Dia"] = 1
        for col in COLUNAS_FUSOS:
            if col not in df_fusos.columns:
                df_fusos[col] = 0 if col == "Quantidade_Quebras" else ""
    except Exception:
        df_fusos = pd.DataFrame(columns=COLUNAS_FUSOS)
        df_fusos.to_excel(ARQUIVO_FUSOS, index=False)
else:
    df_fusos = pd.DataFrame(columns=COLUNAS_FUSOS)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# ==========================================
# 2. INICIALIZAÇÃO DA BASE DE CORREIAS
# ==========================================
df_correias = None
if os.path.exists(ARQUIVO_CORREIAS):
    try:
        df_correias = pd.read_excel(ARQUIVO_CORREIAS)
    except Exception:
        df_correias = pd.DataFrame(columns=COLUNAS_CORREIAS)
        df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
else:
    df_correias = pd.DataFrame(columns=COLUNAS_CORREIAS)
    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

for col in COLUNAS_CORREIAS:
    if col not in df_correias.columns:
        df_correias[col] = ""
    df_correias[col] = df_correias[col].astype(object)

# =========================================================================
# 3. CARGA FIXA HISTÓRICA CONSOLIDADA DE CORREIAS
# =========================================================================
DADOS_HISTORICOS_CORREIAS = [
    # Setor A
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

    # Setor B
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

    # Setor Látex
    ("Setor Látex", "B-72", "33.990", "2026-01-02", "", ""),
    ("Setor Látex", "B-73", "33.990", "2026-06-13", "34.870", "2026-06-13"),
    ("Setor Látex", "B-74", "33.990", "2026-03-10", "34.870", "2025-02-15"),
    ("Setor Látex", "B-78", "", "", "34.870", "2025-12-29"),
    ("Setor Látex", "B-79", "", "", "34.870", "2024-11-30"),

    # Setor Menegatto
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
        if m1_cor and str(df_correias.loc[idx, "Tipo_Correia_1"]).strip() in ["", "nan", "None"]:
            df_correias.loc[idx, "Tipo_Correia_1"] = str(m1_cor)
            df_correias.loc[idx, "Data_Instalacao_1"] = str(dt1_cor)
            salvar_cor_init = True
        if m2_cor and str(df_correias.loc[idx, "Tipo_Correia_2"]).strip() in ["", "nan", "None"]:
            df_correias.loc[idx, "Tipo_Correia_2"] = str(m2_cor)
            df_correias.loc[idx, "Data_Instalacao_2"] = str(dt2_cor)
            salvar_cor_init = True

if salvar_cor_init:
    gerar_backup_seguro(ARQUIVO_CORREIAS)
    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

df_correias["Tipo_Correia_1"] = df_correias["Tipo_Correia_1"].apply(formatar_modelo)
df_correias["Data_Instalacao_1"] = df_correias["Data_Instalacao_1"].astype(str).replace({"nan": "", "NaT": "", "None": ""})
df_correias["Tipo_Correia_2"] = df_correias["Tipo_Correia_2"].apply(formatar_modelo)
df_correias["Data_Instalacao_2"] = df_correias["Data_Instalacao_2"].astype(str).replace({"nan": "", "NaT": "", "None": ""})

# =========================================================================
# 4. CARGA FIXA HISTÓRICA CONSOLIDADA DE FUSOS
# =========================================================================
dados_janeiro_b = [("L-29", 1), ("L-30", 13), ("L-31", 10), ("L-32", 7), ("L-33", 5), ("L-34", 2), ("L-35", 2), ("L-36", 8), ("L-37", 2), ("L-38", 1), ("L-39", 1), ("L-40", 4), ("L-50", 7), ("L-51", 0), ("L-41", 1), ("L-42", 2), ("L-43", 0), ("L-44", 3), ("L-45", 0), ("L-46", 2), ("L-52", 3), ("L-53", 6)]
dados_fevereiro_b = [("L-29", 4), ("L-30", 23), ("L-31", 9), ("L-32", 34), ("L-33", 22), ("L-34", 19), ("L-35", 42), ("L-36", 22), ("L-37", 7), ("L-38", 4), ("L-39", 5), ("L-40", 11), ("L-50", 45), ("L-51", 2), ("L-41", 1), ("L-42", 2), ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 1), ("L-52", 2), ("L-53", 5)]
dados_marco_b = [("L-29", 1), ("L-30", 15), ("L-31", 4), ("L-32", 8), ("L-33", 14), ("L-34", 6), ("L-35", 8), ("L-36", 9), ("L-37", 18), ("L-38", 2), ("L-39", 5), ("L-40", 6), ("L-50", 0), ("L-51", 4), ("L-41", 1), ("L-42", 1), ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 0), ("L-52", 3), ("L-53", 6)]
dados_abril_b = [("L-29", 2), ("L-30", 3), ("L-31", 1), ("L-32", 2), ("L-33", 17), ("L-34", 8), ("L-35", 7), ("L-36", 15), ("L-37", 17), ("L-38", 5), ("L-39", 12), ("L-40", 3), ("L-50", 2), ("L-51", 3), ("L-41", 3), ("L-42", 4), ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 1), ("L-52", 1), ("L-53", 4)]
dados_maio_b = [("L-29", 0), ("L-30", 5), ("L-31", 8), ("L-32", 0), ("L-33", 6), ("L-34", 2), ("L-35", 7), ("L-36", 10), ("L-37", 12), ("L-38", 2), ("L-39", 5), ("L-40", 3), ("L-50", 1), ("L-51", 0), ("L-41", 2), ("L-42", 1), ("L-43", 8), ("L-44", 1), ("L-45", 0), ("L-46", 2), ("L-52", 5), ("L-53", 5)]
dados_junho_b = [("L-29", 6), ("L-30", 16), ("L-31", 35), ("L-32", 4), ("L-33", 10), ("L-34", 4), ("L-35", 8), ("L-36", 11), ("L-37", 6), ("L-38", 1), ("L-39", 9), ("L-40", 11), ("L-50", 1), ("L-51", 1), ("L-41", 6), ("L-42", 1), ("L-43", 2), ("L-44", 0), ("L-45", 0), ("L-46", 0), ("L-52", 1), ("L-53", 2)]
dados_julho_b = [("L-29", 0), ("L-30", 31), ("L-31", 4), ("L-32", 13), ("L-33", 26), ("L-34", 7), ("L-35", 14), ("L-36", 2), ("L-37", 14), ("L-38", 0), ("L-39", 2), ("L-40", 11), ("L-50", 2), ("L-51", 6), ("L-41", 3), ("L-42", 2), ("L-43", 3), ("L-44", 5), ("L-45", 0), ("L-46", 0), ("L-52", 6), ("L-53", 4)]

mapa_cargas_setor_b = [
    ("Janeiro", dados_janeiro_b), ("Fevereiro", dados_fevereiro_b), ("Março", dados_marco_b),
    ("Abril", dados_abril_b), ("Maio", dados_maio_b), ("Junho", dados_junho_b), ("Julho", dados_julho_b),
]

dados_janeiro_a = [("L-01", 8), ("L-02", 9), ("L-03", 2), ("L-04", 3), ("L-05", 3), ("L-06", 13), ("L-07", 6), ("L-08", 4), ("L-09", 8), ("L-10", 9), ("L-11", 3), ("L-12", 6), ("L-13", 5), ("L-14", 8), ("L-15", 9), ("L-16", 4), ("L-17", 2), ("L-18", 3), ("L-19", 10), ("L-20", 5), ("L-21", 4), ("L-22", 4), ("L-23", 6), ("L-24", 10), ("L-25", 4), ("L-26", 3), ("L-27", 4), ("L-28", 3)]
dados_fevereiro_a = [("L-01", 3), ("L-02", 6), ("L-03", 4), ("L-04", 4), ("L-05", 2), ("L-06", 5), ("L-07", 7), ("L-08", 3), ("L-09", 4), ("L-10", 2), ("L-11", 4), ("L-12", 14), ("L-13", 14), ("L-14", 13), ("L-15", 4), ("L-16", 10), ("L-17", 1), ("L-18", 4), ("L-19", 0), ("L-20", 6), ("L-21", 7), ("L-22", 4), ("L-23", 3), ("L-24", 4), ("L-25", 4), ("L-26", 0), ("L-27", 0), ("L-28", 8)]
dados_marco_a = [("L-01", 4), ("L-02", 7), ("L-03", 6), ("L-04", 3), ("L-05", 1), ("L-06", 5), ("L-07", 10), ("L-08", 2), ("L-09", 12), ("L-10", 8), ("L-11", 10), ("L-12", 9), ("L-13", 9), ("L-14", 9), ("L-15", 12), ("L-16", 3), ("L-17", 8), ("L-18", 20), ("L-19", 10), ("L-20", 3), ("L-21", 4), ("L-22", 5), ("L-23", 3), ("L-24", 8), ("L-25", 4), ("L-26", 9), ("L-27", 2), ("L-28", 2)]
dados_abril_a = [("L-01", 1), ("L-02", 8), ("L-03", 3), ("L-04", 6), ("L-05", 2), ("L-06", 7), ("L-07", 8), ("L-08", 0), ("L-09", 2), ("L-10", 6), ("L-11", 9), ("L-12", 6), ("L-13", 5), ("L-14", 2), ("L-15", 2), ("L-16", 7), ("L-17", 5), ("L-18", 8), ("L-19", 2), ("L-20", 5), ("L-21", 3), ("L-22", 12), ("L-23", 5), ("L-24", 9), ("L-25", 2), ("L-26", 7), ("L-27", 5), ("L-28", 2)]
dados_maio_a = [("L-01", 8), ("L-02", 5), ("L-03", 11), ("L-04", 1), ("L-05", 5), ("L-06", 7), ("L-07", 6), ("L-08", 2), ("L-09", 8), ("L-10", 1), ("L-11", 4), ("L-12", 1), ("L-13", 2), ("L-14", 6), ("L-15", 6), ("L-16", 10), ("L-17", 4), ("L-18", 5), ("L-19", 7), ("L-20", 13), ("L-21", 3), ("L-22", 4), ("L-23", 9), ("L-24", 7), ("L-25", 1), ("L-26", 2), ("L-27", 4), ("L-28", 4)]
dados_junho_a = [("L-01", 5), ("L-02", 2), ("L-03", 4), ("L-04", 4), ("L-05", 5), ("L-06", 3), ("L-07", 8), ("L-08", 1), ("L-09", 6), ("L-10", 5), ("L-11", 9), ("L-12", 5), ("L-13", 8), ("L-14", 3), ("L-15", 2), ("L-16", 5), ("L-17", 2), ("L-18", 4), ("L-19", 3), ("L-20", 5), ("L-21", 5), ("L-22", 3), ("L-23", 4), ("L-24", 4), ("L-25", 3), ("L-26", 9), ("L-27", 3), ("L-28", 1)]
dados_julho_a = [("L-01", 5), ("L-02", 3), ("L-03", 3), ("L-04", 10), ("L-05", 2), ("L-06", 6), ("L-07", 11), ("L-08", 8), ("L-09", 11), ("L-10", 7), ("L-11", 5), ("L-12", 3), ("L-13", 2), ("L-14", 7), ("L-15", 6), ("L-16", 9), ("L-17", 4), ("L-18", 8), ("L-19", 8), ("L-20", 5), ("L-21", 11), ("L-22", 18), ("L-23", 9), ("L-24", 10), ("L-25", 7), ("L-26", 9), ("L-27", 3), ("L-28", 9)]
dados_agosto_a = [("L-01", 5), ("L-02", 8), ("L-03", 3), ("L-04", 5), ("L-05", 11), ("L-06", 8), ("L-07", 15), ("L-08", 5), ("L-09", 14), ("L-10", 6), ("L-11", 9), ("L-12", 5), ("L-13", 7), ("L-14", 9), ("L-15", 11), ("L-16", 5), ("L-17", 6), ("L-18", 12), ("L-19", 10), ("L-20", 10), ("L-21", 9), ("L-22", 16), ("L-23", 11), ("L-24", 8), ("L-25", 4), ("L-26", 2), ("L-27", 5), ("L-28", 10)]
dados_setembro_a = [("L-01", 1), ("L-02", 2), ("L-03", 2), ("L-04", 4), ("L-05", 1), ("L-06", 0), ("L-07", 1), ("L-08", 0), ("L-09", 3), ("L-10", 4), ("L-11", 7), ("L-12", 2), ("L-13", 4), ("L-14", 3), ("L-15", 1), ("L-16", 0), ("L-17", 1), ("L-18", 1), ("L-19", 4), ("L-20", 4), ("L-21", 2), ("L-22", 3), ("L-23", 1), ("L-24", 5), ("L-25", 1), ("L-26", 0), ("L-27", 2), ("L-28", 4)]

mapa_cargas_setor_a = [
    ("Janeiro", dados_janeiro_a), ("Fevereiro", dados_fevereiro_a), ("Março", dados_marco_a),
    ("Abril", dados_abril_a), ("Maio", dados_maio_a), ("Junho", dados_junho_a),
    ("Julho", dados_julho_a), ("Agosto", dados_agosto_a), ("Setembro", dados_setembro_a),
]

dados_janeiro_latex = [("B-71", 7), ("B-72", 10), ("B-73", 13), ("B-74", 18), ("B-75", 4), ("B-76", 9), ("B-77", 10), ("B-78", 5), ("B-79", 1), ("B-80", 0), ("B-83", 6), ("B-84", 7), ("B-85", 9), ("B-86", 5), ("B-87", 5), ("B-88", 10), ("B-89", 1), ("B-102", 0), ("B-103", 0), ("B-104", 0)]
dados_fevereiro_latex = [("B-71", 1), ("B-72", 2), ("B-73", 5), ("B-74", 0), ("B-75", 0), ("B-76", 1), ("B-77", 2), ("B-78", 3), ("B-79", 4), ("B-80", 0), ("B-83", 1), ("B-84", 2), ("B-85", 4), ("B-86", 4), ("B-87", 2), ("B-88", 0), ("B-89", 0), ("B-102", 0), ("B-103", 0), ("B-104", 0)]
dados_marco_latex = [("B-71", 5), ("B-72", 7), ("B-73", 2), ("B-74", 1), ("B-75", 0), ("B-76", 3), ("B-77", 4), ("B-78", 4), ("B-79", 0), ("B-80", 1), ("B-83", 9), ("B-84", 18), ("B-85", 9), ("B-86", 9), ("B-87", 14), ("B-88", 0), ("B-89", 8), ("B-102", 0), ("B-103", 0), ("B-104", 0)]
dados_abril_latex = [("B-71", 5), ("B-72", 9), ("B-73", 3), ("B-74", 5), ("B-75", 2), ("B-76", 4), ("B-77", 5), ("B-78", 3), ("B-79", 0), ("B-80", 0), ("B-83", 10), ("B-84", 18), ("B-85", 8), ("B-86", 12), ("B-87", 9), ("B-88", 0), ("B-89", 10), ("B-102", 0), ("B-103", 0), ("B-104", 0)]
dados_maio_latex = [("B-71", 1), ("B-72", 4), ("B-73", 7), ("B-74", 5), ("B-75", 2), ("B-76", 4), ("B-77", 1), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 15), ("B-84", 18), ("B-85", 6), ("B-86", 12), ("B-87", 7), ("B-88", 0), ("B-89", 17), ("B-102", 0), ("B-103", 0), ("B-104", 1)]
dados_junho_latex = [("B-71", 1), ("B-72", 1), ("B-73", 7), ("B-74", 0), ("B-75", 7), ("B-76", 1), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 9), ("B-84", 7), ("B-85", 10), ("B-86", 13), ("B-87", 5), ("B-88", 1), ("B-89", 8), ("B-102", 0), ("B-103", 0), ("B-104", 0)]
dados_julho_latex = [("B-71", 3), ("B-72", 1), ("B-73", 3), ("B-74", 0), ("B-75", 0), ("B-76", 0), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 8), ("B-84", 10), ("B-85", 7), ("B-86", 9), ("B-87", 10), ("B-88", 0), ("B-89", 7), ("B-102", 0), ("B-103", 0), ("B-104", 0)]
dados_agosto_latex = [("B-71", 5), ("B-72", 1), ("B-73", 0), ("B-74", 0), ("B-75", 0), ("B-76", 0), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0), ("B-83", 1), ("B-84", 3), ("B-85", 6), ("B-86", 1), ("B-87", 1), ("B-88", 0), ("B-89", 1), ("B-102", 0), ("B-103", 0), ("B-104", 0)]

mapa_cargas_setor_latex = [
    ("Janeiro", dados_janeiro_latex), ("Fevereiro", dados_fevereiro_latex), ("Março", dados_marco_latex),
    ("Abril", dados_abril_latex), ("Maio", dados_maio_latex), ("Junho", dados_junho_latex),
    ("Julho", dados_julho_latex), ("Agosto", dados_agosto_latex),
]

dados_janeiro_menegatto = [("B-47", 4), ("B-48", 3), ("B-49", 5), ("B-81", 3), ("B-82", 3), ("B-93", 10), ("B-94", 9), ("B-95", 4), ("B-96", 6), ("B-97", 4), ("B-98", 1), ("B-99", 6), ("B-100", 6), ("B-101", 5), ("B-107", 0), ("B-108", 0)]
dados_fevereiro_menegatto = [("B-47", 4), ("B-48", 3), ("B-49", 2), ("B-81", 2), ("B-82", 0), ("B-93", 1), ("B-94", 0), ("B-95", 3), ("B-96", 4), ("B-97", 6), ("B-98", 1), ("B-99", 6), ("B-100", 16), ("B-101", 1), ("B-107", 0), ("B-108", 0)]
dados_marco_menegatto = [("B-47", 0), ("B-48", 5), ("B-49", 2), ("B-81", 3), ("B-82", 3), ("B-93", 3), ("B-94", 7), ("B-95", 12), ("B-96", 2), ("B-97", 5), ("B-98", 1), ("B-99", 9), ("B-100", 0), ("B-101", 3), ("B-107", 0), ("B-108", 0)]
dados_abril_menegatto = [("B-47", 9), ("B-48", 8), ("B-49", 3), ("B-81", 8), ("B-82", 5), ("B-93", 3), ("B-94", 2), ("B-95", 25), ("B-96", 5), ("B-97", 3), ("B-98", 2), ("B-99", 14), ("B-100", 2), ("B-101", 4), ("B-107", 0), ("B-108", 0)]
dados_maio_menegatto = [("B-47", 5), ("B-48", 9), ("B-49", 1), ("B-81", 7), ("B-82", 1), ("B-93", 1), ("B-94", 0), ("B-95", 4), ("B-96", 0), ("B-97", 9), ("B-98", 0), ("B-99", 1), ("B-100", 4), ("B-101", 3), ("B-107", 0), ("B-108", 0)]
dados_junho_menegatto = [("B-47", 4), ("B-48", 6), ("B-49", 1), ("B-81", 1), ("B-82", 3), ("B-93", 5), ("B-94", 7), ("B-95", 15), ("B-96", 5), ("B-97", 9), ("B-98", 0), ("B-99", 6), ("B-100", 3), ("B-101", 1), ("B-107", 0), ("B-108", 0)]
dados_julho_menegatto = [("B-47", 11), ("B-48", 5), ("B-49", 7), ("B-81", 6), ("B-82", 5), ("B-93", 1), ("B-94", 0), ("B-95", 1), ("B-96", 3), ("B-97", 1), ("B-98", 1), ("B-99", 3), ("B-100", 2), ("B-101", 5), ("B-107", 1), ("B-108", 0)]
dados_agosto_menegatto = [("B-47", 7), ("B-48", 1), ("B-49", 4), ("B-81", 0), ("B-82", 3), ("B-93", 2), ("B-94", 0), ("B-95", 1), ("B-96", 2), ("B-97", 3), ("B-98", 0), ("B-99", 0), ("B-100", 1), ("B-101", 2), ("B-107", 0), ("B-108", 0)]

mapa_cargas_setor_menegatto = [
    ("Janeiro", dados_janeiro_menegatto), ("Fevereiro", dados_fevereiro_menegatto), ("Março", dados_marco_menegatto),
    ("Abril", dados_abril_menegatto), ("Maio", dados_maio_menegatto), ("Junho", dados_junho_menegatto),
    ("Julho", dados_julho_menegatto), ("Agosto", dados_agosto_menegatto),
]

precisa_salvar_fusos = False

for s_alvo, mapa_c, fuso_padrao in [
    ("Setor B", mapa_cargas_setor_b, "TEP"),
    ("Setor A", mapa_cargas_setor_a, "FAG"),
    ("Setor Látex", mapa_cargas_setor_latex, "M4BA"),
    ("Setor Menegatto", mapa_cargas_setor_menegatto, "MENEGATTO"),
]:
    for nome_mes, lista_d in mapa_c:
        mask_m = (df_fusos["Ano"] == 2026) & (df_fusos["Mes"] == nome_mes) & (df_fusos["Setor"] == s_alvo)
        sub = df_fusos[mask_m]
        if sub.empty or sub["Quantidade_Quebras"].sum() == 0:
            df_fusos = df_fusos[~mask_m]
            novos = [
                {"Ano": 2026, "Mes": nome_mes, "Dia": 1, "Setor": s_alvo, "Maquina_TAG": mq, "Quantidade_Quebras": int(q), "Tipo_Fuso": fuso_padrao}
                for mq, q in lista_d
            ]
            df_fusos = pd.concat([df_fusos, pd.DataFrame(novos)], ignore_index=True)
            precisa_salvar_fusos = True

if precisa_salvar_fusos:
    gerar_backup_seguro(ARQUIVO_FUSOS)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

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
# PROCESSAMENTO DE CORREIAS
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

        /* Estilização dos botões das máquinas */
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

    # Detalhe / HUD da Máquina Clicada
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

    # Mosaico de Máquinas Exibindo Todos os Setores Concomitantemente
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

    # VISÃO GERAL (FÁBRICA COMPLETA)
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

    # VISÃO DETALHADA POR SETOR
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
                color=alt.condition("datum.Quantidade_Quebras >= 10", alt.value("#ffffff"), alt.value("#0f172a"))
            )
            st.altair_chart((rect + txt).properties(height=max(320, len(maqs_setor) * 23)), use_container_width=True)

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

        tipos_disp = sorted(df_sa["Tipo_Fuso"].dropna().unique().tolist()) if not df_sa.empty else OPCOES_TIPO_FUSO
        if not tipos_disp:
            tipos_disp = OPCOES_TIPO_FUSO

        with st.container(border=True):
            c_dt1, c_dt2 = st.columns([2.5, 1.5])
            with c_dt1:
                st.markdown(f"<div style='font-size:1.02rem; font-weight:800;'>🔬 Desempenho Operacional por Tipo de Fuso — {s_ativo}</div>", unsafe_allow_html=True)
            with c_dt2:
                tipo_sel_analise = st.selectbox("Selecione o Tipo de Fuso:", tipos_disp, key=f"sel_tipo_diag_{s_ativo}")

            df_tipo_esp = df_sa[df_sa["Tipo_Fuso"] == tipo_sel_analise].copy()
            tot_f_tipo = int(df_tipo_esp["Quantidade_Quebras"].sum()) if not df_tipo_esp.empty else 0
            med_f_tipo = round(tot_f_tipo / div_meses, 1)

            maqs_tipo = sorted(df_tipo_esp["Maquina_TAG"].unique().tolist()) if not df_tipo_esp.empty else []
            falhas_tipo = df_tipo_esp[df_tipo_esp["Quantidade_Quebras"] > 0]
            maqs_falharam_tipo = falhas_tipo["Maquina_TAG"].nunique()

            if not falhas_tipo.empty:
                agrup_top_t = falhas_tipo.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().sort_values(ascending=False)
                top_maq_t = agrup_top_t.index[0]
                qtd_top_t = int(agrup_top_t.iloc[0])
            else:
                top_maq_t, qtd_top_t = "Nenhuma", 0

            cf_k1, cf_k2, cf_k3, cf_k4 = st.columns(4)
            cf_k1.markdown(f"<div class='card-kpi-bonito c-total' style='height:58px;'><div><div class='kpi-lbl'>Total ({tipo_sel_analise})</div><div class='kpi-val'>{tot_f_tipo} un.</div></div><div>🏷️</div></div>", unsafe_allow_html=True)
            cf_k2.markdown(f"<div class='card-kpi-bonito c-ok' style='height:58px;'><div><div class='kpi-lbl'>Média Mensal</div><div class='kpi-val' style='color:#059669;'>{med_f_tipo} /mês</div></div><div>📉</div></div>", unsafe_allow_html=True)
            cf_k3.markdown(f"<div class='card-kpi-bonito c-warn' style='height:58px;'><div><div class='kpi-lbl'>Máquinas Atreladas</div><div class='kpi-val' style='color:#d97706;'>{maqs_falharam_tipo} falharam / {len(maqs_tipo)}</div></div><div>⚙️</div></div>", unsafe_allow_html=True)
            cf_k4.markdown(f"<div class='card-kpi-bonito c-crit' style='height:58px;'><div><div class='kpi-lbl'>Maior Ofensor</div><div class='kpi-val' style='color:#dc2626; font-size:1.05rem;'>{top_maq_t} ({qtd_top_t})</div></div><div>🚨</div></div>", unsafe_allow_html=True)

            st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

            if not maqs_tipo:
                st.info(f"Nenhuma máquina possui registros vinculados ao fuso {tipo_sel_analise} no {s_ativo} em {ano_f}.")
            else:
                grid_tipo = pd.MultiIndex.from_product([maqs_tipo, LISTA_MESES_PUROS], names=["MAQ", "Mes"]).to_frame().reset_index(drop=True)
                grid_tipo["MES"] = grid_tipo["Mes"].map(MAPA_MES_ABREV)
                agrup_esp = df_tipo_esp.groupby(["Maquina_TAG", "Mes"])["Quantidade_Quebras"].sum().reset_index()
                m_calor_esp = pd.merge(grid_tipo, agrup_esp, left_on=["MAQ", "Mes"], right_on=["Maquina_TAG", "Mes"], how="left").fillna(0)

                rect_t = alt.Chart(m_calor_esp).mark_rect(stroke="#fff", strokeWidth=1).encode(
                    x=alt.X("MES:N", sort=ORDEM_MESES_ABREV, title="Mês", axis=alt.Axis(orient="top", labelAngle=0, labelFontWeight="bold")),
                    y=alt.Y("MAQ:N", sort=maqs_tipo, title="Máquina", axis=alt.Axis(labelFontWeight="bold")),
                    color=alt.Color("Quantidade_Quebras:Q", scale=alt.Scale(domain=[0, 3, 8, 15], range=["#dcfce7", "#fef08a", "#f97316", "#dc2626"]), legend=alt.Legend(title="Quebras")),
                    tooltip=[alt.Tooltip("MAQ:N", title="Máquina"), alt.Tooltip("MES:N", title="Mês"), alt.Tooltip("Quantidade_Quebras:Q", title="Quebras")]
                )
                txt_t = alt.Chart(m_calor_esp).mark_text(baseline="middle", fontSize=11, fontWeight=700).encode(
                    x=alt.X("MES:N", sort=ORDEM_MESES_ABREV),
                    y=alt.Y("MAQ:N", sort=maqs_tipo),
                    text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")),
                    color=alt.condition("datum.Quantidade_Quebras >= 10", alt.value("#ffffff"), alt.value("#0f172a"))
                )
                st.altair_chart((rect_t + txt_t).properties(height=max(220, len(maqs_tipo) * 23)), use_container_width=True)

# ------------------------------------------
# 3. PAINEL GERENCIAL DE SETORES (NOVO)
# ------------------------------------------
elif tela == "Painel Setores":
    st.markdown("<h2 style='margin:0; font-weight:900;'>🏭 Painel Executivo de Setores</h2>", unsafe_allow_html=True)
    st.caption("Visão consolidada e comparativa de desempenho operacional entre todos os setores da fábrica.")

    # Matriz de indicadores por setor
    dados_resumo_setores = []
    for s_nome in DICIONARIO_SETORES.keys():
        maqs_s = obter_maquinas_setor(s_nome, df_correias, df_fusos)
        tot_q_fusos = int(df_fusos[df_fusos["Setor"] == s_nome]["Quantidade_Quebras"].sum()) if not df_fusos.empty else 0
        
        # Correias críticas do setor
        crit_cor = len([r for r in lista_correias_criticas if r["setor"] == s_nome])
        novas_cor = len([r for r in lista_correias_novas if r["setor"] == s_nome])
        meia_cor = len([r for r in lista_correias_meia if r["setor"] == s_nome])

        dados_resumo_setores.append({
            "Setor": s_nome,
            "Total Máquinas": len(maqs_s),
            "Quebras Fusos (Total)": tot_q_fusos,
            "Correias Críticas (Troca)": crit_cor,
            "Correias Meia-Vida": meia_cor,
            "Correias Novas": novas_cor,
        })

    df_res_setores = pd.DataFrame(dados_resumo_setores)

    # 4 Cards superiores de resumo geral
    cs1, cs2, cs3, cs4 = st.columns(4)
    cs1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total de Setores</div><div class='kpi-val'>{len(DICIONARIO_SETORES)}</div></div><div>🏢</div></div>", unsafe_allow_html=True)
    cs2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Parque de Máquinas</div><div class='kpi-val' style='color:#059669;'>{len(todas_maquinas_totais)}</div></div><div>⚙️</div></div>", unsafe_allow_html=True)
    cs3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Correias Meia-Vida</div><div class='kpi-val' style='color:#d97706;'>{len(lista_correias_meia)}</div></div><div>🟡</div></div>", unsafe_allow_html=True)
    cs4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Total Correias Críticas</div><div class='kpi-val' style='color:#dc2626;'>{len(lista_correias_criticas)}</div></div><div>🚨</div></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    c_gset1, c_gset2 = st.columns(2)
    with c_gset1:
        with st.container(border=True):
            st.markdown("<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🔩 Total de Quebras de Fusos por Setor</div>", unsafe_allow_html=True)
            chart_s_fuso = alt.Chart(df_res_setores).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#2563eb").encode(
                x=alt.X("Setor:N", title=None, axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                y=alt.Y("Quebras Fusos (Total):Q", title="Quebras"),
                tooltip=["Setor", "Quebras Fusos (Total)"]
            )
            txt_s_fuso = chart_s_fuso.mark_text(dy=-8, fontSize=11, fontWeight=700).encode(text="Quebras Fusos (Total):Q")
            st.altair_chart((chart_s_fuso + txt_s_fuso).properties(height=240), use_container_width=True)

    with c_gset2:
        with st.container(border=True):
            st.markdown("<div style='font-size:0.95rem; font-weight:800; margin-bottom:8px;'>🚨 Correias com Troca Necessária por Setor</div>", unsafe_allow_html=True)
            chart_s_cor = alt.Chart(df_res_setores).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#dc2626").encode(
                x=alt.X("Setor:N", title=None, axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                y=alt.Y("Correias Críticas (Troca):Q", title="Críticas"),
                tooltip=["Setor", "Correias Críticas (Troca)"]
            )
            txt_s_cor = chart_s_cor.mark_text(dy=-8, fontSize=11, fontWeight=700).encode(text="Correias Críticas (Troca):Q")
            st.altair_chart((chart_s_cor + txt_s_cor).properties(height=240), use_container_width=True)

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("<div style='font-size:0.95rem; font-weight:800; margin-bottom:6px;'>📋 Matriz Comparativa de Setores</div>", unsafe_allow_html=True)
        st.dataframe(df_res_setores, use_container_width=True, hide_index=True)

# ------------------------------------------
# 4. PAINEL GERENCIAL DE MÁQUINAS (NOVO)
# ------------------------------------------
elif tela == "Painel Maquinas":
    st.markdown("<h2 style='margin:0; font-weight:900;'>⚙️ Prontuário Individual da Máquina</h2>", unsafe_allow_html=True)
    st.caption("Consulte o histórico detalhado, dados de correias e registros de quebras por TAG de máquina.")

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

        # Cards de Visão Geral da Máquina
        cm1, cm2, cm3, cm4 = st.columns(4)
        cm1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Setor Ativo</div><div class='kpi-val' style='font-size:1.05rem;'>{setor_selecionado_maq}</div></div><div>🏭</div></div>", unsafe_allow_html=True)
        cm2.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Status Correia</div><div class='kpi-val' style='font-size:1.05rem;'>{info_cor_maq.get('dot', '⚪')} {info_cor_maq.get('status_label', 'Sem Dados')}</div></div><div>🔄</div></div>", unsafe_allow_html=True)
        cm3.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Total Quebras Fusos</div><div class='kpi-val' style='color:#dc2626;'>{tot_falhas_maq}</div></div><div>🔩</div></div>", unsafe_allow_html=True)
        cm4.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Tipo Fuso Padrão</div><div class='kpi-val' style='font-size:1.05rem;'>{sub_fusos_maq['Tipo_Fuso'].iloc[0] if not sub_fusos_maq.empty else 'N/A'}</div></div><div>🏷️</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        # Dados das Correias da Máquina
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

        # Histórico de Quebras de Fusos da Máquina
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

    tab_fusos_db, tab_correias_db, tab_backups_db = st.tabs(["🔩 Base de Fusos", "🔄 Base de Correias", "🛡️ Histórico de Backups"])

    # Aba Fusos: 1 único arquivo com todos os 12 meses do ano
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
            st.caption(f"Descarrega um ficheiro Excel com 12 abas (Jan a Dez) de {ano_db_fuso} para o {setor_db_fuso}.")
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
            st.caption(f"Envie um ficheiro Excel (.xlsx) contendo abas dos meses (`Janeiro`, `Fevereiro`, etc.) no modelo `MAQUINA | 1 | 2 ... 31`.")
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
                            st.success(f"✅ {len(meses_atualizados)} meses atualizados com sucesso ({', '.join(meses_atualizados[:4])}...)!")
                            st.rerun()
                        else:
                            st.warning("Nenhuma aba com nome de mês correspondente ou coluna 'MAQUINA' foi encontrada.")

                    except Exception as erro_up:
                        st.error(f"Erro ao processar o ficheiro anual: {erro_up}")

    # Aba Correias: Estritamente Download e Upload de Dados
    with tab_correias_db:
        st.markdown("### 🔄 Troca de Dados de Correias")
        st.caption("Descarregue a planilha modelo com todas as máquinas cadastradas ou envie novos dados atualizados.")

        col_d_cor, col_u_cor = st.columns([1.5, 2.5])

        # 1. DOWNLOAD DA BASE DE CORREIAS
        with col_d_cor:
            st.markdown("#### 📥 Descarregar Planilha")
            st.caption("Gera um arquivo .xlsx limpo pronto para preenchimento de todas as máquinas.")

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

        # 2. UPLOAD E ATUALIZAÇÃO DA BASE DE CORREIAS
        with col_u_cor:
            st.markdown("#### 📤 Enviar Dados Atualizados")
            st.caption("Suba o arquivo Excel preenchido (.xlsx). Colunas esperadas: `Setor`, `Maquina_TAG`, `Tipo_Correia_1`, `Data_Instalacao_1`, `Tipo_Correia_2`, `Data_Instalacao_2`.")

            up_arquivo_cor = st.file_uploader(
                "Carregar nova planilha de correias (.xlsx)",
                type=["xlsx"],
                key="uploader_novas_correias",
            )

            modo_gravacao = st.radio(
                "Modo de Atualização:",
                [
                    "Mesclar e Atualizar (Preserva outras máquinas e atualiza as enviadas)",
                    "Substituição Completa (Sobrescreve toda a base atual)"
                ],
                key="radio_modo_up_cor"
            )

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

                        colunas_obrigatorias = ["Setor", "Maquina_TAG"]
                        if not all(col in df_novo_cor.columns for col in colunas_obrigatorias):
                            st.error("❌ O arquivo precisa conter pelo menos as colunas 'Setor' e 'Maquina_TAG' (ou 'Máquina').")
                        else:
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
                                mascara_manter = [
                                    (r["Setor"], r["Maquina_TAG"]) not in chaves_enviadas
                                    for _, r in df_correias.iterrows()
                                ]
                                df_base_restante = df_correias[mascara_manter]
                                df_final_up_c = pd.concat([df_base_restante, df_novo_cor], ignore_index=True)

                            df_final_up_c.to_excel(ARQUIVO_CORREIAS, index=False)
                            st.success(f"✅ Base de Correias atualizada com sucesso! ({len(df_novo_cor)} máquinas processadas)")
                            st.rerun()

                    except Exception as erro_proc:
                        st.error(f"Erro ao processar o arquivo de correias: {erro_proc}")

    # Aba Backups Automáticos
    with tab_backups_db:
        st.markdown("#### Cópias de Segurança Geradas Automaticamente")
        st.caption("Cada vez que uma alteração é salva no sistema, uma cópia completa com data e hora é arquivada.")

        if os.path.exists("backups"):
            arquivos_bkp = sorted(os.listdir("backups"), reverse=True)
            if arquivos_bkp:
                dados_bkp = []
                for bkp in arquivos_bkp:
                    caminho_bkp = os.path.join("backups", bkp)
                    tam_kb = round(os.path.getsize(caminho_bkp) / 1024, 1)
                    dados_bkp.append({"Arquivo de Backup": bkp, "Tamanho": f"{tam_kb} KB"})
                st.dataframe(pd.DataFrame(dados_bkp), use_container_width=True, height=300)
            else:
                st.info("Nenhum backup gerado ainda.")
        else:
            st.info("Pasta de backups ainda não inicializada.")

# ------------------------------------------
# 6. GESTÃO CADASTRAL DE MÁQUINAS (SISTEMA & DADOS) (NOVO)
# ------------------------------------------
elif tela == "Gestao Maquinas":
    st.markdown("<h2 style='margin:0; font-weight:900;'>🏭 Gestão Cadastral de Máquinas</h2>", unsafe_allow_html=True)
    st.caption("Cadastre novas máquinas no parque de ativos ou desative TAGs dos setores.")

    c_cad_s, _ = st.columns([2, 3])
    with c_cad_s:
        setor_gerenc = st.selectbox("Selecione o Setor:", list(DICIONARIO_SETORES.keys()), key="sel_setor_gestao_maq")

    maqs_atuais_gestao = obter_maquinas_setor(setor_gerenc, df_correias, df_fusos)

    st.markdown("---")
    col_add, col_del = st.columns(2)

    with col_add:
        with st.container(border=True):
            st.markdown(f"#### ➕ Adicionar Nova Máquina — {setor_gerenc}")
            st.caption("A nova máquina será incorporada nas bases de fusos e correias.")
            nova_tag_input = st.text_input("Código / TAG da Máquina:", placeholder="Ex: L-54 ou B-109", key="inp_nova_tag_gerenc").strip().upper()

            if st.button("Cadastrar Máquina", type="primary", key="btn_cadastrar_maq_gerenc"):
                if not nova_tag_input:
                    st.warning("Informe o código da máquina.")
                elif nova_tag_input in maqs_atuais_gestao:
                    st.warning(f"A máquina {nova_tag_input} já existe no {setor_gerenc}.")
                else:
                    # Adiciona à base de correias se não existir
                    mask_c = (df_correias["Setor"] == setor_gerenc) & (df_correias["Maquina_TAG"] == nova_tag_input)
                    if not mask_c.any():
                        novo_cor = {
                            "Setor": setor_gerenc, "Maquina_TAG": nova_tag_input,
                            "Tipo_Correia_1": "", "Data_Instalacao_1": "",
                            "Tipo_Correia_2": "", "Data_Instalacao_2": "",
                        }
                        df_correias = pd.concat([df_correias, pd.DataFrame([novo_cor])], ignore_index=True)
                        gerar_backup_seguro(ARQUIVO_CORREIAS)
                        df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

                    # Adiciona à base de fusos para o ano vigente
                    fuso_padrao = "FAG" if setor_gerenc == "Setor A" else "TEP" if setor_gerenc == "Setor B" else "M4BA" if setor_gerenc == "Setor Látex" else "MENEGATTO"
                    mask_f = (df_fusos["Setor"] == setor_gerenc) & (df_fusos["Maquina_TAG"] == nova_tag_input)
                    if not mask_f.any():
                        novos_fusos_ano = [
                            {"Ano": 2026, "Mes": m_nome, "Dia": 1, "Setor": setor_gerenc, "Maquina_TAG": nova_tag_input, "Quantidade_Quebras": 0, "Tipo_Fuso": fuso_padrao}
                            for m_nome in LISTA_MESES_PUROS
                        ]
                        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_fusos_ano)], ignore_index=True)
                        gerar_backup_seguro(ARQUIVO_FUSOS)
                        df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

                    st.success(f"✅ Máquina {nova_tag_input} cadastrada com sucesso no {setor_gerenc}!")
                    st.rerun()

    with col_del:
        with st.container(border=True):
            st.markdown(f"#### 🗑️ Desativar / Remover Máquina — {setor_gerenc}")
            st.caption("Remove a TAG e desvincula os apontamentos existentes das bases.")
            tag_del_sel = st.selectbox("Selecione a TAG para remover:", ["-- Selecione --"] + maqs_atuais_gestao, key="sel_tag_excluir_gerenc")

            if st.button("Remover Máquina", type="secondary", key="btn_remover_maq_gerenc"):
                if tag_del_sel and tag_del_sel != "-- Selecione --":
                    gerar_backup_seguro(ARQUIVO_CORREIAS)
                    df_correias = df_correias[~((df_correias["Setor"] == setor_gerenc) & (df_correias["Maquina_TAG"] == tag_del_sel))]
                    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

                    gerar_backup_seguro(ARQUIVO_FUSOS)
                    df_fusos = df_fusos[~((df_fusos["Setor"] == setor_gerenc) & (df_fusos["Maquina_TAG"] == tag_del_sel))]
                    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

                    st.success(f"🗑️ Máquina {tag_del_sel} removida com sucesso do {setor_gerenc}!")
                    st.rerun()
                else:
                    st.warning("Selecione uma máquina válida para remover.")

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown(f"#### 📋 Lista de Máquinas Ativas — {setor_gerenc} ({len(maqs_atuais_gestao)} no total)")
        st.write(", ".join(maqs_atuais_gestao))
