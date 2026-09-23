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

# Ficheiros de dados blindados e separados
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
# INICIALIZAÇÃO E CARGA DAS BASES
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

df_correias = None
if os.path.exists(ARQUIVO_CORREIAS):
    try:
        df_correias = pd.read_excel(ARQUIVO_CORREIAS)
    except Exception:
        df_correias = pd.DataFrame(columns=COLUNAS_CORREIAS)
        df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
else:
    if os.path.exists("lancamentos_correias_v3.xlsx"):
        try:
            df_antigo = pd.read_excel("lancamentos_correias_v3.xlsx")
            df_migrado = pd.DataFrame(columns=COLUNAS_CORREIAS)
            df_migrado["Setor"] = df_antigo.get("Setor", "")
            df_migrado["Maquina_TAG"] = df_antigo.get("Maquina_TAG", "")
            df_migrado["Tipo_Correia_1"] = df_antigo.get("Tipo_Correia", "")
            df_migrado["Data_Instalacao_1"] = df_antigo.get("Data_Instalacao", "")
            df_migrado["Tipo_Correia_2"] = ""
            df_migrado["Data_Instalacao_2"] = ""
            df_migrado.to_excel(ARQUIVO_CORREIAS, index=False)
            df_correias = df_migrado
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
# CARGA FIXA HISTÓRICA CONSOLIDADA DE CORREIAS
# =========================================================================
DADOS_HISTORICOS_CORREIAS = [
    # --- Setor A ---
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

    # --- Setor B ---
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

    # --- Setor Látex ---
    ("Setor Látex", "B-72", "33.990", "2026-01-02", "", ""),
    ("Setor Látex", "B-73", "33.990", "2026-06-13", "34.870", "2026-06-13"),
    ("Setor Látex", "B-74", "33.990", "2026-03-10", "34.870", "2025-02-15"),
    ("Setor Látex", "B-78", "", "", "34.870", "2025-12-29"),
    ("Setor Látex", "B-79", "", "", "34.870", "2024-11-30"),

    # --- Setor Menegatto ---
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
        if m1_cor:
            df_correias.loc[idx, "Tipo_Correia_1"] = str(m1_cor)
            df_correias.loc[idx, "Data_Instalacao_1"] = str(dt1_cor)
            salvar_cor_init = True
        elif dt1_cor and str(df_correias.loc[idx, "Data_Instalacao_1"]).strip() in ["", "nan", "None", "NaT"]:
            df_correias.loc[idx, "Data_Instalacao_1"] = str(dt1_cor)
            salvar_cor_init = True

        if m2_cor:
            df_correias.loc[idx, "Tipo_Correia_2"] = str(m2_cor)
            df_correias.loc[idx, "Data_Instalacao_2"] = str(dt2_cor)
            salvar_cor_init = True
        elif dt2_cor and str(df_correias.loc[idx, "Data_Instalacao_2"]).strip() in ["", "nan", "None", "NaT"]:
            df_correias.loc[idx, "Data_Instalacao_2"] = str(dt2_cor)
            salvar_cor_init = True

if salvar_cor_init:
    gerar_backup_seguro(ARQUIVO_CORREIAS)
    df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

df_correias["Tipo_Correia_1"] = df_correias["Tipo_Correia_1"].apply(formatar_modelo)
df_correias["Data_Instalacao_1"] = df_correias["Data_Instalacao_1"].astype(str).replace({"nan": "", "NaT": "", "None": ""})
df_correias["Tipo_Correia_2"] = df_correias["Tipo_Correia_2"].apply(formatar_modelo)
df_correias["Data_Instalacao_2"] = df_correias["Data_Instalacao_2"].astype(str).replace({"nan": "", "NaT": "", "None": ""})

# ==========================================
# DADOS HISTÓRICOS DE FUSOS
# ==========================================
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

# ------------------------------------------
# VISÃO GERAL (FÁBRICA COMPLETA)
# ------------------------------------------
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

    # ------------------------------------------
    # VISÃO DETALHADA POR SETOR (COM TODOS OS GRÁFICOS)
    # ------------------------------------------
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

        # 1. Gráficos de Evolução Mensal e Distribuição por Marca/Tipo
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

        # 2. Mapa de Calor Operacional do Setor
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

        # 3. Diagnóstico e Mapa Térmico Específico por Marca / Tipo de Fuso
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
# 4. LANÇAMENTO DE FUSOS (MATRIZ DIÁRIA)
# ------------------------------------------
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

# ------------------------------------------
# 5. BANCO DE DADOS & SEGURANÇA
# ------------------------------------------
elif tela == "Banco de Dados":
    st.title("🗄️ Banco de Dados & Gestão de Backups")
    st.caption("Visualização crua dos registros, downloads manuais, restaurações e histórico de backups automáticos")

    tab_fusos_db, tab_correias_db, tab_backups_db = st.tabs(["🔩 Base de Fusos", "🔄 Base de Correias", "🛡️ Histórico de Backups"])

    # Aba Fusos
    with tab_fusos_db:
        c1, c2, c3 = st.columns([2, 2, 2])
        c1.metric("Total de Linhas (Fusos)", len(df_fusos))
        c2.metric("Total de Quebras Acumuladas", int(df_fusos["Quantidade_Quebras"].sum()) if not df_fusos.empty else 0)
        c3.metric("Máquinas Cadastradas", df_fusos["Maquina_TAG"].nunique() if not df_fusos.empty else 0)

        filtro_setor_fuso_db = st.selectbox("Filtrar visualização por Setor:", ["Todos"] + list(DICIONARIO_SETORES.keys()), key="f_setor_db_f")
        df_exibir_f = df_fusos if filtro_setor_fuso_db == "Todos" else df_fusos[df_fusos["Setor"] == filtro_setor_fuso_db]

        st.dataframe(df_exibir_f, use_container_width=True, height=380)

        col_d_f, col_u_f = st.columns(2)
        with col_d_f:
            if os.path.exists(ARQUIVO_FUSOS):
                with open(ARQUIVO_FUSOS, "rb") as f_down_fus:
                    st.download_button(
                        "📥 Baixar Base de Fusos (.xlsx)",
                        data=f_down_fus,
                        file_name=f"backup_fusos_{date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
        with col_u_f:
            up_fuso = st.file_uploader("Restaurar Base de Fusos (.xlsx)", type=["xlsx"], key="up_fuso_db")
            if up_fuso is not None:
                if st.button("Substituir Base de Fusos por este Arquivo", type="primary", key="btn_subst_fusos"):
                    gerar_backup_seguro(ARQUIVO_FUSOS)
                    with open(ARQUIVO_FUSOS, "wb") as f_out_f:
                        f_out_f.write(up_fuso.getbuffer())
                    st.success("✅ Base de Fusos substituída com sucesso! Backup anterior gerado.")
                    st.rerun()

    # Aba Correias
    with tab_correias_db:
        c1_c, c2_c, c3_c = st.columns([2, 2, 2])
        c1_c.metric("Total de Linhas (Correias)", len(df_correias))
        c2_c.metric("Máquinas Vinculadas", df_correias["Maquina_TAG"].nunique() if not df_correias.empty else 0)
        c3_c.metric("Setores com Dados", df_correias["Setor"].nunique() if not df_correias.empty else 0)

        filtro_setor_cor_db = st.selectbox("Filtrar visualização por Setor:", ["Todos"] + list(DICIONARIO_SETORES.keys()), key="f_setor_db_c")
        df_exibir_c = df_correias if filtro_setor_cor_db == "Todos" else df_correias[df_correias["Setor"] == filtro_setor_cor_db]

        st.dataframe(df_exibir_c, use_container_width=True, height=380)

        col_d_c, col_u_c = st.columns(2)
        with col_d_c:
            if os.path.exists(ARQUIVO_CORREIAS):
                with open(ARQUIVO_CORREIAS, "rb") as f_down_cor:
                    st.download_button(
                        "📥 Baixar Base de Correias (.xlsx)",
                        data=f_down_cor,
                        file_name=f"backup_correias_{date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
        with col_u_c:
            up_correia = st.file_uploader("Restaurar Base de Correias (.xlsx)", type=["xlsx"], key="up_cor_db")
            if up_correia is not None:
                if st.button("Substituir Base de Correias por este Arquivo", type="primary", key="btn_subst_cor"):
                    gerar_backup_seguro(ARQUIVO_CORREIAS)
                    with open(ARQUIVO_CORREIAS, "wb") as f_out_c:
                        f_out_c.write(up_correia.getbuffer())
                    st.success("✅ Base de Correias substituída com sucesso! Backup anterior gerado.")
                    st.rerun()

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

elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Máquinas")
