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

# ==========================================
# LEITURA E RECUPERAÇÃO AUTOMÁTICA DAS BASES (BLINDAGEM CONTRA BADZIPFILE)
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
# DADOS HISTÓRICOS: SETOR B (JAN A JUL/2026)
# ==========================================
dados_janeiro_setor_b = [
    ("L-29", 1), ("L-30", 13), ("L-31", 10), ("L-32", 7),
    ("L-33", 5), ("L-34", 2), ("L-35", 2), ("L-36", 8),
    ("L-37", 2), ("L-38", 1), ("L-39", 1), ("L-40", 4),
    ("L-50", 7), ("L-51", 0), ("L-41", 1), ("L-42", 2),
    ("L-43", 0), ("L-44", 3), ("L-45", 0), ("L-46", 2),
    ("L-52", 3), ("L-53", 6)
]

dados_fevereiro_setor_b = [
    ("L-29", 4), ("L-30", 23), ("L-31", 9), ("L-32", 34),
    ("L-33", 22), ("L-34", 19), ("L-35", 42), ("L-36", 22),
    ("L-37", 7), ("L-38", 4), ("L-39", 5), ("L-40", 11),
    ("L-50", 45), ("L-51", 2), ("L-41", 1), ("L-42", 2),
    ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 1),
    ("L-52", 2), ("L-53", 5)
]

dados_marco_setor_b = [
    ("L-29", 1), ("L-30", 15), ("L-31", 4), ("L-32", 8),
    ("L-33", 14), ("L-34", 6), ("L-35", 8), ("L-36", 9),
    ("L-37", 18), ("L-38", 2), ("L-39", 5), ("L-40", 6),
    ("L-50", 0), ("L-51", 4), ("L-41", 1), ("L-42", 1),
    ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 0),
    ("L-52", 3), ("L-53", 6)
]

dados_abril_setor_b = [
    ("L-29", 2), ("L-30", 3), ("L-31", 1), ("L-32", 2),
    ("L-33", 17), ("L-34", 8), ("L-35", 7), ("L-36", 15),
    ("L-37", 17), ("L-38", 5), ("L-39", 12), ("L-40", 3),
    ("L-50", 2), ("L-51", 3), ("L-41", 3), ("L-42", 4),
    ("L-43", 4), ("L-44", 2), ("L-45", 0), ("L-46", 1),
    ("L-52", 1), ("L-53", 4)
]

dados_maio_setor_b = [
    ("L-29", 0), ("L-30", 5), ("L-31", 8), ("L-32", 0),
    ("L-33", 6), ("L-34", 2), ("L-35", 7), ("L-36", 10),
    ("L-37", 12), ("L-38", 2), ("L-39", 5), ("L-40", 3),
    ("L-50", 1), ("L-51", 0), ("L-41", 2), ("L-42", 1),
    ("L-43", 8), ("L-44", 1), ("L-45", 0), ("L-46", 2),
    ("L-52", 5), ("L-53", 5)
]

dados_junho_setor_b = [
    ("L-29", 6), ("L-30", 16), ("L-31", 35), ("L-32", 4),
    ("L-33", 10), ("L-34", 4), ("L-35", 8), ("L-36", 11),
    ("L-37", 6), ("L-38", 1), ("L-39", 9), ("L-40", 11),
    ("L-50", 1), ("L-51", 1), ("L-41", 6), ("L-42", 1),
    ("L-43", 2), ("L-44", 0), ("L-45", 0), ("L-46", 0),
    ("L-52", 1), ("L-53", 2)
]

dados_julho_setor_b = [
    ("L-29", 0), ("L-30", 31), ("L-31", 4), ("L-32", 13),
    ("L-33", 26), ("L-34", 7), ("L-35", 14), ("L-36", 2),
    ("L-37", 14), ("L-38", 0), ("L-39", 2), ("L-40", 11),
    ("L-50", 2), ("L-51", 6), ("L-41", 3), ("L-42", 2),
    ("L-43", 3), ("L-44", 5), ("L-45", 0), ("L-46", 0),
    ("L-52", 6), ("L-53", 4)
]

mapa_cargas_setor_b = [
    ("Janeiro", dados_janeiro_setor_b),
    ("Fevereiro", dados_fevereiro_setor_b),
    ("Março", dados_marco_setor_b),
    ("Abril", dados_abril_setor_b),
    ("Maio", dados_maio_setor_b),
    ("Junho", dados_junho_setor_b),
    ("Julho", dados_julho_setor_b),
]

# ==========================================
# DADOS HISTÓRICOS: SETOR A (JAN A SET/2026)
# ==========================================
dados_janeiro_setor_a = [
    ("L-01", 8), ("L-02", 9), ("L-03", 2), ("L-04", 3), ("L-05", 3),
    ("L-06", 13), ("L-07", 6), ("L-08", 4), ("L-09", 8), ("L-10", 9),
    ("L-11", 3), ("L-12", 6), ("L-13", 5), ("L-14", 8), ("L-15", 9),
    ("L-16", 4), ("L-17", 2), ("L-18", 3), ("L-19", 10), ("L-20", 5),
    ("L-21", 4), ("L-22", 4), ("L-23", 6), ("L-24", 10), ("L-25", 4),
    ("L-26", 3), ("L-27", 4), ("L-28", 3)
]

dados_fevereiro_setor_a = [
    ("L-01", 3), ("L-02", 6), ("L-03", 4), ("L-04", 4), ("L-05", 2),
    ("L-06", 5), ("L-07", 7), ("L-08", 3), ("L-09", 4), ("L-10", 2),
    ("L-11", 4), ("L-12", 14), ("L-13", 14), ("L-14", 13), ("L-15", 4),
    ("L-16", 10), ("L-17", 1), ("L-18", 4), ("L-19", 0), ("L-20", 6),
    ("L-21", 7), ("L-22", 4), ("L-23", 3), ("L-24", 4), ("L-25", 4),
    ("L-26", 0), ("L-27", 0), ("L-28", 8)
]

dados_marco_setor_a = [
    ("L-01", 4), ("L-02", 7), ("L-03", 6), ("L-04", 3), ("L-05", 1),
    ("L-06", 5), ("L-07", 10), ("L-08", 2), ("L-09", 12), ("L-10", 8),
    ("L-11", 10), ("L-12", 9), ("L-13", 9), ("L-14", 9), ("L-15", 12),
    ("L-16", 3), ("L-17", 8), ("L-18", 20), ("L-19", 10), ("L-20", 3),
    ("L-21", 4), ("L-22", 5), ("L-23", 3), ("L-24", 8), ("L-25", 4),
    ("L-26", 9), ("L-27", 2), ("L-28", 2)
]

dados_abril_setor_a = [
    ("L-01", 1), ("L-02", 8), ("L-03", 3), ("L-04", 6), ("L-05", 2),
    ("L-06", 7), ("L-07", 8), ("L-08", 0), ("L-09", 2), ("L-10", 6),
    ("L-11", 9), ("L-12", 6), ("L-13", 5), ("L-14", 2), ("L-15", 2),
    ("L-16", 7), ("L-17", 5), ("L-18", 8), ("L-19", 2), ("L-20", 5),
    ("L-21", 3), ("L-22", 12), ("L-23", 5), ("L-24", 9), ("L-25", 2),
    ("L-26", 7), ("L-27", 5), ("L-28", 2)
]

dados_maio_setor_a = [
    ("L-01", 8), ("L-02", 5), ("L-03", 11), ("L-04", 1), ("L-05", 5),
    ("L-06", 7), ("L-07", 6), ("L-08", 2), ("L-09", 8), ("L-10", 1),
    ("L-11", 4), ("L-12", 1), ("L-13", 2), ("L-14", 6), ("L-15", 6),
    ("L-16", 10), ("L-17", 4), ("L-18", 5), ("L-19", 7), ("L-20", 13),
    ("L-21", 3), ("L-22", 4), ("L-23", 9), ("L-24", 7), ("L-25", 1),
    ("L-26", 2), ("L-27", 4), ("L-28", 4)
]

dados_junho_setor_a = [
    ("L-01", 5), ("L-02", 2), ("L-03", 4), ("L-04", 4), ("L-05", 5),
    ("L-06", 3), ("L-07", 8), ("L-08", 1), ("L-09", 6), ("L-10", 5),
    ("L-11", 9), ("L-12", 5), ("L-13", 8), ("L-14", 3), ("L-15", 2),
    ("L-16", 5), ("L-17", 2), ("L-18", 4), ("L-19", 3), ("L-20", 5),
    ("L-21", 5), ("L-22", 3), ("L-23", 4), ("L-24", 4), ("L-25", 3),
    ("L-26", 9), ("L-27", 3), ("L-28", 1)
]

dados_julho_setor_a = [
    ("L-01", 5), ("L-02", 3), ("L-03", 3), ("L-04", 10), ("L-05", 2),
    ("L-06", 6), ("L-07", 11), ("L-08", 8), ("L-09", 11), ("L-10", 7),
    ("L-11", 5), ("L-12", 3), ("L-13", 2), ("L-14", 7), ("L-15", 6),
    ("L-16", 9), ("L-17", 4), ("L-18", 8), ("L-19", 8), ("L-20", 5),
    ("L-21", 11), ("L-22", 18), ("L-23", 9), ("L-24", 10), ("L-25", 7),
    ("L-26", 9), ("L-27", 3), ("L-28", 9)
]

dados_agosto_setor_a = [
    ("L-01", 5), ("L-02", 8), ("L-03", 3), ("L-04", 5), ("L-05", 11),
    ("L-06", 8), ("L-07", 15), ("L-08", 5), ("L-09", 14), ("L-10", 6),
    ("L-11", 9), ("L-12", 5), ("L-13", 7), ("L-14", 9), ("L-15", 11),
    ("L-16", 5), ("L-17", 6), ("L-18", 12), ("L-19", 10), ("L-20", 10),
    ("L-21", 9), ("L-22", 16), ("L-23", 11), ("L-24", 8), ("L-25", 4),
    ("L-26", 2), ("L-27", 5), ("L-28", 10)
]

dados_setembro_setor_a = [
    ("L-01", 1), ("L-02", 2), ("L-03", 2), ("L-04", 4), ("L-05", 1),
    ("L-06", 0), ("L-07", 1), ("L-08", 0), ("L-09", 3), ("L-10", 4),
    ("L-11", 7), ("L-12", 2), ("L-13", 4), ("L-14", 3), ("L-15", 1),
    ("L-16", 0), ("L-17", 1), ("L-18", 1), ("L-19", 4), ("L-20", 4),
    ("L-21", 2), ("L-22", 3), ("L-23", 1), ("L-24", 5), ("L-25", 1),
    ("L-26", 0), ("L-27", 2), ("L-28", 4)
]

mapa_cargas_setor_a = [
    ("Janeiro", dados_janeiro_setor_a),
    ("Fevereiro", dados_fevereiro_setor_a),
    ("Março", dados_marco_setor_a),
    ("Abril", dados_abril_setor_a),
    ("Maio", dados_maio_setor_a),
    ("Junho", dados_junho_setor_a),
    ("Julho", dados_julho_setor_a),
    ("Agosto", dados_agosto_setor_a),
    ("Setembro", dados_setembro_setor_a),
]

# =========================================================================
# DADOS HISTÓRICOS: SETOR LÁTEX (JAN A AGO/2026) - 20 MÁQUINAS OFICIAIS
# =========================================================================
dados_janeiro_setor_latex = [
    ("B-71", 7), ("B-72", 10), ("B-73", 13), ("B-74", 18), ("B-75", 4),
    ("B-76", 9), ("B-77", 10), ("B-78", 5), ("B-79", 1), ("B-80", 0),
    ("B-83", 6), ("B-84", 7), ("B-85", 9), ("B-86", 5), ("B-87", 5),
    ("B-88", 10), ("B-89", 1), ("B-102", 0), ("B-103", 0), ("B-104", 0)
]

dados_fevereiro_setor_latex = [
    ("B-71", 1), ("B-72", 2), ("B-73", 5), ("B-74", 0), ("B-75", 0),
    ("B-76", 1), ("B-77", 2), ("B-78", 3), ("B-79", 4), ("B-80", 0),
    ("B-83", 1), ("B-84", 2), ("B-85", 4), ("B-86", 4), ("B-87", 2),
    ("B-88", 0), ("B-89", 0), ("B-102", 0), ("B-103", 0), ("B-104", 0)
]

dados_marco_setor_latex = [
    ("B-71", 5), ("B-72", 7), ("B-73", 2), ("B-74", 1), ("B-75", 0),
    ("B-76", 3), ("B-77", 4), ("B-78", 4), ("B-79", 0), ("B-80", 1),
    ("B-83", 9), ("B-84", 18), ("B-85", 9), ("B-86", 9), ("B-87", 14),
    ("B-88", 0), ("B-89", 8), ("B-102", 0), ("B-103", 0), ("B-104", 0)
]

dados_abril_setor_latex = [
    ("B-71", 5), ("B-72", 9), ("B-73", 3), ("B-74", 5), ("B-75", 2),
    ("B-76", 4), ("B-77", 5), ("B-78", 3), ("B-79", 0), ("B-80", 0),
    ("B-83", 10), ("B-84", 18), ("B-85", 8), ("B-86", 12), ("B-87", 9),
    ("B-88", 0), ("B-89", 10), ("B-102", 0), ("B-103", 0), ("B-104", 0)
]

dados_maio_setor_latex = [
    ("B-71", 1), ("B-72", 4), ("B-73", 7), ("B-74", 5), ("B-75", 2),
    ("B-76", 4), ("B-77", 1), ("B-78", 0), ("B-79", 0), ("B-80", 0),
    ("B-83", 15), ("B-84", 18), ("B-85", 6), ("B-86", 12), ("B-87", 7),
    ("B-88", 0), ("B-89", 17), ("B-102", 0), ("B-103", 0), ("B-104", 1)
]

dados_junho_setor_latex = [
    ("B-71", 1), ("B-72", 1), ("B-73", 7), ("B-74", 0), ("B-75", 7),
    ("B-76", 1), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0),
    ("B-83", 9), ("B-84", 7), ("B-85", 10), ("B-86", 13), ("B-87", 5),
    ("B-88", 1), ("B-89", 8), ("B-102", 0), ("B-103", 0), ("B-104", 0)
]

dados_julho_setor_latex = [
    ("B-71", 3), ("B-72", 1), ("B-73", 3), ("B-74", 0), ("B-75", 0),
    ("B-76", 0), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0),
    ("B-83", 8), ("B-84", 10), ("B-85", 7), ("B-86", 9), ("B-87", 10),
    ("B-88", 0), ("B-89", 7), ("B-102", 0), ("B-103", 0), ("B-104", 0)
]

dados_agosto_setor_latex = [
    ("B-71", 5), ("B-72", 1), ("B-73", 0), ("B-74", 0), ("B-75", 0),
    ("B-76", 0), ("B-77", 0), ("B-78", 0), ("B-79", 0), ("B-80", 0),
    ("B-83", 1), ("B-84", 3), ("B-85", 6), ("B-86", 1), ("B-87", 1),
    ("B-88", 0), ("B-89", 1), ("B-102", 0), ("B-103", 0), ("B-104", 0)
]

mapa_cargas_setor_latex = [
    ("Janeiro", dados_janeiro_setor_latex),
    ("Fevereiro", dados_fevereiro_setor_latex),
    ("Março", dados_marco_setor_latex),
    ("Abril", dados_abril_setor_latex),
    ("Maio", dados_maio_setor_latex),
    ("Junho", dados_junho_setor_latex),
    ("Julho", dados_julho_setor_latex),
    ("Agosto", dados_agosto_setor_latex),
]

# =========================================================================
# DADOS HISTÓRICOS: SETOR MENEGATTO (JAN A AGO/2026)
# =========================================================================
dados_janeiro_setor_menegatto = [
    ("B-47", 4), ("B-48", 3), ("B-49", 5), ("B-81", 3), ("B-82", 3),
    ("B-93", 10), ("B-94", 9), ("B-95", 4), ("B-96", 6), ("B-97", 4),
    ("B-98", 1), ("B-99", 6), ("B-100", 6), ("B-101", 5), ("B-107", 0),
    ("B-108", 0)
]

dados_fevereiro_setor_menegatto = [
    ("B-47", 4), ("B-48", 3), ("B-49", 2), ("B-81", 2), ("B-82", 0),
    ("B-93", 1), ("B-94", 0), ("B-95", 3), ("B-96", 4), ("B-97", 6),
    ("B-98", 1), ("B-99", 6), ("B-100", 16), ("B-101", 1), ("B-107", 0),
    ("B-108", 0)
]

dados_marco_setor_menegatto = [
    ("B-47", 0), ("B-48", 5), ("B-49", 2), ("B-81", 3), ("B-82", 3),
    ("B-93", 3), ("B-94", 7), ("B-95", 12), ("B-96", 2), ("B-97", 5),
    ("B-98", 1), ("B-99", 9), ("B-100", 0), ("B-101", 3), ("B-107", 0),
    ("B-108", 0)
]

dados_abril_setor_menegatto = [
    ("B-47", 9), ("B-48", 8), ("B-49", 3), ("B-81", 8), ("B-82", 5),
    ("B-93", 3), ("B-94", 2), ("B-95", 25), ("B-96", 5), ("B-97", 3),
    ("B-98", 2), ("B-99", 14), ("B-100", 2), ("B-101", 4), ("B-107", 0),
    ("B-108", 0)
]

dados_maio_setor_menegatto = [
    ("B-47", 5), ("B-48", 9), ("B-49", 1), ("B-81", 7), ("B-82", 1),
    ("B-93", 1), ("B-94", 0), ("B-95", 4), ("B-96", 0), ("B-97", 9),
    ("B-98", 0), ("B-99", 1), ("B-100", 4), ("B-101", 3), ("B-107", 0),
    ("B-108", 0)
]

dados_junho_setor_menegatto = [
    ("B-47", 4), ("B-48", 6), ("B-49", 1), ("B-81", 1), ("B-82", 3),
    ("B-93", 5), ("B-94", 7), ("B-95", 15), ("B-96", 5), ("B-97", 9),
    ("B-98", 0), ("B-99", 6), ("B-100", 3), ("B-101", 1), ("B-107", 0),
    ("B-108", 0)
]

dados_julho_setor_menegatto = [
    ("B-47", 11), ("B-48", 5), ("B-49", 7), ("B-81", 6), ("B-82", 5),
    ("B-93", 1), ("B-94", 0), ("B-95", 1), ("B-96", 3), ("B-97", 1),
    ("B-98", 1), ("B-99", 3), ("B-100", 2), ("B-101", 5), ("B-107", 1),
    ("B-108", 0)
]

dados_agosto_setor_menegatto = [
    ("B-47", 7), ("B-48", 1), ("B-49", 4), ("B-81", 0), ("B-82", 3),
    ("B-93", 2), ("B-94", 0), ("B-95", 1), ("B-96", 2), ("B-97", 3),
    ("B-98", 0), ("B-99", 0), ("B-100", 1), ("B-101", 2), ("B-107", 0),
    ("B-108", 0)
]

mapa_cargas_setor_menegatto = [
    ("Janeiro", dados_janeiro_setor_menegatto),
    ("Fevereiro", dados_fevereiro_setor_menegatto),
    ("Março", dados_marco_setor_menegatto),
    ("Abril", dados_abril_setor_menegatto),
    ("Maio", dados_maio_setor_menegatto),
    ("Junho", dados_junho_setor_menegatto),
    ("Julho", dados_julho_setor_menegatto),
    ("Agosto", dados_agosto_setor_menegatto),
]

precisa_salvar_fusos = False

# Injeção Setor B
for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_b:
    linhas_mes = df_fusos[
        (df_fusos["Ano"] == 2026)
        & (df_fusos["Mes"] == nome_mes_carga)
        & (df_fusos["Setor"] == "Setor B")
    ]
    if linhas_mes.empty or linhas_mes["Quantidade_Quebras"].sum() == 0:
        df_fusos = df_fusos[
            ~(
                (df_fusos["Ano"] == 2026)
                & (df_fusos["Mes"] == nome_mes_carga)
                & (df_fusos["Setor"] == "Setor B")
            )
        ]
        novos_reg = [
            {
                "Ano": 2026,
                "Mes": nome_mes_carga,
                "Setor": "Setor B",
                "Maquina_TAG": maq,
                "Quantidade_Quebras": int(qtd),
                "Tipo_Fuso": "TEP",
            }
            for maq, qtd in lista_dados_carga
        ]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

# Injeção Setor A
for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_a:
    linhas_mes = df_fusos[
        (df_fusos["Ano"] == 2026)
        & (df_fusos["Mes"] == nome_mes_carga)
        & (df_fusos["Setor"] == "Setor A")
    ]
    if linhas_mes.empty or linhas_mes["Quantidade_Quebras"].sum() == 0:
        df_fusos = df_fusos[
            ~(
                (df_fusos["Ano"] == 2026)
                & (df_fusos["Mes"] == nome_mes_carga)
                & (df_fusos["Setor"] == "Setor A")
            )
        ]
        novos_reg = [
            {
                "Ano": 2026,
                "Mes": nome_mes_carga,
                "Setor": "Setor A",
                "Maquina_TAG": maq,
                "Quantidade_Quebras": int(qtd),
                "Tipo_Fuso": "FAG",
            }
            for maq, qtd in lista_dados_carga
        ]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

# Injeção Setor Látex
for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_latex:
    linhas_mes = df_fusos[
        (df_fusos["Ano"] == 2026)
        & (df_fusos["Mes"] == nome_mes_carga)
        & (df_fusos["Setor"] == "Setor Látex")
    ]
    if linhas_mes.empty or len(linhas_mes) < 20:
        df_fusos = df_fusos[
            ~(
                (df_fusos["Ano"] == 2026)
                & (df_fusos["Mes"] == nome_mes_carga)
                & (df_fusos["Setor"] == "Setor Látex")
            )
        ]
        novos_reg = [
            {
                "Ano": 2026,
                "Mes": nome_mes_carga,
                "Setor": "Setor Látex",
                "Maquina_TAG": maq,
                "Quantidade_Quebras": int(qtd),
                "Tipo_Fuso": "M4BA",
            }
            for maq, qtd in lista_dados_carga
        ]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

# Injeção Setor Menegatto
for nome_mes_carga, lista_dados_carga in mapa_cargas_setor_menegatto:
    linhas_mes = df_fusos[
        (df_fusos["Ano"] == 2026)
        & (df_fusos["Mes"] == nome_mes_carga)
        & (df_fusos["Setor"] == "Setor Menegatto")
    ]
    if linhas_mes.empty or linhas_mes["Quantidade_Quebras"].sum() == 0:
        df_fusos = df_fusos[
            ~(
                (df_fusos["Ano"] == 2026)
                & (df_fusos["Mes"] == nome_mes_carga)
                & (df_fusos["Setor"] == "Setor Menegatto")
            )
        ]
        novos_reg = [
            {
                "Ano": 2026,
                "Mes": nome_mes_carga,
                "Setor": "Setor Menegatto",
                "Maquina_TAG": maq,
                "Quantidade_Quebras": int(qtd),
                "Tipo_Fuso": "MENEGATTO",
            }
            for maq, qtd in lista_dados_carga
        ]
        df_fusos = pd.concat([df_fusos, pd.DataFrame(novos_reg)], ignore_index=True)
        precisa_salvar_fusos = True

if precisa_salvar_fusos:
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Base de Correias com blindagem contra corrupção
df_correias = None
if os.path.exists(ARQUIVO_CORREIAS):
    try:
        df_correias = pd.read_excel(ARQUIVO_CORREIAS)
    except Exception:
        df_correias = pd.DataFrame(columns=colunas_correias)
        df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
else:
    if os.path.exists("lancamentos_correias_v3.xlsx"):
        try:
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
        except Exception:
            df_correias = pd.DataFrame(columns=colunas_correias)
            df_correias.to_excel(ARQUIVO_CORREIAS, index=False)
    else:
        df_correias = pd.DataFrame(columns=colunas_correias)
        df_correias.to_excel(ARQUIVO_CORREIAS, index=False)

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

# Mapeamento oficial de ativos por setor (atualizado com novas tags do Menegatto)
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
    "B-94", "B-95", "B-96", "B-97", "B-98", "B-99", "B-100", "B-101", "B-107", "B-108"
]

DICIONARIO_SETORES = {
    "Setor A": maquinas_setor_a,
    "Setor B": maquinas_setor_b,
    "Setor Látex": maquinas_setor_latex,
    "Setor Menegatto": maquinas_setor_menegatto,
}

# Função auxiliar para obter todas as máquinas de um setor (base estática + adicionadas)
def obter_maquinas_setor(setor_nome, df_ref_correias=None, df_ref_fusos=None):
    lista_base = list(DICIONARIO_SETORES.get(setor_nome, []))
    extras = set()
    if df_ref_correias is not None and not df_ref_correias.empty:
        maq_c = df_ref_correias[df_ref_correias["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique()
        extras.update(maq_c)
    if df_ref_fusos is not None and not df_ref_fusos.empty:
        maq_f = df_ref_fusos[df_ref_fusos["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique()
        extras.update(maq_f)
    todas = set(lista_base).union(extras)
    return sorted(list(todas))

mapa_setor_maquina = {}
for setor_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m:
        mapa_setor_maquina[m] = setor_nome

# Atualiza mapa de setor com registos do excel caso haja novas máquinas
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
# CÁLCULO E ANÁLISE DE CORREIAS (BLINDADO)
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
    setor_m = mapa_setor_maquina.get(maq_tag, "Setor A")
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

        /* Overlay Invisível */
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

        .chart-header-row {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-bottom: 6px;
            margin-bottom: 6px;
            border-bottom: 1px solid #f1f5f9;
        }}
        .chart-header-title {{
            font-size: 0.96rem;
            font-weight: 800;
            color: #0f172a;
        }}
        .chart-header-badge {{
            font-size: 0.82rem;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 6px;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# BARRA LATERAL (SIDEBAR) CUSTOMIZADA
# ==========================================
with st.sidebar:
    st.markdown(
        """
        <div style="margin-top: -1.5rem; margin-bottom: 0.5rem;">
            <h1 style="font-size: 1.8rem; margin-bottom: 0px; color: #0f172a;">⚙️ Portal PCM</h1>
            <p style="font-size: 0.9rem; font-weight: 700; color: #64748b; margin-top: 2px;">Controle MEC</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
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

MAPA_MES_ABREV = {
    "Janeiro": "JAN", "Fevereiro": "FEV", "Março": "MAR", "Abril": "ABR",
    "Maio": "MAI", "Junho": "JUN", "Julho": "JUL", "Agosto": "AGO",
    "Setembro": "SET", "Outubro": "OUT", "Novembro": "NOV", "Dezembro": "DEZ"
}
ORDEM_MESES_ABREV = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

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
    for setor_nome in DICIONARIO_SETORES.keys():
        if filtro_setor != "Todos os Setores" and setor_nome != filtro_setor:
            continue
        for m in obter_maquinas_setor(setor_nome, df_correias, df_fusos):
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
        col_setor, col_add, col_del = st.columns([1.6, 2.2, 2.2])
        with col_setor:
            setor_selecionado = st.selectbox(
                "🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_setor_correias_direto"
            )

        df_atual_cor = pd.read_excel(ARQUIVO_CORREIAS)
        maquinas_do_setor = obter_maquinas_setor(setor_selecionado, df_atual_cor, df_fusos)

        # Inserção de Nova Máquina
        with col_add:
            st.markdown("<div style='font-size:0.8rem; font-weight:700; color:#334155; margin-bottom:2px;'>➕ Adicionar Máquina</div>", unsafe_allow_html=True)
            c_input_add, c_btn_add = st.columns([1.5, 1.0])
            with c_input_add:
                nova_maq_cor = st.text_input("Nova Máquina", placeholder="Ex: L-99 ou B-105", key=f"inp_add_cor_{setor_selecionado}", label_visibility="collapsed").strip().upper()
            with c_btn_add:
                if st.button("Adicionar", key=f"btn_add_cor_{setor_selecionado}", use_container_width=True):
                    if nova_maq_cor:
                        if nova_maq_cor in maquinas_do_setor:
                            st.warning(f"A máquina {nova_maq_cor} já existe no {setor_selecionado}.")
                        else:
                            novo_reg_cor = pd.DataFrame([{
                                "Setor": setor_selecionado,
                                "Maquina_TAG": nova_maq_cor,
                                "Tipo_Correia_1": "",
                                "Data_Instalacao_1": "",
                                "Tipo_Correia_2": "",
                                "Data_Instalacao_2": ""
                            }])
                            df_atual_cor = pd.concat([df_atual_cor, novo_reg_cor], ignore_index=True)
                            df_atual_cor.to_excel(ARQUIVO_CORREIAS, index=False)
                            st.success(f"Máquina {nova_maq_cor} adicionada ao {setor_selecionado}!")
                            st.rerun()
                    else:
                        st.warning("Informe o código da máquina.")

        # Exclusão de Máquina
        with col_del:
            st.markdown("<div style='font-size:0.8rem; font-weight:700; color:#334155; margin-bottom:2px;'>🗑️ Excluir Máquina</div>", unsafe_allow_html=True)
            c_sel_del, c_btn_del = st.columns([1.5, 1.0])
            with c_sel_del:
                maq_del_cor = st.selectbox("Excluir", ["-- Selecione --"] + maquinas_do_setor, key=f"sel_del_cor_{setor_selecionado}", label_visibility="collapsed")
            with c_btn_del:
                if st.button("Excluir", key=f"btn_del_cor_{setor_selecionado}", use_container_width=True, type="secondary"):
                    if maq_del_cor and maq_del_cor != "-- Selecione --":
                        df_atual_cor = df_atual_cor[~((df_atual_cor["Setor"] == setor_selecionado) & (df_atual_cor["Maquina_TAG"] == maq_del_cor))]
                        df_atual_cor.to_excel(ARQUIVO_CORREIAS, index=False)
                        st.success(f"Máquina {maq_del_cor} removida do {setor_selecionado}!")
                        st.rerun()
                    else:
                        st.warning("Selecione uma máquina para excluir.")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(f"Apontamento de Correias — {setor_selecionado}")

    df_filtrado_cor = df_atual_cor[df_atual_cor["Setor"] == setor_selecionado]
    maquinas_do_setor = obter_maquinas_setor(setor_selecionado, df_atual_cor, df_fusos)

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
# 3. PAINEL GERENCIAL DE FUSOS
# ------------------------------------------
elif tela == "Painel Fusos":
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

    # ==========================================
    # CÁLCULO DINÂMICO DE MESES TRANSCORRIDOS
    # ==========================================
    data_hoje_ref = date.today()
    ano_atual_ref = data_hoje_ref.year
    mes_atual_num_ref = data_hoje_ref.month

    if int(ano_painel) < ano_atual_ref:
        meses_divisor = 12
        desc_meses_divisor = "12 meses"
    elif int(ano_painel) == ano_atual_ref:
        meses_divisor = max(1, mes_atual_num_ref)
        nome_mes_vigente_abrev = ORDEM_MESES_ABREV[meses_divisor - 1]
        desc_meses_divisor = f"Jan a {nome_mes_vigente_abrev}"
    else:
        meses_divisor = 1
        desc_meses_divisor = "Previsto"

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

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

    df_dados_fusos = pd.read_excel(ARQUIVO_FUSOS)
    df_fuso_ano = df_dados_fusos[df_dados_fusos["Ano"] == int(ano_painel)].copy()

    # ==========================================
    # CASO 1: ABA GERAL (FÁBRICA COMPLETA)
    # ==========================================
    if st.session_state.aba_setor_fuso == "Geral":
        total_geral_quebras = int(df_fuso_ano["Quantidade_Quebras"].sum()) if not df_fuso_ano.empty else 0
        media_mensal_fabrica = round(total_geral_quebras / meses_divisor, 1)

        # Identificar o último mês que possui apontamentos na fábrica
        ultimo_mes_fabrica = "Nenhum"
        total_quebras_mes_atual = 0
        setor_ofensor_mes = "Nenhum"
        qtd_setor_ofensor_mes = 0

        if not df_fuso_ano.empty and total_geral_quebras > 0:
            df_reais_fabrica = df_fuso_ano[df_fuso_ano["Quantidade_Quebras"] > 0]
            for m_teste in reversed(lista_meses_puros):
                df_sub_fab = df_reais_fabrica[df_reais_fabrica["Mes"] == m_teste]
                if not df_sub_fab.empty and df_sub_fab["Quantidade_Quebras"].sum() > 0:
                    ultimo_mes_fabrica = m_teste
                    total_quebras_mes_atual = int(df_sub_fab["Quantidade_Quebras"].sum())
                    agrup_setor_mes = (
                        df_sub_fab.groupby("Setor")["Quantidade_Quebras"]
                        .sum()
                        .sort_values(ascending=False)
                    )
                    setor_ofensor_mes = agrup_setor_mes.index[0]
                    qtd_setor_ofensor_mes = int(agrup_setor_mes.iloc[0])
                    break

        lbl_setor_critico = f"Setor Crítico ({ultimo_mes_fabrica})" if ultimo_mes_fabrica != "Nenhum" else "Setor Crítico"
        lbl_quebras_mes_atual = f"Quebras ({ultimo_mes_fabrica})" if ultimo_mes_fabrica != "Nenhum" else "Quebras no Mês"

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
                        <div class="kpi-lbl">Média Mensal ({desc_meses_divisor})</div>
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
                        <div class="kpi-lbl">{lbl_setor_critico}</div>
                        <div class="kpi-val" style="color:#d97706; font-size:1.1rem;">{setor_ofensor_mes} ({qtd_setor_ofensor_mes})</div>
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
                        <div class="kpi-lbl">{lbl_quebras_mes_atual}</div>
                        <div class="kpi-val" style="color:#dc2626; font-size:1.35rem;">{total_quebras_mes_atual}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">🚨</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        def gerar_grafico_setor(nome_setor, cor_primaria):
            df_s = df_fuso_ano[df_fuso_ano["Setor"] == nome_setor]
            agrup_s = df_s.groupby("Mes")["Quantidade_Quebras"].sum().reset_index()
            df_base_meses = pd.DataFrame({"Mes": lista_meses_puros})
            df_consolidado = pd.merge(df_base_meses, agrup_s, on="Mes", how="left").fillna(0)
            df_consolidado["Quantidade_Quebras"] = df_consolidado["Quantidade_Quebras"].astype(int)
            df_consolidado["Mes_Abrev"] = df_consolidado["Mes"].map(MAPA_MES_ABREV)
            total_setor = df_consolidado["Quantidade_Quebras"].sum()

            barras = (
                alt.Chart(df_consolidado)
                .mark_bar(color=cor_primaria, cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                    tooltip=[
                        alt.Tooltip("Mes:N", title="Mês"),
                        alt.Tooltip("Quantidade_Quebras:Q", title="Quebras"),
                    ],
                )
            )

            rotulos = (
                alt.Chart(df_consolidado)
                .mark_text(dy=-6, fontSize=11, fontWeight=700, color="#334155")
                .encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV),
                    y=alt.Y("Quantidade_Quebras:Q"),
                    text=alt.condition(
                        alt.datum.Quantidade_Quebras > 0,
                        alt.Text("Quantidade_Quebras:Q"),
                        alt.value("")
                    ),
                )
            )

            chart = (barras + rotulos).properties(height=210)
            return chart, total_setor

        col_g1, col_g2 = st.columns(2)

        with col_g1:
            with st.container(border=True):
                chart_a, tot_a = gerar_grafico_setor("Setor A", "#2563eb")
                st.markdown(
                    f"""
                    <div class="chart-header-row">
                        <span class="chart-header-title">🏭 Setor A</span>
                        <span class="chart-header-badge" style="color:#2563eb;">Total: <b>{tot_a} fusos</b></span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.altair_chart(chart_a, use_container_width=True)

            with st.container(border=True):
                chart_latex, tot_latex = gerar_grafico_setor("Setor Látex", "#059669")
                st.markdown(
                    f"""
                    <div class="chart-header-row">
                        <span class="chart-header-title">🌿 Setor Látex</span>
                        <span class="chart-header-badge" style="color:#059669;">Total: <b>{tot_latex} fusos</b></span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.altair_chart(chart_latex, use_container_width=True)

        with col_g2:
            with st.container(border=True):
                chart_b, tot_b = gerar_grafico_setor("Setor B", "#d97706")
                st.markdown(
                    f"""
                    <div class="chart-header-row">
                        <span class="chart-header-title">🏭 Setor B</span>
                        <span class="chart-header-badge" style="color:#d97706;">Total: <b>{tot_b} fusos</b></span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.altair_chart(chart_b, use_container_width=True)

            with st.container(border=True):
                chart_men, tot_men = gerar_grafico_setor("Setor Menegatto", "#dc2626")
                st.markdown(
                    f"""
                    <div class="chart-header-row">
                        <span class="chart-header-title">⚙️ Setor Menegatto</span>
                        <span class="chart-header-badge" style="color:#dc2626;">Total: <b>{tot_men} fusos</b></span>
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
        
        maquinas_setor_lista = obter_maquinas_setor(setor_ativo, df_correias, df_fusos)
        qtd_maquinas_setor = len(maquinas_setor_lista)
        media_mensal_setor = round(total_setor_quebras / meses_divisor, 1)

        # Identificação do último mês com apontamentos no setor ativo
        ultimo_mes_nome = "Nenhum"
        top_maq_ultimo_mes = "Nenhuma"
        qtd_top_ultimo_mes = 0
        quebras_ultimo_mes_setor = 0

        if not df_setor.empty and total_setor_quebras > 0:
            df_reais_setor = df_setor[df_setor["Quantidade_Quebras"] > 0]
            for m_teste in reversed(lista_meses_puros):
                df_sub_m = df_reais_setor[df_reais_setor["Mes"] == m_teste]
                if not df_sub_m.empty and df_sub_m["Quantidade_Quebras"].sum() > 0:
                    ultimo_mes_nome = m_teste
                    quebras_ultimo_mes_setor = int(df_sub_m["Quantidade_Quebras"].sum())
                    agrup_ult_m = (
                        df_sub_m.groupby("Maquina_TAG")["Quantidade_Quebras"]
                        .sum()
                        .reset_index()
                        .sort_values(by="Quantidade_Quebras", ascending=False)
                    )
                    top_maq_ultimo_mes = agrup_ult_m.iloc[0]["Maquina_TAG"]
                    qtd_top_ultimo_mes = int(agrup_ult_m.iloc[0]["Quantidade_Quebras"])
                    break

        if qtd_maquinas_setor > 0 and quebras_ultimo_mes_setor > 0:
            quebras_por_maquina_calc = round(quebras_ultimo_mes_setor / qtd_maquinas_setor, 1)
        else:
            quebras_por_maquina_calc = 0.0

        lbl_quebras_maq = f"Quebras / Máquina ({ultimo_mes_nome})" if ultimo_mes_nome != "Nenhum" else "Quebras / Máquina"
        lbl_maior_quebra = f"Maior Quebra ({ultimo_mes_nome})" if ultimo_mes_nome != "Nenhum" else "Maior Quebra"

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
                        <div class="kpi-lbl">Média Mensal ({desc_meses_divisor})</div>
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
                        <div class="kpi-lbl">{lbl_quebras_maq}</div>
                        <div class="kpi-val" style="color:#d97706;">{quebras_por_maquina_calc}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">⚙️</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with ks4:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-crit">
                    <div>
                        <div class="kpi-lbl">{lbl_maior_quebra}</div>
                        <div class="kpi-val" style="color:#dc2626; font-size:1.1rem;">{top_maq_ultimo_mes} ({qtd_top_ultimo_mes})</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">⚠️</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # Gráfico de Colunas Mensal do Setor + Gráfico de Rosca por Tipo de Fuso
        c_linha_s, c_tipo_s = st.columns([1.55, 1.45])

        with c_linha_s:
            with st.container(border=True):
                st.markdown(
                    f"""
                    <div class="chart-header-row">
                        <span class="chart-header-title">📊 Evolução Mensal de Quebras ({setor_ativo} - {ano_painel})</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                df_base_meses = pd.DataFrame({"Mes": lista_meses_puros})
                agrup_mes_setor = df_setor.groupby("Mes")["Quantidade_Quebras"].sum().reset_index()
                df_evol_setor = pd.merge(df_base_meses, agrup_mes_setor, on="Mes", how="left").fillna(0)
                df_evol_setor["Quantidade_Quebras"] = df_evol_setor["Quantidade_Quebras"].astype(int)
                df_evol_setor["Mes_Abrev"] = df_evol_setor["Mes"].map(MAPA_MES_ABREV)

                barras_setor = (
                    alt.Chart(df_evol_setor)
                    .mark_bar(color="#2563eb", cornerRadiusTopLeft=5, cornerRadiusTopRight=5)
                    .encode(
                        x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title="Mês", axis=alt.Axis(labelAngle=0, labelFontWeight="bold")),
                        y=alt.Y("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                        tooltip=[
                            alt.Tooltip("Mes:N", title="Mês"),
                            alt.Tooltip("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                        ],
                    )
                )

                rotulos_setor = (
                    alt.Chart(df_evol_setor)
                    .mark_text(dy=-8, fontSize=11, fontWeight=700, color="#1e293b")
                    .encode(
                        x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV),
                        y=alt.Y("Quantidade_Quebras:Q"),
                        text=alt.condition(
                            alt.datum.Quantidade_Quebras > 0,
                            alt.Text("Quantidade_Quebras:Q"),
                            alt.value("")
                        ),
                    )
                )

                chart_colunas_setor = (barras_setor + rotulos_setor).properties(height=280)
                st.altair_chart(chart_colunas_setor, use_container_width=True)

        with c_tipo_s:
            with st.container(border=True):
                st.markdown(
                    f"""
                    <div class="chart-header-row">
                        <span class="chart-header-title">🍩 Distribuição por Tipo de Fuso</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if not df_setor.empty and total_setor_quebras > 0:
                    df_tipo_agrup = (
                        df_setor[df_setor["Quantidade_Quebras"] > 0]
                        .groupby("Tipo_Fuso")["Quantidade_Quebras"]
                        .sum()
                        .reset_index()
                    )
                    
                    chart_rosca = (
                        alt.Chart(df_tipo_agrup)
                        .mark_arc(innerRadius=60, outerRadius=110, stroke="#ffffff", strokeWidth=2)
                        .encode(
                            theta=alt.Theta("Quantidade_Quebras:Q", stack=True),
                            color=alt.Color(
                                "Tipo_Fuso:N",
                                title="Marca / Tipo",
                                scale=alt.Scale(scheme="category10"),
                                legend=alt.Legend(orient="right", labelFontSize=11, titleFontSize=12)
                            ),
                            tooltip=[
                                alt.Tooltip("Tipo_Fuso:N", title="Tipo de Fuso"),
                                alt.Tooltip("Quantidade_Quebras:Q", title="Total de Quebras"),
                            ],
                        )
                        .properties(height=280)
                    )
                    st.altair_chart(chart_rosca, use_container_width=True)
                else:
                    st.info(f"Sem registos de tipos de fuso para {setor_ativo} em {ano_painel}.")

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        # =========================================================
        # 1. MAPA DE CALOR: QUEBRAS POR MÁQUINA X MÊS (GERAL DO SETOR)
        # =========================================================
        with st.container(border=True):
            st.markdown(
                f"""
                <div class="chart-header-row">
                    <span class="chart-header-title">🔥 Mapa de Calor Operacional — Quebras por Máquina x Mês ({setor_ativo} - {ano_painel})</span>
                    <span class="chart-header-badge" style="color:#d97706;">Eixo Cronológico: Jan a Dez</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            linhas_grade_calor = []
            for maq in maquinas_setor_lista:
                for mes_completo in lista_meses_puros:
                    mes_abrev = MAPA_MES_ABREV[mes_completo]
                    linhas_grade_calor.append({
                        "MAQ": maq,
                        "Mes_Completo": mes_completo,
                        "MES": mes_abrev,
                        "Quantidade_Quebras": 0
                    })
            
            df_calor_base = pd.DataFrame(linhas_grade_calor)

            if not df_setor.empty:
                df_setor_agrup = df_setor.groupby(["Maquina_TAG", "Mes"])["Quantidade_Quebras"].sum().reset_index()
                df_setor_agrup.rename(columns={"Maquina_TAG": "MAQ", "Mes": "Mes_Completo", "Quantidade_Quebras": "QTD_REAL"}, inplace=True)
                df_calor_mesclado = pd.merge(df_calor_base, df_setor_agrup, on=["MAQ", "Mes_Completo"], how="left")
                df_calor_mesclado["Quantidade_Quebras"] = df_calor_mesclado["QTD_REAL"].fillna(0).astype(int)
            else:
                df_calor_mesclado = df_calor_base

            altura_calor = max(380, len(maquinas_setor_lista) * 24)

            rect_heatmap = (
                alt.Chart(df_calor_mesclado)
                .mark_rect(stroke="#ffffff", strokeWidth=1)
                .encode(
                    x=alt.X("MES:N", sort=ORDEM_MESES_ABREV, title="Mês", axis=alt.Axis(orient="top", labelAngle=0, labelFontWeight="bold", labelColor="#0f172a")),
                    y=alt.Y("MAQ:N", sort=maquinas_setor_lista, title="Máquina", axis=alt.Axis(labelFontWeight="bold", labelColor="#0f172a")),
                    color=alt.Color(
                        "Quantidade_Quebras:Q",
                        scale=alt.Scale(
                            domain=[0, 3, 8, 15],
                            range=["#dcfce7", "#fef08a", "#f97316", "#dc2626"]
                        ),
                        legend=alt.Legend(title="Escala de Quebras", orient="right")
                    ),
                    tooltip=[
                        alt.Tooltip("MAQ:N", title="Máquina"),
                        alt.Tooltip("MES:N", title="Mês"),
                        alt.Tooltip("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                    ],
                )
            )

            text_heatmap = (
                alt.Chart(df_calor_mesclado)
                .mark_text(baseline="middle", fontSize=11, fontWeight=700)
                .encode(
                    x=alt.X("MES:N", sort=ORDEM_MESES_ABREV),
                    y=alt.Y("MAQ:N", sort=maquinas_setor_lista),
                    text=alt.Text("Quantidade_Quebras:Q"),
                    color=alt.condition(
                        alt.datum.Quantidade_Quebras >= 10,
                        alt.value("#ffffff"),
                        alt.value("#0f172a")
                    ),
                    tooltip=[
                        alt.Tooltip("MAQ:N", title="Máquina"),
                        alt.Tooltip("MES:N", title="Mês"),
                        alt.Tooltip("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                    ],
                )
            )

            chart_calor_final = (rect_heatmap + text_heatmap).properties(
                height=altura_calor
            )

            st.altair_chart(chart_calor_final, use_container_width=True)

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        # =========================================================
        # 2. DIAGNÓSTICO POR TIPO DE FUSO: CARDS + MAPA DE CALOR DO TIPO ESCOLHIDO
        # =========================================================
        tipos_disponiveis_setor = sorted(df_setor["Tipo_Fuso"].dropna().unique().tolist()) if not df_setor.empty else []
        if not tipos_disponiveis_setor:
            tipos_disponiveis_setor = OPCOES_TIPO_FUSO

        with st.container(border=True):
            col_diag_t, col_diag_sel = st.columns([2.5, 1.5])
            with col_diag_t:
                st.markdown(
                    f"""
                    <div style="font-size:1.05rem; font-weight:800; color:#0f172a; display:flex; align-items:center; gap:8px;">
                        <span>🔬 Desempenho Operacional por Tipo de Fuso — {setor_ativo}</span>
                    </div>
                    <div style="font-size:0.8rem; color:#64748b; font-weight:600; margin-top:2px;">
                        Acompanhe o mapa térmico de máquinas atreladas especificamente a cada tipo/marca de fuso ao longo dos 12 meses.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with col_diag_sel:
                fuso_selecionado_analise = st.selectbox(
                    "Tipo de Fuso para Análise Detalhada:",
                    tipos_disponiveis_setor,
                    key=f"sel_tipo_fuso_detalhe_{setor_ativo}",
                )

            st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)

            # Filtra o setor pelo tipo de fuso escolhido
            df_tipo_especifico = df_setor[df_setor["Tipo_Fuso"] == fuso_selecionado_analise].copy()

            tot_fuso_sel = int(df_tipo_especifico["Quantidade_Quebras"].sum()) if not df_tipo_especifico.empty else 0
            media_fuso_sel = round(tot_fuso_sel / meses_divisor, 1)
            
            maqs_vinculadas = sorted(df_tipo_especifico["Maquina_TAG"].unique().tolist())
            qtd_maqs_vinculadas = len(maqs_vinculadas)
            
            df_tipo_falhas = df_tipo_especifico[df_tipo_especifico["Quantidade_Quebras"] > 0]
            maqs_com_quebra_tipo = df_tipo_falhas["Maquina_TAG"].nunique()

            if not df_tipo_falhas.empty:
                agrup_top_fuso = df_tipo_falhas.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().reset_index().sort_values(by="Quantidade_Quebras", ascending=False)
                top_maq_fuso = agrup_top_fuso.iloc[0]["Maquina_TAG"]
                qtd_top_fuso = int(agrup_top_fuso.iloc[0]["Quantidade_Quebras"])
            else:
                top_maq_fuso = "Nenhuma"
                qtd_top_fuso = 0

            c_f1, c_f2, c_f3, c_f4 = st.columns(4)
            with c_f1:
                st.markdown(
                    f"""
                    <div class="card-kpi-bonito c-total" style="height:58px;">
                        <div>
                            <div class="kpi-lbl">Total Quebras ({fuso_selecionado_analise})</div>
                            <div class="kpi-val" style="color:#0f172a; font-size:1.2rem;">{tot_fuso_sel} un.</div>
                        </div>
                        <div style="font-size:1.3rem;">🏷️</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c_f2:
                st.markdown(
                    f"""
                    <div class="card-kpi-bonito c-ok" style="height:58px;">
                        <div>
                            <div class="kpi-lbl">Média Mensal ({desc_meses_divisor})</div>
                            <div class="kpi-val" style="color:#059669; font-size:1.2rem;">{media_fuso_sel} /mês</div>
                        </div>
                        <div style="font-size:1.3rem;">📉</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c_f3:
                st.markdown(
                    f"""
                    <div class="card-kpi-bonito c-warn" style="height:58px;">
                        <div>
                            <div class="kpi-lbl">Máquinas Atreladas</div>
                            <div class="kpi-val" style="color:#d97706; font-size:1.2rem;">{maqs_com_quebra_tipo} falharam / {qtd_maqs_vinculadas}</div>
                        </div>
                        <div style="font-size:1.3rem;">⚙️</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c_f4:
                st.markdown(
                    f"""
                    <div class="card-kpi-bonito c-crit" style="height:58px;">
                        <div>
                            <div class="kpi-lbl">Maior Ofensor ({fuso_selecionado_analise})</div>
                            <div class="kpi-val" style="color:#dc2626; font-size:1.05rem;">{top_maq_fuso} ({qtd_top_fuso})</div>
                        </div>
                        <div style="font-size:1.3rem;">🚨</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

            st.markdown(
                f"""
                <div class="chart-header-row">
                    <span class="chart-header-title">🔥 Mapa Térmico por Máquina — Fuso: {fuso_selecionado_analise} ({setor_ativo} - {ano_painel})</span>
                    <span class="chart-header-badge" style="color:#4f46e5;">Filtrado: {fuso_selecionado_analise}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if not maqs_vinculadas:
                st.info(f"Nenhuma máquina possui registos associados ao fuso {fuso_selecionado_analise} no {setor_ativo} em {ano_painel}.")
            else:
                linhas_calor_tipo = []
                for maq in maqs_vinculadas:
                    for mes_completo in lista_meses_puros:
                        mes_abrev = MAPA_MES_ABREV[mes_completo]
                        linhas_calor_tipo.append({
                            "MAQ": maq,
                            "Mes_Completo": mes_completo,
                            "MES": mes_abrev,
                            "Quantidade_Quebras": 0
                        })

                df_calor_base_tipo = pd.DataFrame(linhas_calor_tipo)

                df_tipo_agrup_calor = (
                    df_tipo_especifico.groupby(["Maquina_TAG", "Mes"])["Quantidade_Quebras"]
                    .sum()
                    .reset_index()
                )
                df_tipo_agrup_calor.rename(
                    columns={"Maquina_TAG": "MAQ", "Mes": "Mes_Completo", "Quantidade_Quebras": "QTD_REAL"},
                    inplace=True
                )

                df_calor_tipo_mesclado = pd.merge(
                    df_calor_base_tipo, df_tipo_agrup_calor, on=["MAQ", "Mes_Completo"], how="left"
                )
                df_calor_tipo_mesclado["Quantidade_Quebras"] = df_calor_tipo_mesclado["QTD_REAL"].fillna(0).astype(int)

                altura_calor_tipo = max(240, len(maqs_vinculadas) * 24)

                rect_heatmap_tipo = (
                    alt.Chart(df_calor_tipo_mesclado)
                    .mark_rect(stroke="#ffffff", strokeWidth=1)
                    .encode(
                        x=alt.X("MES:N", sort=ORDEM_MESES_ABREV, title="Mês", axis=alt.Axis(orient="top", labelAngle=0, labelFontWeight="bold", labelColor="#0f172a")),
                        y=alt.Y("MAQ:N", sort=maqs_vinculadas, title="Máquina", axis=alt.Axis(labelFontWeight="bold", labelColor="#0f172a")),
                        color=alt.Color(
                            "Quantidade_Quebras:Q",
                            scale=alt.Scale(
                                domain=[0, 3, 8, 15],
                                range=["#dcfce7", "#fef08a", "#f97316", "#dc2626"]
                            ),
                            legend=alt.Legend(title="Escala de Quebras", orient="right")
                        ),
                        tooltip=[
                            alt.Tooltip("MAQ:N", title="Máquina"),
                            alt.Tooltip("MES:N", title="Mês"),
                            alt.Tooltip("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                        ],
                    )
                )

                text_heatmap_tipo = (
                    alt.Chart(df_calor_tipo_mesclado)
                    .mark_text(baseline="middle", fontSize=11, fontWeight=700)
                    .encode(
                        x=alt.X("MES:N", sort=ORDEM_MESES_ABREV),
                        y=alt.Y("MAQ:N", sort=maqs_vinculadas),
                        text=alt.Text("Quantidade_Quebras:Q"),
                        color=alt.condition(
                            alt.datum.Quantidade_Quebras >= 10,
                            alt.value("#ffffff"),
                            alt.value("#0f172a")
                        ),
                        tooltip=[
                            alt.Tooltip("MAQ:N", title="Máquina"),
                            alt.Tooltip("MES:N", title="Mês"),
                            alt.Tooltip("Quantidade_Quebras:Q", title="Quebras Apontadas"),
                        ],
                    )
                )

                chart_calor_tipo_final = (rect_heatmap_tipo + text_heatmap_tipo).properties(
                    height=altura_calor_tipo
                )

                st.altair_chart(chart_calor_tipo_final, use_container_width=True)

# ------------------------------------------
# 4. LANÇAMENTOS: FUSOS (COM INSERÇÃO E EXCLUSÃO)
# ------------------------------------------
elif tela == "Lançamento Fusos":
    st.title("🔩 Lançamento: Fechamento Mensal de Fusos")
    st.caption("Preenchimento rápido de quebras por máquina, seleção de tipo de fuso e fechamento por setor")

    with st.container(border=True):
        col_ano, col_setor, col_add_f, col_del_f = st.columns([1.1, 1.4, 2.1, 2.1])
        with col_ano:
            ano_selecionado = st.selectbox(
                "📅 Ano:", [2024, 2025, 2026, 2027, 2028], index=2, key="sel_ano_fusos"
            )
        with col_setor:
            setor_selecionado = st.selectbox(
                "🏭 Setor:", list(DICIONARIO_SETORES.keys()), key="sel_setor_fusos"
            )

        df_atual_fuso_ctrl = pd.read_excel(ARQUIVO_FUSOS)
        maquinas_do_setor_f = obter_maquinas_setor(setor_selecionado, df_correias, df_atual_fuso_ctrl)

        # Inserção de Máquina em Fusos
        with col_add_f:
            st.markdown("<div style='font-size:0.8rem; font-weight:700; color:#334155; margin-bottom:2px;'>➕ Adicionar Máquina</div>", unsafe_allow_html=True)
            c_inp_af, c_btn_af = st.columns([1.5, 1.0])
            with c_inp_af:
                nova_maq_fuso = st.text_input("Nova Máquina", placeholder="Ex: L-99 ou B-105", key=f"inp_add_fusos_{setor_selecionado}", label_visibility="collapsed").strip().upper()
            with c_btn_af:
                if st.button("Adicionar", key=f"btn_add_fuso_{setor_selecionado}", use_container_width=True):
                    if nova_maq_fuso:
                        if nova_maq_fuso in maquinas_do_setor_f:
                            st.warning(f"A máquina {nova_maq_fuso} já existe no {setor_selecionado}.")
                        else:
                            novos_regs_ano = [
                                {
                                    "Ano": int(ano_selecionado),
                                    "Mes": m_nome,
                                    "Setor": setor_selecionado,
                                    "Maquina_TAG": nova_maq_fuso,
                                    "Quantidade_Quebras": 0,
                                    "Tipo_Fuso": OPCOES_TIPO_FUSO[0],
                                }
                                for m_nome in lista_meses_puros
                            ]
                            df_atual_fuso_ctrl = pd.concat([df_atual_fuso_ctrl, pd.DataFrame(novos_regs_ano)], ignore_index=True)
                            df_atual_fuso_ctrl.to_excel(ARQUIVO_FUSOS, index=False)
                            st.success(f"Máquina {nova_maq_fuso} adicionada ao {setor_selecionado} para {ano_selecionado}!")
                            st.rerun()
                    else:
                        st.warning("Informe o código da máquina.")

        # Exclusão de Máquina em Fusos
        with col_del_f:
            st.markdown("<div style='font-size:0.8rem; font-weight:700; color:#334155; margin-bottom:2px;'>🗑️ Excluir Máquina</div>", unsafe_allow_html=True)
            c_sel_df, c_btn_df = st.columns([1.5, 1.0])
            with c_sel_df:
                maq_del_fuso = st.selectbox("Excluir", ["-- Selecione --"] + maquinas_do_setor_f, key=f"sel_del_fuso_{setor_selecionado}", label_visibility="collapsed")
            with c_btn_df:
                if st.button("Excluir", key=f"btn_del_fuso_{setor_selecionado}", use_container_width=True, type="secondary"):
                    if maq_del_fuso and maq_del_fuso != "-- Selecione --":
                        df_atual_fuso_ctrl = df_atual_fuso_ctrl[~((df_atual_fuso_ctrl["Setor"] == setor_selecionado) & (df_atual_fuso_ctrl["Maquina_TAG"] == maq_del_fuso))]
                        df_atual_fuso_ctrl.to_excel(ARQUIVO_FUSOS, index=False)
                        st.success(f"Máquina {maq_del_fuso} removida do {setor_selecionado}!")
                        st.rerun()
                    else:
                        st.warning("Selecione uma máquina para excluir.")

    st.markdown("<br>", unsafe_allow_html=True)

    abas_meses = st.tabs(lista_meses_puros)
    maquinas_do_setor = obter_maquinas_setor(setor_selecionado, df_correias, df_fusos)

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
