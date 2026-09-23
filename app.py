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

# Estilização CSS global para cartões, KPIs e harmonia visual
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
        .pill-legenda {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: #f1f5f9;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
            color: #334155;
            margin-left: 6px;
        }
        .dot-legenda {
            width: 8px;
            height: 8px;
            border-radius: 50%;
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

DICIONARIO_SETORES = {
    "Setor A": [f"L-{i:02d}" for i in range(1, 29)],
    "Setor B": [f"L-{i:02d}" for i in [18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 50, 51, 52, 53] if i <= 46 or i >= 50],
    "Setor Látex": [f"B-{i:03d}" for i in list(range(71, 81)) + list(range(83, 90)) + list(range(102, 105))],
    "Setor Menegatto": [f"B-{i:03d}" for i in list(range(47, 50)) + list(range(81, 83)) + list(range(90, 102)) + [107, 108]],
}

# Ajuste fino das listas reais baseadas no parque oficial
DICIONARIO_SETORES["Setor A"] = [f"L-{i:02d}" for i in range(1, 29)]
DICIONARIO_SETORES["Setor B"] = [f"L-{i:02d}" for i in list(range(29, 47)) + list(range(50, 54))]
DICIONARIO_SETORES["Setor Látex"] = [f"B-{i:02d}" if i < 100 else f"B-{i}" for i in list(range(71, 81)) + list(range(83, 90)) + list(range(102, 105))]
DICIONARIO_SETORES["Setor Látex"] = ["B-71", "B-72", "B-73", "B-74", "B-75", "B-76", "B-77", "B-78", "B-79", "B-80", "B-83", "B-84", "B-85", "B-86", "B-87", "B-88", "B-89", "B-102", "B-103", "B-104"]
DICIONARIO_SETORES["Setor Menegatto"] = ["B-47", "B-48", "B-49", "B-81", "B-82", "B-90", "B-91", "B-92", "B-93", "B-94", "B-95", "B-96", "B-97", "B-98", "B-99", "B-100", "B-101"]

if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Painel Fusos"
if "aba_setor_fuso" not in st.session_state:
    st.session_state.aba_setor_fuso = "Geral"

def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina

### ==========================================
### LEITURA E RECUPERAÇÃO AUTOMÁTICA DAS BASES
### ==========================================
precisa_salvar_fusos = False

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

### ==========================================
### BARRA LATERAL (SIDEBAR) REFORMULADA
### ==========================================
with st.sidebar:
    st.title("⚙️ Portal PCM")
    st.caption("Planejamento e Controle de Manutenção")
    st.markdown("<div style='margin-top: -10px;'></div>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("<div style='font-size:0.85rem; font-weight:700; color:#64748b; margin-bottom:6px;'>📊 PAINÉIS GERENCIAIS</div>", unsafe_allow_html=True)
    col_sb_1, col_sb_2 = st.columns(2)
    with col_sb_1:
        t_pf = "primary" if st.session_state.pagina_atual == "Painel Fusos" else "secondary"
        if st.button("🔩 Fusos", key="btn_nav_painel_fusos", use_container_width=True, type=t_pf):
            navegar("Painel Fusos")
            st.rerun()
    with col_sb_2:
        t_pc = "primary" if st.session_state.pagina_atual == "Painel Correias" else "secondary"
        if st.button("🔄 Correias", key="btn_nav_painel_correias", use_container_width=True, type=t_pc):
            navegar("Painel Correias")
            st.rerun()

    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.85rem; font-weight:700; color:#64748b; margin-bottom:6px;'>📝 MÓDULOS DE LANÇAMENTO</div>", unsafe_allow_html=True)
    
    col_sb_3, col_sb_4 = st.columns(2)
    with col_sb_3:
        t_lf = "primary" if st.session_state.pagina_atual == "Lançamento Fusos" else "secondary"
        if st.button("🔩 Fusos", key="btn_nav_lancto_fusos", use_container_width=True, type=t_lf):
            navegar("Lançamento Fusos")
            st.rerun()
    with col_sb_4:
        t_lc = "primary" if st.session_state.pagina_atual == "Correias" else "secondary"
        if st.button("🔄 Correias", key="btn_nav_lancto_correias", use_container_width=True, type=t_lc):
            navegar("Correias")
            st.rerun()

    col_sb_5, col_sb_6 = st.columns(2)
    with col_sb_5:
        t_prev = "primary" if st.session_state.pagina_atual == "Preventiva" else "secondary"
        if st.button("🛠️ Preventiva", key="btn_nav_lancto_preventiva", use_container_width=True, type=t_prev):
            navegar("Preventiva")
            st.rerun()
    with col_sb_6:
        t_maq = "primary" if st.session_state.pagina_atual == "Máquinas" else "secondary"
        if st.button("🏭 Máquinas", key="btn_nav_lancto_maquinas", use_container_width=True, type=t_maq):
            navegar("Máquinas")
            st.rerun()

    st.markdown("---")
    st.caption("PCM • Versão Gerencial Corporativa")

### ==========================================
### ÁREA PRINCIPAL
### ==========================================
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

### ------------------------------------------
### 1. PAINEL GERENCIAL DE CORREIAS
### ------------------------------------------
if tela == "Painel Correias":
    st.title("🔄 Dashboard Correias")
    st.caption("Acompanhamento técnico e vida útil das correias instaladas nas máquinas")
    st.info("Painel de correias ativo e operando com padrão decimal de 3 dígitos.")

### ------------------------------------------
### 2. LANÇAMENTOS: CORREIAS (COM ADICIONAR/EXCLUIR MÁQUINAS)
### ------------------------------------------
elif tela == "Correias":
    st.title("🔄 Lançamento: Gestão de Correias")
    st.caption("Cadastro contínuo de modelos e datas de instalação por máquina (Superior / Cabeceira e Inferior / Traseira)")

    with st.container(border=True):
        col_setor, col_nova_maq, col_exc_maq = st.columns([2, 2, 2])
        with col_setor:
            setor_selecionado = st.selectbox(
                "🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_setor_correias_direto"
            )
        with col_nova_maq:
            st.markdown("<div style='height: 27px;'></div>", unsafe_allow_html=True)
            nova_tag_cor = st.text_input("Nova Máquina (TAG):", key="input_nova_tag_cor", placeholder="Ex: L-29 ou B-105")
            if st.button("➕ Adicionar Máquina", key="btn_add_maq_cor"):
                if nova_tag_cor and nova_tag_cor.strip():
                    tag_limpa = nova_tag_cor.strip().upper()
                    if tag_limpa not in DICIONARIO_SETORES[setor_selecionado]:
                        DICIONARIO_SETORES[setor_selecionado].append(tag_limpa)
                        df_cor_atual = pd.read_excel(ARQUIVO_CORREIAS)
                        novo_reg_linha = pd.DataFrame([{
                            "Setor": setor_selecionado,
                            "Maquina_TAG": tag_limpa,
                            "Tipo_Correia_1": "",
                            "Data_Instalacao_1": "",
                            "Tipo_Correia_2": "",
                            "Data_Instalacao_2": ""
                        }])
                        df_cor_atual = pd.concat([df_cor_atual, novo_reg_linha], ignore_index=True)
                        df_cor_atual.to_excel(ARQUIVO_CORREIAS, index=False)
                        st.success(f"Máquina {tag_limpa} adicionada com sucesso ao {setor_selecionado}!")
                        st.rerun()
                    else:
                        st.warning("Esta máquina já existe neste setor!")
        with col_exc_maq:
            st.markdown("<div style='height: 27px;'></div>", unsafe_allow_html=True)
            maq_para_excluir = st.selectbox("Excluir Máquina:", ["Selecione..."] + DICIONARIO_SETORES[setor_selecionado], key="sel_exc_maq_cor")
            if st.button("🗑️ Remover Máquina", key="btn_exc_maq_cor"):
                if maq_para_excluir != "Selecione...":
                    if maq_para_excluir in DICIONARIO_SETORES[setor_selecionado]:
                        DICIONARIO_SETORES[setor_selecionado].remove(maq_para_excluir)
                    df_cor_atual = pd.read_excel(ARQUIVO_CORREIAS)
                    df_cor_atual = df_cor_atual[~((df_cor_atual["Setor"] == setor_selecionado) & (df_cor_atual["Maquina_TAG"] == maq_para_excluir))]
                    df_cor_atual.to_excel(ARQUIVO_CORREIAS, index=False)
                    st.success(f"Máquina {maq_para_excluir} removida com sucesso!")
                    st.rerun()

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
        novos_regs_cor = []
        for _, linha in tabela_editada_cor.iterrows():
            d1_str = str(linha["Data (Superior / Cabeceira)"]) if pd.notna(linha["Data (Superior / Cabeceira)"]) else ""
            d2_str = str(linha["Data (Inferior / Traseira)"]) if pd.notna(linha["Data (Inferior / Traseira)"]) else ""
            novos_regs_cor.append({
                "Setor": setor_selecionado,
                "Maquina_TAG": linha["Máquina"],
                "Tipo_Correia_1": formatar_modelo(linha["Modelo (Superior / Cabeceira)"]),
                "Data_Instalacao_1": d1_str,
                "Tipo_Correia_2": formatar_modelo(linha["Modelo (Inferior / Traseira)"]),
                "Data_Instalacao_2": d2_str,
            })
        df_final_cor = pd.concat([df_limpo_cor, pd.DataFrame(novos_regs_cor)], ignore_index=True)
        df_final_cor.to_excel(ARQUIVO_CORREIAS, index=False)
        st.success(f"✅ Apontamentos de correias do {setor_selecionado} salvos com sucesso!")
        st.rerun()

### ------------------------------------------
### 3. PAINEL GERENCIAL DE FUSOS
### ------------------------------------------
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

    # CASO 1: ABA GERAL (FÁBRICA COMPLETA)
    if st.session_state.aba_setor_fuso == "Geral":
        total_geral_quebras = int(df_fuso_ano["Quantidade_Quebras"].sum()) if not df_fuso_ano.empty else 0
        media_mensal_fabrica = round(total_geral_quebras / meses_divisor, 1)

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

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            with st.container(border=True):
                st.subheader("📊 Quebras por Setor")
                if not df_fuso_ano.empty:
                    df_setores_tot = df_fuso_ano.groupby("Setor")["Quantidade_Quebras"].sum().reset_index()
                    chart_setores = alt.Chart(df_setores_tot).mark_bar(color="#3b82f6", cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                        x=alt.X("Setor:N", title="Setor", sort=None),
                        y=alt.Y("Quantidade_Quebras:Q", title="Total de Quebras"),
                        tooltip=["Setor", "Quantidade_Quebras"]
                    ).properties(height=280)
                    st.altair_chart(chart_setores, use_container_width=True)
                else:
                    st.info("Sem dados para exibir no período.")
        with col_g2:
            with st.container(border=True):
                st.subheader("🍩 Distribuição por Tipo de Fuso")
                if not df_fuso_ano.empty:
                    df_tipos_tot = df_fuso_ano.groupby("Tipo_Fuso")["Quantidade_Quebras"].sum().reset_index()
                    chart_donut = alt.Chart(df_tipos_tot).mark_arc(innerRadius=60, outerRadius=100).encode(
                        theta=alt.Theta(field="Quantidade_Quebras", type="quantitative"),
                        color=alt.Color(field="Tipo_Fuso", type="nominal", legend=alt.Legend(title="Tipo de Fuso")),
                        tooltip=["Tipo_Fuso", "Quantidade_Quebras"]
                    ).properties(height=280)
                    st.altair_chart(chart_donut, use_container_width=True)
                else:
                    st.info("Sem dados para exibir no período.")

    # CASO 2: ABAS ESPECÍFICAS POR SETOR (Setor A, Setor B, Setor Látex, Setor Menegatto)
    else:
        setor_ativo = st.session_state.aba_setor_fuso
        df_setor = df_fuso_ano[df_fuso_ano["Setor"] == setor_ativo].copy()

        total_quebras_setor = int(df_setor["Quantidade_Quebras"].sum()) if not df_setor.empty else 0
        media_setor = round(total_quebras_setor / meses_divisor, 1)
        total_maqs_setor = len(DICIONARIO_SETORES.get(setor_ativo, []))

        ultimo_mes_setor = "Nenhum"
        quebras_ultimo_mes = 0
        maior_ofensor_mes = "Nenhuma"

        if not df_setor.empty and total_quebras_setor > 0:
            df_reais_setor = df_setor[df_setor["Quantidade_Quebras"] > 0]
            for m_teste in reversed(lista_meses_puros):
                df_sub_setor = df_reais_setor[df_reais_setor["Mes"] == m_teste]
                if not df_sub_setor.empty and df_sub_setor["Quantidade_Quebras"].sum() > 0:
                    ultimo_mes_setor = m_teste
                    quebras_ultimo_mes = int(df_sub_setor["Quantidade_Quebras"].sum())
                    agrup_maq_mes = df_sub_setor.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().reset_index().sort_values(by="Quantidade_Quebras", ascending=False)
                    maior_ofensor_mes = f"{agrup_maq_mes.iloc[0]['Maquina_TAG']} ({int(agrup_maq_mes.iloc[0]['Quantidade_Quebras'])})"
                    break

        taxa_maq_ult = round(quebras_ultimo_mes / total_maqs_setor, 2) if total_maqs_setor > 0 else 0.0

        sk1, sk2, sk3, sk4 = st.columns(4)
        with sk1:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-total">
                    <div>
                        <div class="kpi-lbl">Total Quebras ({setor_ativo})</div>
                        <div class="kpi-val" style="color:#0f172a;">{total_quebras_setor}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">🔩</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with sk2:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-ok">
                    <div>
                        <div class="kpi-lbl">Média Mensal ({desc_meses_divisor})</div>
                        <div class="kpi-val" style="color:#059669;">{media_setor}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">📈</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with sk3:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-warn">
                    <div>
                        <div class="kpi-lbl">Quebras / Máq ({ultimo_mes_setor})</div>
                        <div class="kpi-val" style="color:#d97706;">{taxa_maq_ult}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">⚙️</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with sk4:
            st.markdown(
                f"""
                <div class="card-kpi-bonito c-crit">
                    <div>
                        <div class="kpi-lbl">Maior Ofensora ({ultimo_mes_setor})</div>
                        <div class="kpi-val" style="color:#dc2626; font-size:1.1rem;">{maior_ofensor_mes}</div>
                    </div>
                    <div style="font-size:1.5rem; opacity:0.8;">🚨</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # Mapa de Calor Operacional (Estilo Planilha: Máquinas na Horizontal, Meses na Vertical ou vice-versa conforme solicitado)
        with st.container(border=True):
            st.subheader(f"🗺️ Mapa Térmico Operacional — {setor_ativo} (Máquinas vs. Meses)")
            st.caption("Visão matricial consolidada com máquinas na horizontal e dias/meses na vertical para simular a planilha antiga.")

            maquinas_lista = DICIONARIO_SETORES.get(setor_ativo, [])
            if not df_setor.empty and len(maquinas_lista) > 0:
                # Pivot table: Máquinas nas colunas, Meses nas linhas
                pivot_heatmap = df_setor.pivot_table(
                    index="Mes",
                    columns="Maquina_TAG",
                    values="Quantidade_Quebras",
                    aggfunc="sum",
                    fill_value=0
                )
                # Reindexar meses na ordem correta
                meses_presentes = [m for m in lista_meses_puros if m in pivot_heatmap.index]
                pivot_heatmap = pivot_heatmap.reindex(index=meses_presentes, columns=maquinas_lista, fill_value=0)
                
                # Exibir tabela interativa formatada
                st.dataframe(pivot_heatmap, use_container_width=True)
            else:
                st.info("Nenhum apontamento registrado para este setor no ano selecionado.")

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # Diagnóstico por Tipo de Fuso
        tipos_disponiveis_setor = OPCOES_TIPO_FUSO
        if not df_setor.empty:
            t_unicos = df_setor["Tipo_Fuso"].dropna().unique().tolist()
            if t_unicos:
                tipos_disponiveis_setor = t_unicos

        with st.container(border=True):
            col_diag_t, col_diag_sel = st.columns([2.5, 1.5])
            with col_diag_t:
                st.markdown(
                    f"""
                    <div style="font-size:1.05rem; font-weight:800; color:#0f172a; display:flex; align-items:center; gap:8px;">
                        <span>🔬 Desempenho Operacional por Tipo de Fuso — {setor_ativo}</span>
                    </div>
                    <div style="font-size:0.8rem; color:#64748b; font-weight:600; margin-top:2px;">
                        Acompanhe o mapa térmico de máquinas atreladas especificamente a cada tipo/marca de fuso ao longo dos meses.
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
            df_tipo_especifico = df_setor[df_setor["Tipo_Fuso"] == fuso_selecionado_analise].copy()
            tot_fuso_sel = int(df_tipo_especifico["Quantidade_Quebras"].sum()) if not df_tipo_especifico.empty else 0
            media_fuso_sel = round(tot_fuso_sel / meses_divisor, 1)

            maqs_vinculadas = sorted(df_tipo_especifico["Maquina_TAG"].unique().tolist()) if not df_tipo_especifico.empty else []
            qtd_maqs_vinculadas = len(maqs_vinculadas)

            df_tipo_falhas = df_tipo_especifico[df_tipo_especifico["Quantidade_Quebras"] > 0]
            if not df_tipo_falhas.empty:
                agrup_top_fuso = df_tipo_falhas.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().reset_index().sort_values(by="Quantidade_Quebras", ascending=False)
                top_maq_fuso = agrup_top_fuso.iloc[0]["Maquina_TAG"]
                qtd_top_fuso = int(agrup_top_fuso.iloc[0]["Quantidade_Quebras"])
            else:
                top_maq_fuso = "Nenhuma"
                qtd_top_fuso = 0

            mini1, mini2, mini3, mini4 = st.columns(4)
            with mini1:
                st.metric("Total Quebras (Fuso)", tot_fuso_sel)
            with mini2:
                st.metric(f"Média Mensal ({desc_meses_divisor})", media_fuso_sel)
            with mini3:
                st.metric("Máquinas Vinculadas", qtd_maqs_vinculadas)
            with mini4:
                st.metric("Maior Ofensora (Fuso)", f"{top_maq_fuso} ({qtd_top_fuso})")

### ------------------------------------------
### 4. LANÇAMENTOS: FUSOS (COM ADICIONAR/EXCLUIR MÁQUINAS)
### ------------------------------------------
elif tela == "Lançamento Fusos":
    st.title("🔩 Lançamento: Fechamento Mensal de Fusos")
    st.caption("Preenchimento rápido de quebras por máquina, seleção de tipo de fuso e fechamento por setor")

    with st.container(border=True):
        col_ano, col_setor, col_nova_maq_f, col_exc_maq_f = st.columns([1.5, 2, 2, 2])
        with col_ano:
            ano_selecionado = st.selectbox(
                "📅 Ano de Fechamento:", [2024, 2025, 2026, 2027, 2028], index=2, key="sel_ano_fusos"
            )
        with col_setor:
            setor_selecionado = st.selectbox(
                "🏭 Setor Operacional:", list(DICIONARIO_SETORES.keys()), key="sel_setor_fusos"
            )
        with col_nova_maq_f:
            st.markdown("<div style='height: 27px;'></div>", unsafe_allow_html=True)
            nova_tag_fuso = st.text_input("Nova Máquina (TAG):", key="input_nova_tag_fuso", placeholder="Ex: L-29 ou B-105")
            if st.button("➕ Adicionar Máquina", key="btn_add_maq_fuso"):
                if nova_tag_fuso and nova_tag_fuso.strip():
                    tag_limpa_f = nova_tag_fuso.strip().upper()
                    if tag_limpa_f not in DICIONARIO_SETORES[setor_selecionado]:
                        DICIONARIO_SETORES[setor_selecionado].append(tag_limpa_f)
                        st.success(f"Máquina {tag_limpa_f} adicionada com sucesso ao {setor_selecionado}!")
                        st.rerun()
                    else:
                        st.warning("Esta máquina já existe neste setor!")
        with col_exc_maq_f:
            st.markdown("<div style='height: 27px;'></div>", unsafe_allow_html=True)
            maq_exc_fuso = st.selectbox("Excluir Máquina:", ["Selecione..."] + DICIONARIO_SETORES[setor_selecionado], key="sel_exc_maq_fuso")
            if st.button("🗑️ Remover Máquina", key="btn_exc_maq_fuso"):
                if maq_exc_fuso != "Selecione...":
                    if maq_exc_fuso in DICIONARIO_SETORES[setor_selecionado]:
                        DICIONARIO_SETORES[setor_selecionado].remove(maq_exc_fuso)
                    df_fuso_atual = pd.read_excel(ARQUIVO_FUSOS)
                    df_fuso_atual = df_fuso_atual[~((df_fuso_atual["Setor"] == setor_selecionado) & (df_fuso_atual["Maquina_TAG"] == maq_exc_fuso))]
                    df_fuso_atual.to_excel(ARQUIVO_FUSOS, index=False)
                    st.success(f"Máquina {maq_exc_fuso} removida com sucesso!")
                    st.rerun()

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
                f"Total de fusos apontados no {setor_selecionado} em {nome_mes}:  **{total_mes} unid.** "
            )

elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
    st.info("Módulo de preventivas em desenvolvimento.")

elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Cadastro de Máquinas")
    st.info("Módulo de cadastro de ativos corporativos.")
