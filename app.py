import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="Painel Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 1. Guarda na memória qual tela está aberta (começa na Visão Geral)
if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Visão Geral (KPIs)"

def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina

# ==========================================
# BARRA LATERAL (SIDEBAR COM BOTÕES)
# ==========================================
with st.sidebar:
    st.title("⚙️ Painel Manutenção")
    st.markdown("---")
    
    # SEÇÃO 1: PAINÉIS
    st.subheader("📊 Painéis & Indicadores")
    
    # Botão 1: Visão Geral
    tipo_btn1 = "primary" if st.session_state.pagina_atual == "Visão Geral (KPIs)" else "secondary"
    st.button("📈 Visão Geral (KPIs)", use_container_width=True, type=tipo_btn1, on_click=navegar, args=("Visão Geral (KPIs)",))
    
    # Botão 2: Backlog
    tipo_btn2 = "primary" if st.session_state.pagina_atual == "Controle de Backlog" else "secondary"
    st.button("⏳ Controle de Backlog", use_container_width=True, type=tipo_btn2, on_click=navegar, args=("Controle de Backlog",))
    
    # Botão 3: Histórico por TAG
    tipo_btn3 = "primary" if st.session_state.pagina_atual == "Histórico por TAG" else "secondary"
    st.button("🔍 Histórico por TAG", use_container_width=True, type=tipo_btn3, on_click=navegar, args=("Histórico por TAG",))
    
    st.markdown("---")
    
    # SEÇÃO 2: ENTRADA DE DADOS / REGISTROS
    st.subheader("📝 Registros & Lançamentos")
    
    # Botão 4: Abertura de OS
    tipo_btn4 = "primary" if st.session_state.pagina_atual == "Abertura de OS" else "secondary"
    st.button("➕ Abertura de OS", use_container_width=True, type=tipo_btn4, on_click=navegar, args=("Abertura de OS",))
    
    # Botão 5: Baixa de OS
    tipo_btn5 = "primary" if st.session_state.pagina_atual == "Baixa / Encerramento" else "secondary"
    st.button("✔️ Baixa / Encerramento", use_container_width=True, type=tipo_btn5, on_click=navegar, args=("Baixa / Encerramento",))
    
    # Botão 6: Cadastro de TAGs
    tipo_btn6 = "primary" if st.session_state.pagina_atual == "Cadastro de TAGs" else "secondary"
    st.button("⚙️ Cadastro de TAGs", use_container_width=True, type=tipo_btn6, on_click=navegar, args=("Cadastro de TAGs",))
    
    st.markdown("---")
    st.caption("Perfil: Analista de PCM")

# ==========================================
# ÁREA DA DIREITA (CONTEÚDO DINÂMICO)
# ==========================================
tela = st.session_state.pagina_atual

if tela == "Visão Geral (KPIs)":
    st.header("📊 Painel Geral de Indicadores")
    c1, c2, c3 = st.columns(3)
    c1.metric("MTBF Global", "140 h", "+8%")
    c2.metric("MTTR Médio", "2.1 h", "-15%")
    c3.metric("Disponibilidade", "96.4%", "+1.2%")

elif tela == "Controle de Backlog":
    st.header("⏳ Controle de Backlog de Ordens")
    st.info("Distribuição de horas pendentes por setor da fábrica.")

elif tela == "Histórico por TAG":
    st.header("🔍 Histórico e Falhas por TAG")
    st.info("Consulte aqui a ficha corrida de cada equipamento.")

elif tela == "Abertura de OS":
    st.header("➕ Formulário de Abertura de OS")
    st.info("Área para abertura de chamados e solicitações.")

elif tela == "Baixa / Encerramento":
    st.header("✔️ Baixa e Apontamento de Manutenção")
    st.info("Área do mecânico/eletricista para fechar a OS e apontar horas.")

elif tela == "Cadastro de TAGs":
    st.header("⚙️ Cadastro de Equipamentos e Máquinas")
    st.info("Base de dados com código do TAG, setor e criticidade.")
