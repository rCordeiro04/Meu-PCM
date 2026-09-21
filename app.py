import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="Painel Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    # Título solicitado
    st.title("⚙️ Painel Manutenção")
    st.markdown("---")
    
    # 1. Seletor de área principal
    modulo = st.selectbox(
        "Selecione a Área:",
        ["📊 Painéis & Indicadores", "📝 Registros & Lançamentos"]
    )
    
    st.markdown("---")
    
    # 2. Subpáginas dinâmicas conforme a escolha
    if modulo == "📊 Painéis & Indicadores":
        pagina = st.radio(
            "Painéis Disponíveis:",
            [
                "Visão Geral (KPIs)",
                "Controle de Backlog",
                "Histórico & Falhas por TAG"
            ]
        )
    else:
        pagina = st.radio(
            "Lançamentos Disponíveis:",
            [
                "Abertura de OS",
                "Apontamento / Baixa de OS",
                "Cadastro de Equipamentos"
            ]
        )
        
    st.markdown("---")
    st.caption("Perfil: Analista de PCM")

# ==========================================
# ÁREA PRINCIPAL (LADO DIREITO)
# ==========================================

# --- MÓDULO 1: PAINÉIS ---
if modulo == "📊 Painéis & Indicadores":
    if pagina == "Visão Geral (KPIs)":
        st.header("📊 Painel Geral de Indicadores")
        st.write("Visão consolidada da operação e desempenho dos ativos.")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("MTBF Global", "140 h", "+8%")
        c2.metric("MTTR Médio", "2.1 h", "-15%")
        c3.metric("Disponibilidade", "96.4%", "+1.2%")
        
    elif pagina == "Controle de Backlog":
        st.header("⏳ Controle de Backlog")
        st.write("Acompanhamento da carga horária de ordens pendentes.")
        
    elif pagina == "Histórico & Falhas por TAG":
        st.header("🔍 Histórico por Equipamento")
        st.write("Consulta detalhada das intervenções por ativo.")

# --- MÓDULO 2: REGISTROS & LANÇAMENTOS ---
else:
    if pagina == "Abertura de OS":
        st.header("📝 Abertura de Ordem de Serviço")
        st.write("Preencha as informações para emitir uma nova solicitação.")
        
    elif pagina == "Apontamento / Baixa de OS":
        st.header("🔧 Baixa e Apontamento de OS")
        st.write("Fechamento de ordens com horas gastas e causa raiz.")
        
    elif pagina == "Cadastro de Equipamentos":
        st.header("⚙️ Cadastro de Equipamentos (TAGs)")
        st.write("Cadastro de novos ativos, criticidades e setores.")
