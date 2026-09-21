import streamlit as st

# Configuração inicial da página
st.set_page_config(
    page_title="Sistema PCM",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    st.title("⚙️ Portal PCM")
    st.caption("Planeamento e Controlo de Manutenção")
    st.markdown("---")
    
    # 1. Escolha do Módulo Principal
    modulo = st.selectbox(
        "Selecione o Módulo:",
        ["📊 Painéis & Indicadores", "📝 Registos & Lançamentos"]
    )
    
    st.markdown("---")
    
    # 2. Subpáginas dependendo do módulo selecionado
    if modulo == "📊 Painéis & Indicadores":
        pagina = st.radio(
            "Visualizações Disponíveis:",
            [
                "Visão Geral (KPIs)",
                "Controlo de Backlog",
                "Histórico & Falhas por TAG"
            ]
        )
    else:
        pagina = st.radio(
            "Tipos de Lançamento:",
            [
                "Abertura de OS",
                "Apontamento / Baixa de OS",
                "Cadastro de Equipamentos"
            ]
        )
        
    st.markdown("---")
    st.caption("Utilizador: Analista PCM")

# ==========================================
# ÁREA PRINCIPAL (DIREITA)
# ==========================================

# --- MÓDULO 1: PAINÉIS ---
if modulo == "📊 Painéis & Indicadores":
    if pagina == "Visão Geral (KPIs)":
        st.header("📊 Painel Geral de Manutenção")
        st.write("Aqui ficarão os indicadores consolidados (MTBF, MTTR, Disponibilidade).")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("MTBF Global", "140 h", "+8%")
        c2.metric("MTTR Médio", "2.1 h", "-15%")
        c3.metric("Disponibilidade", "96.4%", "+1.2%")
        
    elif pagina == "Controlo de Backlog":
        st.header("⏳ Acompanhamento de Backlog")
        st.write("Distribuição das ordens pendentes por prioridade e tempo de espera.")
        
    elif pagina == "Histórico & Falhas por TAG":
        st.header("🔍 Análise por Equipamento")
        st.write("Consulte o histórico detalhado de intervenções de uma máquina específica.")

# --- MÓDULO 2: REGISTOS / LANÇAMENTOS ---
else:
    if pagina == "Abertura de OS":
        st.header("📝 Formulário de Abertura de Ordem de Serviço")
        st.write("Registo inicial de novas ordens para o chão de fábrica.")
        
    elif pagina == "Apontamento / Baixa de OS":
        st.header("🔧 Encerramento de Manutenção")
        st.write("Registo de horas trabalhadas, tempo de paragem e troca de componentes.")
        
    elif pagina == "Cadastro de Equipamentos":
        st.header("⚙️ Cadastro de TAGs e Máquinas")
        st.write("Adicione novos equipamentos, setores e criticidades à base de dados.")
