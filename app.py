import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="Painel Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 1. Guarda na memória qual tela está aberta
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
    st.subheader("📊 Painéis")
    
    tipo_btn1 = "primary" if st.session_state.pagina_atual == "Visão Geral (KPIs)" else "secondary"
    st.button("📈 Visão Geral (KPIs)", use_container_width=True, type=tipo_btn1, on_click=navegar, args=("Visão Geral (KPIs)",))
    
    tipo_btn2 = "primary" if st.session_state.pagina_atual == "Controle de Backlog" else "secondary"
    st.button("⏳ Controle de Backlog", use_container_width=True, type=tipo_btn2, on_click=navegar, args=("Controle de Backlog",))
    
    tipo_btn3 = "primary" if st.session_state.pagina_atual == "Histórico por TAG" else "secondary"
    st.button("🔍 Histórico por TAG", use_container_width=True, type=tipo_btn3, on_click=navegar, args=("Histórico por TAG",))
    
    st.markdown("---")
    
    # SEÇÃO 2: LANÇAMENTOS (NOVOS ITENS)
    st.subheader("📝 Lançamentos")
    
    tipo_fusos = "primary" if st.session_state.pagina_atual == "Fusos" else "secondary"
    st.button("🔩 Fusos", use_container_width=True, type=tipo_fusos, on_click=navegar, args=("Fusos",))
    
    tipo_correias = "primary" if st.session_state.pagina_atual == "Correias" else "secondary"
    st.button("🔄 Correias", use_container_width=True, type=tipo_correias, on_click=navegar, args=("Correias",))
    
    tipo_preventiva = "primary" if st.session_state.pagina_atual == "Preventiva" else "secondary"
    st.button("🛠️ Preventiva", use_container_width=True, type=tipo_preventiva, on_click=navegar, args=("Preventiva",))
    
    tipo_maquinas = "primary" if st.session_state.pagina_atual == "Máquinas" else "secondary"
    st.button("🏭 Máquinas", use_container_width=True, type=tipo_maquinas, on_click=navegar, args=("Máquinas",))
    
    st.markdown("---")
    st.caption("Perfil: Analista de PCM")

# ==========================================
# ÁREA DA DIREITA (CONTEÚDO DINÂMICO)
# ==========================================
tela = st.session_state.pagina_atual

# --- TELAS DOS PAINÉIS ---
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
    st.info("Consulte aqui a ficha de manutenção de cada equipamento.")

# --- TELAS DOS NOVOS LANÇAMENTOS ---
elif tela == "Fusos":
    st.header("🔩 Lançamentos: Troca e Controle de Fusos")
    st.write("Área para registro de trocas, lubrificação e vida útil de fusos.")

elif tela == "Correias":
    st.header("🔄 Lançamentos: Controle de Correias")
    st.write("Registro de tensionamento, trocas e inspeção de correias.")

elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Planos de Manutenção Preventiva")
    st.write("Checklists e fechamentos de rotinas preventivas periódicas.")

elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Cadastro e Dados de Máquinas")
    st.write("Cadastro de TAGs, capacidades, setores e criticidades.")
