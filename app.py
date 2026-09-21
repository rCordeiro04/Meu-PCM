import streamlit as st

# Configura a página para ocupar a tela toda
st.set_page_config(
    page_title="Sistema PCM",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. MENU LATERAL ESQUERDO (Aba de Seleção)
# ==========================================
with st.sidebar:
    st.title("⚙️ Sistema PCM")
    st.markdown("---")
    
    # As opções que você clica para mudar a tela da direita
    pagina = st.radio(
        "Navegação:",
        [
            "📊 Dashboard de KPIs",
            "📝 Abertura de OS",
            "🔧 Baixa de OS",
            "📂 Histórico por TAG",
            "⚙️ Configurações / Cadastros"
        ]
    )
    
    st.markdown("---")
    st.caption("Perfil: Analista de PCM")

# ==========================================
# 2. ÁREA DA DIREITA (Muda conforme o clique)
# ==========================================
if pagina == "📊 Dashboard de KPIs":
    st.header("📊 Painel Geral de Manutenção")
    st.write("Aqui ficarão os cartões com MTTR, MTBF, Disponibilidade e gráficos de parada.")
    
    # Exemplo de cartões rápidos
    c1, c2, c3 = st.columns(3)
    c1.metric("MTBF", "124 h", "+12%")
    c2.metric("MTTR", "2.4 h", "-8%")
    c3.metric("OS Pendentes", "6 OS", "Backlog")

elif pagina == "📝 Abertura de OS":
    st.header("📝 Nova Ordem de Serviço")
    st.write("Preencha os dados abaixo para abrir uma solicitação:")
    
    with st.form("form_abertura"):
        st.text_input("TAG do Equipamento (Ex: BOMBA-01)")
        st.selectbox("Tipo de Manutenção", ["Corretiva", "Preventiva", "Preditiva"])
        st.selectbox("Criticidade", ["Alta (A)", "Média (B)", "Baixa (C)"])
        st.text_area("Descrição da Falha ou Serviço")
        st.form_submit_button("Criar Ordem de Serviço")

elif pagina == "🔧 Baixa de OS":
    st.header("🔧 Encerramento e Apontamento de Horas")
    st.info("Selecione uma OS aberta para apontar horas paradas, técnicos e causa raiz.")

elif pagina == "📂 Histórico por TAG":
    st.header("📂 Histórico do Equipamento")
    st.selectbox("Selecione o Equipamento:", ["BOMBA-01", "COMPRESSOR-02", "ESTEIRA-05"])
    st.write("Aqui aparecerão todas as ordens passadas desse equipamento.")

elif pagina == "⚙️ Configurações / Cadastros":
    st.header("⚙️ Cadastros do Sistema")
    st.write("Área para cadastrar novas máquinas, setores e equipe de mecânicos/eletricistas.")
