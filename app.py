from datetime import date
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Painel Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_FUSOS = "lancamentos_fusos_v3.xlsx"

colunas_obrigatorias = [
    "Ano",
    "Mes",
    "Maquina_TAG",
    "Quantidade_Quebras",
    "Observacoes",
]

# Inicializa o arquivo com a estrutura correta se não existir
if not os.path.exists(ARQUIVO_FUSOS):
    pd.DataFrame(columns=colunas_obrigatorias).to_excel(
        ARQUIVO_FUSOS, index=False
    )

df_fusos = pd.read_excel(ARQUIVO_FUSOS)
if not all(col in df_fusos.columns for col in colunas_obrigatorias):
    df_fusos = pd.DataFrame(columns=colunas_obrigatorias)
    df_fusos.to_excel(ARQUIVO_FUSOS, index=False)

# Controle de navegação da página
if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "Lançamento Fusos"


def navegar(nome_pagina):
    st.session_state.pagina_atual = nome_pagina


# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    st.title("⚙️ Painel Manutenção")
    st.markdown("---")

    # PAINÉIS
    st.subheader("📊 Painéis")
    tipo_painel_fusos = (
        "primary"
        if st.session_state.pagina_atual == "Painel Fusos"
        else "secondary"
    )
    st.button(
        "🔩 Fusos",
        use_container_width=True,
        type=tipo_painel_fusos,
        on_click=navegar,
        args=("Painel Fusos",),
    )

    st.markdown("---")

    # LANÇAMENTOS
    st.subheader("📝 Lançamentos")
    tipo_fusos = (
        "primary"
        if st.session_state.pagina_atual == "Lançamento Fusos"
        else "secondary"
    )
    st.button(
        "🔩 Fusos",
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
        use_container_width=True,
        type=tipo_maq,
        on_click=navegar,
        args=("Máquinas",),
    )

    st.markdown("---")
    st.caption("Perfil: Analista de PCM")

# ==========================================
# ÁREA PRINCIPAL (TELA DA DIREITA)
# ==========================================
tela = st.session_state.pagina_atual

lista_meses_puros = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
]

# ------------------------------------------
# 1. LANÇAMENTOS: FUSOS (LISTA DE MÁQUINAS 1 A 10 POR MÊS)
# ------------------------------------------
if tela == "Lançamento Fusos":
    st.header("🔩 Lançamentos: Grade de Quebras de Fusos")
    st.write(
        "Escolha o ano, selecione a aba do mês correspondente no topo e aponte a quantidade de quebras para as máquinas de 1 a 10."
    )

    # 1. Seleção do Ano no topo
    col_ano, _ = st.columns([2, 4])
    with col_ano:
        ano_selecionado = st.selectbox(
            "📅 Selecione o Ano:", [2024, 2025, 2026, 2027, 2028], index=2
        )

    st.markdown("---")

    # 2. Abas horizontais dos 12 meses na parte superior
    abas_meses = st.tabs(lista_meses_puros)

    # Lista fixa das máquinas de 1 a 10
    lista_maquinas = [f"Máquina {i}" for i in range(1, 11)]

    for idx, nome_mes in enumerate(lista_meses_puros):
        with abas_meses[idx]:
            st.subheader(f"Apontamento: {nome_mes} / {ano_selecionado}")

            # Lê os dados atuais
            df_atual = pd.read_excel(ARQUIVO_FUSOS)

            # Filtra os dados já salvos desse ano e mês
            df_mes_ano = df_atual[
                (df_atual["Ano"] == ano_selecionado)
                & (df_atual["Mes"] == nome_mes)
            ]

            # Monta a tabela base com as 10 máquinas
            dados_grade = []
            for maq in lista_maquinas:
                # Se já tiver registro salvo, recupera o valor; senão coloca 0
                registro_existente = df_mes_ano[df_mes_ano["Maquina_TAG"] == maq]
                if not registro_existente.empty:
                    qtd = int(registro_existente.iloc[0]["Quantidade_Quebras"])
                    obs = str(registro_existente.iloc[0]["Observacoes"])
                    if obs == "nan":
                        obs = ""
                else:
                    qtd = 0
                    obs = ""

                dados_grade.append(
                    {
                        "Máquina": maq,
                        "Quantidade de Quebras": qtd,
                        "Observações": obs,
                    }
                )

            df_grade = pd.DataFrame(dados_grade)

            st.write(
                "👉 Dê dois cliques no número para editar a quantidade de quebras de cada máquina:"
            )

            # Editor interativo na tela
            tabela_editada = st.data_editor(
                df_grade,
                disabled=["Máquina"],  # Trava o nome das máquinas para ninguém apagar
                hide_index=True,
                use_container_width=True,
                key=f"editor_{ano_selecionado}_{nome_mes}",
            )

            col_btn, col_info = st.columns([2, 4])
            with col_btn:
                salvar_mes = st.button(
                    f"💾 Salvar {nome_mes}/{ano_selecionado}",
                    key=f"btn_{ano_selecionado}_{nome_mes}",
                    type="primary",
                )

            if salvar_mes:
                # 1. Remove qualquer registro antigo desse mês/ano para não duplicar
                df_limpo = df_atual[
                    ~(
                        (df_atual["Ano"] == ano_selecionado)
                        & (df_atual["Mes"] == nome_mes)
                    )
                ]

                # 2. Prepara os novos dados da tabela editada
                novos_registros = []
                for _, linha in tabela_editada.iterrows():
                    novos_registros.append(
                        {
                            "Ano": int(ano_selecionado),
                            "Mes": nome_mes,
                            "Maquina_TAG": linha["Máquina"],
                            "Quantidade_Quebras": int(
                                linha["Quantidade de Quebras"]
                            ),
                            "Observacoes": str(linha["Observações"]),
                        }
                    )

                df_final = pd.concat(
                    [df_limpo, pd.DataFrame(novos_registros)], ignore_index=True
                )
                df_final.to_excel(ARQUIVO_FUSOS, index=False)
                st.success(
                    f"✅ Fechamento de {nome_mes}/{ano_selecionado} salvo com sucesso!"
                )
                st.rerun()

            # Métricas rápidas logo abaixo da tabela
            total_quebras_mes = tabela_editada["Quantidade de Quebras"].sum()
            st.caption(
                f"Total de quebras acumuladas em {nome_mes}: **{total_quebras_mes} fusos**"
            )

# ------------------------------------------
# 2. PAINEL GERENCIAL DE FUSOS
# ------------------------------------------
elif tela == "Painel Fusos":
    st.header("📊 Painel Gerencial: Quebras de Fusos")
    st.write("Análise visual consolidada das quebras apontadas nas máquinas.")

    col_ano, col_mes = st.columns([1, 2])
    lista_anos = ["Todos", "2024", "2025", "2026", "2027", "2028"]
    lista_meses = ["Todos"] + lista_meses_puros

    with col_ano:
        ano_painel = st.selectbox(
            "📅 Filtrar por Ano:", lista_anos, index=3, key="p_ano"
        )
    with col_mes:
        mes_painel = st.selectbox(
            "🗓️ Filtrar por Mês:", lista_meses, index=0, key="p_mes"
        )

    st.markdown("---")

    df_dados = pd.read_excel(ARQUIVO_FUSOS)
    df_filtrado_p = df_dados.copy()

    if ano_painel != "Todos":
        df_filtrado_p = df_filtrado_p[df_filtrado_p["Ano"] == int(ano_painel)]
    if mes_painel != "Todos":
        df_filtrado_p = df_filtrado_p[df_filtrado_p["Mes"] == mes_painel]

    # Filtra apenas quem tem quebras maior que zero para o ranking
    df_quebras_reais = df_filtrado_p[df_filtrado_p["Quantidade_Quebras"] > 0]

    if not df_quebras_reais.empty:
        total_quebras = df_filtrado_p["Quantidade_Quebras"].sum()
        total_maquinas_falharam = df_quebras_reais["Maquina_TAG"].nunique()

        agrupado_maq = (
            df_filtrado_p.groupby("Maquina_TAG")["Quantidade_Quebras"]
            .sum()
            .reset_index()
        )
        agrupado_maq = agrupado_maq.sort_values(
            by="Quantidade_Quebras", ascending=False
        )

        maquina_top = agrupado_maq.iloc[0]["Maquina_TAG"]
        qtd_top = agrupado_maq.iloc[0]["Quantidade_Quebras"]

        m1, m2, m3 = st.columns(3)
        m1.metric("Total de Fusos Quebrados", f"{total_quebras} unid.")
        m2.metric("Máquinas com Quebras", f"{total_maquinas_falharam} de 10")
        m3.metric("Maior Ofensor", f"{maquina_top} ({qtd_top} quebras)")

        st.markdown("---")

        graf_col, tab_col = st.columns([2, 1])
        with graf_col:
            st.subheader("Ranking de Falhas")
            st.bar_chart(
                data=agrupado_maq.set_index("Maquina_TAG")[
                    "Quantidade_Quebras"
                ]
            )

        with tab_col:
            st.subheader("Resumo por Máquina")
            st.dataframe(
                agrupado_maq.rename(
                    columns={
                        "Maquina_TAG": "Máquina",
                        "Quantidade_Quebras": "Total",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
    else:
        st.info("Nenhuma quebra registrada para o período selecionado.")

# DEMAIS TELAS MANTIDAS
elif tela == "Correias":
    st.header("🔄 Lançamentos: Correias")
elif tela == "Preventiva":
    st.header("🛠️ Lançamentos: Preventiva")
elif tela == "Máquinas":
    st.header("🏭 Lançamentos: Máquinas")
