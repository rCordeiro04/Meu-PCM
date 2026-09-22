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
