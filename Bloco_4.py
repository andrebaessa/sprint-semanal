
import streamlit as st
import pandas as pd
import zipfile
import io
import datetime as dt

st.set_page_config(page_title="📊 Dashboard - SPRINT Semanal", layout="wide", page_icon="")

st.markdown("""
    <style>
        section[data-testid="stSidebar"] { background-color: #f8f8f8; }
        .metric-box { border: 1px solid #ddd; padding: 1rem; border-radius: 10px; background-color: white; box-shadow: 0px 0px 4px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard - SPRINT Semanal")
st.caption("Suba as planilhas da semana para calcular o ranking.")

col1, col2 = st.columns(2)
with col1:
    data_inicio = st.date_input("Início da janela de observação", dt.date.today() - dt.timedelta(days=7))
with col2:
    data_fim = st.date_input("Fim da janela de observação", dt.date.today())

# Upload de ZIP
uploaded_zip = st.file_uploader("📁 Envie o arquivo ZIP com as planilhas (.xlsx e .csv)", type="zip")

# Arquivos obrigatórios
required_files = {
    "indicadores_uso.xlsx", "indicacoes.csv", "base_btg.xlsx",
    "nnm.xlsx", "pace_nnm.xlsx", "base_receita.xlsx",
    "pace_receita.xlsx", "nps.xlsx", "participacao_time.xlsx",
    "mapa_nomes.xlsx", "outras_receitas.xlsx"
}

def read_file_from_zip(zip_ref, filename):
    with zip_ref.open(filename) as f:
        if filename.endswith(".csv"):
            return pd.read_csv(f)
        else:
            return pd.read_excel(f)

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
        file_list = set(zip_ref.namelist())
        missing_files = required_files - file_list

        if missing_files:
            st.error(f"Arquivos faltando no .zip: {', '.join(missing_files)}")
        else:
            st.success("✅ Todos os arquivos encontrados no ZIP. Processando...")

            # Leitura dos arquivos em memória
            df_ind_uso = read_file_from_zip(zip_ref, "indicadores_uso.xlsx")
            df_indicacoes = read_file_from_zip(zip_ref, "indicacoes.csv")
            df_participacao = read_file_from_zip(zip_ref, "participacao_time.xlsx")
            df_mapa_nomes = read_file_from_zip(zip_ref, "mapa_nomes.xlsx")
            df_btg = read_file_from_zip(zip_ref, "base_btg.xlsx")
            df_nnm = read_file_from_zip(zip_ref, "nnm.xlsx")
            df_pace_nnm = read_file_from_zip(zip_ref, "pace_nnm.xlsx")
            df_receita = read_file_from_zip(zip_ref, "base_receita.xlsx")
            df_pace_receita = read_file_from_zip(zip_ref, "pace_receita.xlsx")
            df_nps = read_file_from_zip(zip_ref, "nps.xlsx")
            df_outras_receitas = read_file_from_zip(zip_ref, "outras_receitas.xlsx")

            st.success("🧩 Planilhas carregadas com sucesso! Agora é só seguir com os cálculos...")

    # Mapeamento de nomes
    mapa_nomes = {
        str(k).strip().upper(): str(v).strip()
        for k, v in zip(df_mapa_nomes.iloc[:, 0], df_mapa_nomes.iloc[:, 1])
    }

    def traduzir(nome):
        return mapa_nomes.get(str(nome).strip().upper(), str(nome).strip())

    # Aplicar mapeamento
    df_ind_uso["Assessor"] = df_ind_uso.iloc[:, 0].apply(traduzir)
    df_indicacoes["Assessor"] = df_indicacoes.iloc[:, 2].apply(traduzir)
    df_participacao["Assessor"] = df_participacao.iloc[:, 0].apply(traduzir)
    df_btg["Assessor"] = df_btg.iloc[:, 2].apply(traduzir)
    df_nnm["Assessor"] = df_nnm.iloc[:, 0].apply(traduzir)
    df_pace_nnm["Assessor"] = df_pace_nnm.iloc[:, 1].apply(traduzir)
    df_receita["Assessor"] = df_receita.iloc[:, 2].apply(traduzir)
    df_pace_receita["Assessor"] = df_pace_receita.iloc[:, 1].apply(traduzir)
    df_nps["Assessor"] = df_nps["Nome do assessor"].apply(traduzir)
    df_outras_receitas["Assessor"] = df_outras_receitas.iloc[:, 0].apply(traduzir)

    # Indicadores de esforço
    df_esforco = pd.DataFrame()
    df_esforco["Assessor"] = df_participacao["Assessor"]
    df_esforco["Pontuação TIME"] = df_participacao.iloc[:, 1].fillna(0)

    aderencia = df_ind_uso[["Assessor", df_ind_uso.columns[8]]].copy()
    aderencia.columns = ["Assessor", "Score"]
    aderencia["Pontuação ADERÊNCIA SF"] = aderencia["Score"] * 10
    df_esforco = df_esforco.merge(aderencia[["Assessor", "Pontuação ADERÊNCIA SF"]], on="Assessor", how="outer")

    leads = df_ind_uso[["Assessor", df_ind_uso.columns[4]]].copy()
    leads.columns = ["Assessor", "Total Leads"]
    leads["Pontuação LEADS SF"] = pd.cut(
        leads["Total Leads"], [-1, 2, 5, 8, 11, float("inf")], labels=[0, 50, 100, 200, 300]
    ).astype(float).fillna(0).astype(int)
    df_esforco = df_esforco.merge(leads[["Assessor", "Pontuação LEADS SF"]], on="Assessor", how="outer")

    opps = df_ind_uso[["Assessor", df_ind_uso.columns[5]]].copy()
    opps.columns = ["Assessor", "Total OPP"]
    opps["Pontuação OPP SF"] = pd.cut(
        opps["Total OPP"], [-1, 4, 9, 14, float("inf")], labels=[0, 50, 100, 200]
    ).astype(float).fillna(0).astype(int)
    df_esforco = df_esforco.merge(opps[["Assessor", "Pontuação OPP SF"]], on="Assessor", how="outer")

    indic = df_indicacoes.groupby("Assessor").size().reset_index(name="Total IND")
    indic["Pontuação IND SF"] = pd.cut(
        indic["Total IND"], [0, 3, 6, float("inf")], labels=[50, 100, 200], right=True
    ).astype(float).fillna(0).astype(int)
    df_esforco = df_esforco.merge(indic[["Assessor", "Pontuação IND SF"]], on="Assessor", how="outer")

    df_esforco = df_esforco.fillna(0)
    df_esforco["Total Esforço"] = df_esforco[
        ["Pontuação TIME", "Pontuação ADERÊNCIA SF", "Pontuação LEADS SF", "Pontuação OPP SF", "Pontuação IND SF"]
    ].sum(axis=1)

    st.success("✅ Indicadores de esforço calculados com sucesso!")

    # --- Indicadores Finalísticos e Cálculo do Ranking ---
    


    auc = df_btg.groupby("Assessor")["PL Total"].sum().reset_index()
    auc = auc.sort_values(by="PL Total", ascending=False).reset_index(drop=True)
    auc["Ranking"] = auc.index + 1
    auc["Pontuação AUC TOTAL"] = auc["Ranking"].apply(lambda x: max(100 - (x - 1) * 10, 0) if x <= 10 else 0)
    auc = auc[["Assessor", "Pontuação AUC TOTAL"]]

    mes_coluna = 4 + data_inicio.month
    objetivo_nnm = df_pace_nnm.iloc[:, mes_coluna]
    df_nnm["NNM"] = pd.to_numeric(df_nnm.iloc[:, 1], errors="coerce")
    df_nnm = df_nnm.merge(df_pace_nnm[["Assessor", df_pace_nnm.columns[mes_coluna]]], on="Assessor", how="left")
    df_nnm["Obj"] = pd.to_numeric(df_nnm[df_pace_nnm.columns[mes_coluna]], errors="coerce")
    df_nnm["%"] = df_nnm["NNM"] / (df_nnm["Obj"] / 4) * 100
    def pontos_meta(p):
        if p < 0: return -100
        elif p <= 50: return -50
        elif p <= 75: return 0
        elif p <= 100: return 50
        elif p <= 125: return 100
        elif p <= 150: return 200
        else: return 300
    df_nnm["Pontuação NNM"] = df_nnm["%"].apply(pontos_meta)

    df_receita["Receita"] = pd.to_numeric(df_receita.iloc[:, 14], errors="coerce")
    receita_total = df_receita.groupby("Assessor")["Receita"].sum().reset_index()
    outras = df_outras_receitas.groupby("Assessor")[df_outras_receitas.columns[1]].sum().reset_index().rename(columns={df_outras_receitas.columns[1]: "Extra"})
    receita_total = receita_total.merge(outras, on="Assessor", how="outer").fillna(0)
    receita_total["Receita Total"] = receita_total["Receita"] + receita_total["Extra"]
    receita_total = receita_total.merge(df_pace_receita[["Assessor", df_pace_receita.columns[mes_coluna]]], on="Assessor", how="left")
    receita_total["Obj"] = pd.to_numeric(receita_total[df_pace_receita.columns[mes_coluna]], errors="coerce")
    receita_total["%"] = receita_total["Receita Total"] / (receita_total["Obj"] / 4) * 100
    receita_total["Pontuação RECEITA"] = receita_total["%"].apply(pontos_meta)

    df_nps["Nota"] = pd.to_numeric(df_nps["De 0 a 10, qual a probabilidade de você recomendar a assessoria de investimentos da  para um amigo ou familiar?"], errors="coerce")
    nps = df_nps.dropna(subset=["Nota", "Assessor"])
    def classificar(n):
        if n >= 9: return "Promotor"
        elif n >= 7: return "Neutro"
        else: return "Detrator"
    nps["Tipo"] = nps["Nota"].apply(classificar)
    nps_agg = nps.groupby("Assessor").agg(Respostas=("Nota", "count"), Promotores=("Tipo", lambda x: (x == "Promotor").sum()), Detratores=("Tipo", lambda x: (x == "Detrator").sum())).reset_index()
    nps_agg["NPS"] = ((nps_agg["Promotores"] / nps_agg["Respostas"] - nps_agg["Detratores"] / nps_agg["Respostas"]) * 100)
    def pontos_nps(row):
        if 1 <= row["Respostas"] <= 5:
            return 50 if row["NPS"] >= 80 else 0
        elif row["Respostas"] > 5:
            if row["NPS"] >= 95: return 150
            elif row["NPS"] >= 80: return 100
        return 0
    nps_agg["Pontuação NPS"] = nps_agg.apply(pontos_nps, axis=1)

    final = ativacoes.merge(auc, on="Assessor", how="outer")
    final = final.merge(df_nnm[["Assessor", "Pontuação NNM"]], on="Assessor", how="outer")
    final = final.merge(receita_total[["Assessor", "Pontuação RECEITA"]], on="Assessor", how="outer")
    final = final.merge(nps_agg[["Assessor", "Pontuação NPS"]], on="Assessor", how="outer")
    final = final.fillna(0)
    final["Total Finalísticos"] = final[["Pontuação ATIVAÇÕES", "Pontuação AUC TOTAL", "Pontuação NNM", "Pontuação RECEITA", "Pontuação NPS"]].sum(axis=1)

    df_geral = df_esforco.merge(final, on="Assessor", how="outer").fillna(0)
    df_geral["Pontuação Final"] = df_geral["Total Esforço"] + df_geral["Total Finalísticos"]
    df_geral["LEADS + OPP"] = df_geral["Pontuação LEADS SF"] + df_geral["Pontuação OPP SF"]
    df_geral["Elegível TOP3"] = (df_geral["Pontuação Final"] >= 400) & (df_geral["LEADS + OPP"] > 0) & (df_geral["Total Finalísticos"] > 0)
    df_geral = df_geral.sort_values(by="Pontuação Final", ascending=False).reset_index(drop=True)

    # --- Filtro de nomes inválidos ---
    df_geral = df_geral[~df_geral["Assessor"].str.lower().isin(["nan", "total", "assessor", "filtros aplicados..."])]

    # --- Gerar Ranking Geral completo ---
    df_geral["Ranking"] = df_geral["Pontuação Final"].rank(ascending=False, method="min").astype(int)

    colunas_ranking = [
        "Ranking", "Assessor", "Pontuação TIME", "Pontuação ADERÊNCIA SF", "Pontuação LEADS SF",
        "Pontuação OPP SF", "Pontuação IND SF", "Total Esforço",
        "Pontuação ATIVAÇÕES", "Pontuação AUC TOTAL", "Pontuação NNM",
        "Pontuação RECEITA", "Pontuação NPS", "Total Finalísticos",
        "Pontuação Final", "LEADS + OPP", "Elegível TOP3"
    ]

    df_ranking_final = df_geral[colunas_ranking].sort_values(by="Pontuação Final", ascending=False)

    # Função para aplicar destaque nas colunas desejadas
    def destacar_colunas(df):
        def estilo(col):
            if col.name == "Total Esforço":
                return ["background-color: #e6f4ea; color: black"] * len(col)
            elif col.name == "Total Finalísticos":
                return ["background-color: #e6f4ea; color: black"] * len(col)
            elif col.name == "Pontuação Final":
                return ["background-color: #f8f9fb; color: black; font-weight: bold"] * len(col)
            else:
                return [""] * len(col)
         #aplica a formatação apenas em colunas numéricas (int ou float)
        formatacoes = {
            col: (lambda x: f"{int(x):,}".replace(",", "."))
            for col in df.select_dtypes(include=["int", "float"]).columns
        }

        return df.style.apply(estilo).format(formatacoes)

    # Participaram da reunião
    participaram_reuniao_df = df_esforco[df_esforco["Pontuação TIME"] == 50][["Assessor"]]

    # Aderência SF
    aderencia_df = df_esforco[df_esforco["Pontuação ADERÊNCIA SF"] > 0][["Assessor", "Pontuação ADERÊNCIA SF"]]

    # Leads
    leads_df = df_ind_uso[["Assessor", "#Total de Leads"]].copy()
    leads_df.columns = ["Assessor", "Total Leads"]
    leads_df = leads_df[leads_df["Total Leads"] > 0]

    # Oportunidades
    opps_df = df_ind_uso[["Assessor", "#Total de Oportunidades"]].copy()
    opps_df.columns = ["Assessor", "Total Oportunidades"]
    opps_df = opps_df[opps_df["Total Oportunidades"] > 0]

    # Indicações
    indicacoes_df = df_indicacoes["Assessor"].value_counts().reset_index()
    indicacoes_df.columns = ["Assessor", "Total Indicações"]

    # Ativações por faixa
    ativacoes_detalhe = ativ.copy()
    ativacoes_detalhe["Faixa"] = ativacoes_detalhe["PL Total"].apply(
        lambda x: "300+" if x >= 300_000 and x < 1_000_000 else
                "1mm+" if x >= 1_000_000 and x < 5_000_000 else
                "5mm+" if x >= 5_000_000 else None
    )
    ativacoes_pivot = ativacoes_detalhe.dropna(subset=["Faixa"])
    ativacoes_pivot = ativacoes_pivot.pivot_table(index="Assessor", columns="Faixa", aggfunc="size", fill_value=0).reset_index()

    # NNM
    nnm_resumo = df_nnm[["Assessor", "NNM"]].dropna()

    # Receita
    receita_ranking = receita_total[["Assessor", "Pontuação RECEITA"]]
    receita_ranking = receita_ranking[receita_ranking["Pontuação RECEITA"] > 0]

    # NPS
    nps_resumo = nps_agg[nps_agg["Pontuação NPS"] > 0][["Assessor", "NPS"]]

    # Números Gerais Perspective
    soma_nnm = df_nnm["NNM"].sum()
    soma_receita = receita_total["Receita Total"].sum()


    st.subheader("🏅 Destaques da Semana")
    with st.expander("🏭 Números Gerais Perspective"):
        col1, col2 = st.columns(2)
        col1.metric("💵 Soma do NNM", f"R$ {soma_nnm:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        col2.metric("💰 Soma da Receita", f"R$ {soma_receita:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))


    with st.expander("👥 Participaram da Reunião"):
        st.dataframe(participaram_reuniao_df, use_container_width=True)

    with st.expander("📊 Aderência SalesForce"):
        st.dataframe(aderencia_df, use_container_width=True)

    with st.expander("📬 Leads Criadas"):
        st.dataframe(leads_df, use_container_width=True)

    with st.expander("📈 Oportunidades Criadas"):
        st.dataframe(opps_df, use_container_width=True)

    with st.expander("📨 Indicações Recebidas"):
        st.dataframe(indicacoes_df, use_container_width=True)

    with st.expander("🚀 Ativações por Faixa"):
        st.dataframe(ativacoes_pivot, use_container_width=True)

    with st.expander("💵 NNM da Semana"):
        df_formatado = nnm_resumo.copy()
        df_formatado["NNM"] = df_formatado["NNM"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.dataframe(df_formatado, use_container_width=True)

    with st.expander("💰 Receita - Pontuação"):
        st.dataframe(receita_ranking, use_container_width=True)

    with st.expander("🌟 NPS"):
        st.dataframe(nps_resumo, use_container_width=True)


    st.subheader("🥇 TOP 3 da Semana")
    # Filtrar apenas os elegíveis, ordenar por pontuação
    top3 = df_geral[df_geral["Elegível TOP3"]].sort_values(by="Pontuação Final", ascending=False).reset_index(drop=True)
    # Lista dos rótulos corretos
    rotulos = ["🥇 1º Lugar", "🥈 2º Lugar", "🥉 3º Lugar"]
    # Mostrar apenas se houver alguém para aquela posição
    for posicao in range(3):
        if posicao < len(top3):
            row = top3.iloc[posicao]
            with st.expander(rotulos[posicao]):
                st.markdown(f"**{row['Assessor']}** — {row['Pontuação Final']:.0f} pontos")

    st.subheader("📋 Ranking Geral")
    with st.expander("Ver Tabela Completa", expanded=False):
        st.dataframe(destacar_colunas(df_ranking_final), use_container_width=True)

