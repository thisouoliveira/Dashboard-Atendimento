import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Função para formatar números no padrão brasileiro (pt-BR)


def format_brl(numero, casas_decimais=0):
    try:
        if casas_decimais == 0:
            return f"{numero:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return f"{numero:,.{casas_decimais}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(numero)


# Configuração da página
st.set_page_config(
    page_title="Dashboard de Atendimento por Setor",
    page_icon="📊",
    layout="wide"
)

# Função para converter tempo médio de espera para minutos


def convert_time_to_minutes(time_str):
    """Converte string de tempo HH:MM para minutos"""
    try:
        # Lida com formatos como '0:06' ou '00:06'
        if isinstance(time_str, str):
            parts = time_str.split(":")
            if len(parts) == 2:
                hours = int(parts[0])
                minutes = int(parts[1])
                return hours * 60 + minutes
            else:
                return time_str.minute  # Formato inválido
        return time_str.minute  # Não é string
    except ValueError:
        return 0  # Erro de conversão


# Carregar dados do Excel


@st.cache_data
def load_data_from_excel(file_path):
    xls = pd.ExcelFile(file_path)
    all_sheets_df = []
    for sheet_name in xls.sheet_names:
        df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
        all_sheets_df.append(df_sheet)

    df = pd.concat(all_sheets_df, ignore_index=True)

    # Converter tempo médio de espera para minutos
    df["TEMPO_MINUTOS"] = df["TEMPO_MEDIO_ESPERA"].apply(
        convert_time_to_minutes)
    return df

# Carregar dados de atendimentos


@st.cache_data
def load_atendimentos_data(file_path):
    """Carrega dados de todas as abas da planilha de atendimentos"""
    try:
        xls = pd.ExcelFile(file_path)
        df_list = []

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # Limpeza e tipos de dados
            df = df.dropna(subset=['MES', 'ANO', 'TIPO'])
            df['MES'] = df['MES'].astype(int)
            df['ANO'] = df['ANO'].astype(int)
            df['QTDE'] = pd.to_numeric(
                df['QTDE'], errors='coerce').fillna(0).astype(int)

            df_list.append(df)

        if df_list:
            return pd.concat(df_list, ignore_index=True)
        else:
            return pd.DataFrame()

    except Exception as e:
        st.error(f"Erro ao carregar dados de atendimentos: {e}")
        return pd.DataFrame()


# Título principal
st.title("📊 Dashboard de Atendimento por Setor")
st.markdown("---")

# Criar abas
tab1, tab2 = st.tabs(
    ["⏰ Tempo de Atendimento", "📊 Quantidade de Atendimentos"])

# === ABA 1: TEMPO DE ATENDIMENTO ===
with tab1:
    # Carregar dados de tempo
    try:
        df = load_data_from_excel("Tempo_Atendimentos.xlsx")

        # Sidebar para filtros
        st.sidebar.header("🔍 Filtros - Tempo de Atendimento")

        # Filtro por setor
        setores_disponiveis = ["Todos"] + sorted(df["SETOR"].unique().tolist())
        setor_selecionado = st.sidebar.selectbox(
            "Selecione o Setor:", setores_disponiveis, key="setor_tempo")

        # Filtro por ano
        anos_disponiveis = ["Todos"] + sorted(df["ANO"].unique().tolist())
        ano_selecionado = st.sidebar.selectbox(
            "Selecione o Ano:", anos_disponiveis, key="ano_tempo")

        # Filtro por mês
        meses_disponiveis = ["Todos"] + sorted(df["MES"].unique().tolist())
        mes_selecionado = st.sidebar.selectbox(
            "Selecione o Mês:", meses_disponiveis, key="mes_tempo")

        # Aplicar filtros
        df_filtrado = df.copy()

        if setor_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["SETOR"]
                                      == setor_selecionado]

        if ano_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["ANO"] == ano_selecionado]

        if mes_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["MES"] == mes_selecionado]

        # Métricas principais
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            total_atendimentos = df_filtrado["QUANTIDADE"].sum()
            st.metric("Total de Atendimentos",
                      f"{format_brl(total_atendimentos)}")

        with col2:
            tempo_medio_geral = df_filtrado["TEMPO_MINUTOS"].mean()
            st.metric("Tempo Médio de Espera", f"{tempo_medio_geral:.1f} min")

        with col3:
            setores_unicos = df_filtrado["SETOR"].nunique()
            st.metric("Setores Ativos", setores_unicos)

        with col4:
            maior_fila = df_filtrado["QUANTIDADE"].max()
            st.metric("Maior Fila", f"{format_brl(maior_fila)}")

        st.markdown("---")

        # Gráficos principais
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📈 Quantidade de Atendimentos por Setor")

            # Agrupar por setor
            df_setor = df_filtrado.groupby("SETOR").agg({
                "QUANTIDADE": "sum"
            }).reset_index().sort_values("QUANTIDADE", ascending=True)

            fig_setor = px.bar(
                df_setor,
                x="QUANTIDADE",
                y="SETOR",
                orientation="h",
                title="Atendimentos por Setor",
                color="QUANTIDADE",
                color_continuous_scale="Blues"
            )
            fig_setor.update_layout(height=400)
            st.plotly_chart(fig_setor, use_container_width=True)

        with col2:
            st.subheader("⏰ Tempo Médio de Espera por Setor")

            # Agrupar por setor para tempo médio
            df_tempo_setor = df_filtrado.groupby("SETOR").agg({
                "TEMPO_MINUTOS": "mean"
            }).reset_index().sort_values("TEMPO_MINUTOS", ascending=True)

            fig_tempo = px.bar(
                df_tempo_setor,
                x="TEMPO_MINUTOS",
                y="SETOR",
                orientation="h",
                title="Tempo Médio de Espera por Setor (minutos)",
                color="TEMPO_MINUTOS",
                color_continuous_scale="Reds"
            )
            fig_tempo.update_layout(height=400)
            st.plotly_chart(fig_tempo, use_container_width=True)

        # Gráficos de evolução temporal
        st.subheader("📅 Evolução Temporal")

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Atendimentos por Mês")

            # Criar coluna de data para melhor visualização
            df_filtrado["DATA"] = df_filtrado["ANO"].astype(
                str) + "-" + df_filtrado["MES"].astype(str).str.zfill(2)

            df_mensal = df_filtrado.groupby(["DATA", "SETOR"]).agg({
                "QUANTIDADE": "sum"
            }).reset_index()

            fig_mensal = px.line(
                df_mensal,
                x="DATA",
                y="QUANTIDADE",
                color="SETOR",
                title="Evolução Mensal de Atendimentos",
                markers=True
            )
            fig_mensal.update_layout(height=400)
            st.plotly_chart(fig_mensal, use_container_width=True)

        with col2:
            st.subheader("Tempo de Espera por Mês")

            df_tempo_mensal = df_filtrado.groupby(["DATA", "SETOR"]).agg({
                "TEMPO_MINUTOS": "mean"
            }).reset_index()

            fig_tempo_mensal = px.line(
                df_tempo_mensal,
                x="DATA",
                y="TEMPO_MINUTOS",
                color="SETOR",
                title="Evolução Mensal do Tempo de Espera",
                markers=True
            )
            fig_tempo_mensal.update_layout(height=400)
            st.plotly_chart(fig_tempo_mensal, use_container_width=True)

        # Gráfico de correlação
        st.subheader("🔗 Correlação: Quantidade vs Tempo de Espera")

        fig_scatter = px.scatter(
            df_filtrado,
            x="QUANTIDADE",
            y="TEMPO_MINUTOS",
            color="SETOR",
            size="QUANTIDADE",
            hover_data=["ANO", "MES"],
            title="Relação entre Quantidade de Atendimentos e Tempo de Espera"
        )
        fig_scatter.update_layout(height=500)
        st.plotly_chart(fig_scatter, use_container_width=True)

        # Tabela de dados
        st.subheader("📋 Dados Detalhados")

        # Preparar dados para exibição
        df_display = df_filtrado.copy()
        df_display["TEMPO_MEDIO_ESPERA_MIN"] = df_display["TEMPO_MINUTOS"].apply(
            lambda x: f"{int(x//60):02d}:{int(x%60):02d}")

        # Selecionar colunas para exibição
        colunas_display = ["CODIGO", "SETOR", "ANO", "MES",
                           "QUANTIDADE", "TEMPO_MEDIO_ESPERA", "TEMPO_MEDIO_ESPERA_MIN"]
        df_display = df_display[colunas_display]

        # Formatação nas colunas dos dataframes
        if "QUANTIDADE" in df_display.columns:
            df_display["QUANTIDADE"] = df_display["QUANTIDADE"].apply(
                lambda x: format_brl(x))
            df_display["ANO"] = df_display["ANO"].astype(str)

        # Formatação nas colunas dos dataframes
        if "ANO" in df_filtrado.columns:
            df_filtrado["ANO"] = df_filtrado["ANO"].astype(str)

        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True
        )

        # Estatísticas resumidas
        st.subheader("📊 Estatísticas Resumidas")

        col1, col2 = st.columns(2)

        with col1:
            st.write("**Por Setor:**")
            stats_setor = df_filtrado.groupby("SETOR").agg({
                "QUANTIDADE": ["sum", "mean", "max"],
                "TEMPO_MINUTOS": ["mean", "max", "min"]
            }).round(2)

            # Renomear colunas para melhor legibilidade
            stats_setor.columns = [
                "Total Atend.", "Média Atend.", "Máx Atend.",
                "Tempo Médio (min)", "Tempo Máx (min)", "Tempo Mín (min)"
            ]

            # Formatar os números para padrão brasileiro nas colunas numéricas
            for col in stats_setor.columns:
                stats_setor[col] = stats_setor[col].apply(
                    lambda x: format_brl(x, 0))

            st.dataframe(stats_setor, use_container_width=True)

        with col2:
            st.write("**Por Mês:**")
            stats_mes = df_filtrado.groupby(["ANO", "MES"]).agg({
                "QUANTIDADE": ["sum", "mean"],
                "TEMPO_MINUTOS": ["mean"]
            }).round(2)

            stats_mes.columns = ["Total Atend.",
                                 "Média Atend.", "Tempo Médio (min)"]

            # Formatar os números para padrão brasileiro nas colunas numéricas
            for col in stats_mes.columns:
                stats_mes[col] = stats_mes[col].apply(
                    lambda x: format_brl(x, 0))

            st.dataframe(stats_mes, use_container_width=True)

    except FileNotFoundError:
        st.error("❌ Arquivo 'Tempo_Atendimentos.xlsx' não encontrado.")

# === ABA 2: QUANTIDADE DE ATENDIMENTOS ===
with tab2:
    # Carregar dados de atendimentos
    df_atendimentos = load_atendimentos_data("Atendimentos.xlsx")

    if not df_atendimentos.empty:
        # Sidebar para filtros da aba de atendimentos
        st.sidebar.header("🔍 Filtros - Quantidade de Atendimentos")

        # Filtro por setor
        setores_atend = ["Todos"] + \
            sorted(df_atendimentos["SETOR"].unique().tolist())
        setor_atend = st.sidebar.selectbox(
            "Selecione o Setor:", setores_atend, key="setor_atend")

        # Filtro por tipo
        tipos_atend = ["Todos"] + \
            sorted(df_atendimentos["TIPO"].unique().tolist())
        tipo_atend = st.sidebar.selectbox(
            "Selecione o Tipo:", tipos_atend, key="tipo_atend")

        # Filtro por ano
        anos_atend = ["Todos"] + \
            sorted(df_atendimentos["ANO"].unique().tolist())
        ano_atend = st.sidebar.selectbox(
            "Selecione o Ano:", anos_atend, key="ano_atend")

        # Filtro por mês
        meses_atend = ["Todos"] + \
            sorted(df_atendimentos["MES"].unique().tolist())
        mes_atend = st.sidebar.selectbox(
            "Selecione o Mês:", meses_atend, key="mes_atend")

        # Aplicar filtros
        df_atend_filtrado = df_atendimentos.copy()

        if setor_atend != "Todos":
            df_atend_filtrado = df_atend_filtrado[df_atend_filtrado["SETOR"] == setor_atend]

        if tipo_atend != "Todos":
            df_atend_filtrado = df_atend_filtrado[df_atend_filtrado["TIPO"] == tipo_atend]

        if ano_atend != "Todos":
            df_atend_filtrado = df_atend_filtrado[df_atend_filtrado["ANO"] == ano_atend]

        if mes_atend != "Todos":
            df_atend_filtrado = df_atend_filtrado[df_atend_filtrado["MES"] == mes_atend]

        # Métricas principais
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            total_qtde = df_atend_filtrado["QTDE"].sum()
            st.metric("Total de Atendimentos", f"{format_brl(total_qtde)}")

        with col2:
            media_qtde = df_atend_filtrado["QTDE"].mean()
            st.metric("Média de Atendimentos", f"{format_brl(media_qtde)}")

        with col3:
            setores_ativos = df_atend_filtrado["SETOR"].nunique()
            st.metric("Setores Ativos", setores_ativos)

        with col4:
            max_qtde = df_atend_filtrado["QTDE"].max()
            st.metric("Maior Quantidade", f"{format_brl(max_qtde)}")

        st.markdown("---")

        # Gráficos principais
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📊 Quantidade por Setor")

            # Agrupar por setor
            df_setor_qtde = df_atend_filtrado.groupby("SETOR").agg({
                "QTDE": "sum"
            }).reset_index().sort_values("QTDE", ascending=True)

            fig_setor_qtde = px.bar(
                df_setor_qtde,
                x="QTDE",
                y="SETOR",
                orientation="h",
                title="Total de Atendimentos por Setor",
                color="QTDE",
                color_continuous_scale="Viridis"
            )
            fig_setor_qtde.update_layout(height=400)
            st.plotly_chart(fig_setor_qtde, use_container_width=True)

        with col2:
            st.subheader("🎯 Distribuição por Tipo")

            # Gráfico de pizza por tipo
            df_tipo = df_atend_filtrado.groupby("TIPO").agg({
                "QTDE": "sum"
            }).reset_index()

            fig_tipo = px.pie(
                df_tipo,
                values="QTDE",
                names="TIPO",
                title="Distribuição por Tipo de Atendimento"
            )
            fig_tipo.update_layout(height=400)
            st.plotly_chart(fig_tipo, use_container_width=True)

        # Gráficos de evolução temporal
        st.subheader("📅 Evolução Temporal dos Atendimentos")

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Evolução Mensal")

            # Criar coluna de data
            df_atend_filtrado["DATA"] = df_atend_filtrado["ANO"].astype(
                str) + "-" + df_atend_filtrado["MES"].astype(str).str.zfill(2)

            df_mensal_atend = df_atend_filtrado.groupby(["DATA", "TIPO"]).agg({
                "QTDE": "sum"
            }).reset_index()

            fig_mensal_atend = px.line(
                df_mensal_atend,
                x="DATA",
                y="QTDE",
                color="TIPO",
                title="Evolução Mensal por Tipo",
                markers=True
            )
            fig_mensal_atend.update_layout(height=400)
            st.plotly_chart(fig_mensal_atend, use_container_width=True)

        with col2:
            st.subheader("Comparação por Setor e Tipo")

            # Gráfico de barras agrupadas
            df_setor_tipo = df_atend_filtrado.groupby(["SETOR", "TIPO"]).agg({
                "QTDE": "sum"
            }).reset_index()

            fig_setor_tipo = px.bar(
                df_setor_tipo,
                x="SETOR",
                y="QTDE",
                color="TIPO",
                title="Atendimentos por Setor e Tipo",
                barmode="group"
            )
            fig_setor_tipo.update_layout(height=400, xaxis_tickangle=-45)
            st.plotly_chart(fig_setor_tipo, use_container_width=True)

        # Heatmap
        st.subheader("🔥 Mapa de Calor: Setor vs Mês")

        # Criar pivot table para heatmap
        df_heatmap = df_atend_filtrado.pivot_table(
            values="QTDE",
            index="SETOR",
            columns="MES",
            aggfunc="sum",
            fill_value=0
        )

        fig_heatmap = px.imshow(
            df_heatmap,
            title="Quantidade de Atendimentos por Setor e Mês",
            color_continuous_scale="Blues",
            aspect="auto"
        )
        fig_heatmap.update_layout(height=500)
        st.plotly_chart(fig_heatmap, use_container_width=True)

        # Tabela de dados
        st.subheader("📋 Dados Detalhados")

        # Preparar dados para exibição
        df_display_atend = df_atend_filtrado.copy()

        if "QTDE" in df_display_atend.columns:
            df_display_atend["QTDE"] = df_display_atend["QTDE"].apply(
                lambda x: format_brl(x))
            df_display_atend["ANO"] = df_display_atend["ANO"].astype(str)

        st.dataframe(
            df_display_atend,
            use_container_width=True,
            hide_index=True
        )

        # Estatísticas resumidas
        st.subheader("📊 Estatísticas Resumidas")

        col1, col2 = st.columns(2)

        with col1:
            st.write("**Por Setor:**")
            stats_setor_atend = df_atend_filtrado.groupby("SETOR").agg({
                "QTDE": ["sum", "mean", "max", "min"]
            }).round(2)

            stats_setor_atend.columns = ["Total", "Média", "Máximo", "Mínimo"]

            # Formatar os números para padrão brasileiro nas colunas numéricas
            for col in stats_setor_atend.columns:
                stats_setor_atend[col] = stats_setor_atend[col].apply(
                    lambda x: format_brl(x, 0))

            st.dataframe(stats_setor_atend, use_container_width=True)

        with col2:
            st.write("**Por Tipo:**")
            stats_tipo_atend = df_atend_filtrado.groupby("TIPO").agg({
                "QTDE": ["sum", "mean", "max", "min"]
            }).round(2)

            stats_tipo_atend.columns = ["Total", "Média", "Máximo", "Mínimo"]

            # Formatar os números para padrão brasileiro nas colunas numéricas
            for col in stats_tipo_atend.columns:
                stats_tipo_atend[col] = stats_tipo_atend[col].apply(
                    lambda x: format_brl(x, 0))

            st.dataframe(stats_tipo_atend, use_container_width=True)

    else:
        st.error("❌ Não foi possível carregar os dados de atendimentos.")

# Rodapé
st.markdown("---")
st.markdown("*Dashboard desenvolvido com Streamlit e Plotly*")
