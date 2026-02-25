import smartsheet
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import unicodedata

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard 5S", layout="wide")

# --- FUNÇÃO AUXILIAR PARA NORMALIZAÇÃO DE TEXTO ---
def normalizar(texto):
    """Remove acentos, espaços extras e coloca em maiúsculo para comparação segura."""
    if not texto: return ""
    nfkd_form = unicodedata.normalize('NFKD', str(texto))
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).upper().strip()

# --- 1. SISTEMA DE LOGIN VIA SECRETS ---
def check_password():
    if "password_correct" not in st.session_state:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.write("#Acesso Restrito")
            with st.form("login_form"):
                user_input = st.text_input("Usuário ou E-mail")
                password_input = st.text_input("Senha", type="password")
                submit_button = st.form_submit_button("Entrar")
                
                if submit_button:
                    try:
                        usuarios_validos = st.secrets["usuarios"]
                        if user_input in usuarios_validos and str(usuarios_validos[user_input]) == password_input:
                            st.session_state["password_correct"] = True
                            st.rerun()
                        else:
                            st.error("Usuário ou senha incorretos.")
                    except KeyError:
                        st.error("Configuração de 'usuarios' não encontrada nos Secrets.")
        return False
    return True

# --- EXECUÇÃO DO DASHBOARD ---
if check_password():

    if st.sidebar.button("Sair / Logout"):
        del st.session_state["password_correct"]
        st.rerun()

    # --- SECRETS ---
    try:
        TOKEN = st.secrets["SMARTSHEET_TOKEN"]
        ID_PLANILHA = st.secrets["ID_PLANILHA"]
    except KeyError as e:
        st.error(f"Erro: A chave {e} não foi configurada nos Secrets.")
        st.stop()

    # --- CARREGAMENTO DE DADOS ---
    @st.cache_data(ttl=600)
    def carregar_dados():
        smart = smartsheet.Smartsheet(TOKEN)
        sheet = smart.Sheets.get_sheet(ID_PLANILHA)
        
        cols = [column.title for column in sheet.columns]
        data = [[cell.value for cell in row.cells] for row in sheet.rows]
        df = pd.DataFrame(data, columns=cols)

        if df.empty:
            raise ValueError("A planilha está vazia.")

        # Nomes amigáveis para os cards e radar
        sensos = ['Descarte', 'Organização', 'Limpeza', 'Saúde e Higiene', 'Autodisciplina']
        
        # Identifica colunas de Não Conformidade (NC)
        col_ncs_lista = [c for c in df.columns if 'NC' in c.upper()]

        # Garantir colunas essenciais
        if 'Setor' not in df.columns:
            df['Setor'] = 'Indefinido'
        if 'Data' not in df.columns:
            raise ValueError("Coluna 'Data' não encontrada na planilha.")
        if 'Nota' not in df.columns:
            df['Nota'] = 0

        # --- TRATAMENTO ---
        df = df.dropna(how='all')
        df['Setor'] = df['Setor'].fillna('Indefinido').astype(str)
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
        df = df.dropna(subset=['Data'])
        df['Mes_Ano_Ref'] = df['Data'].dt.strftime('%m/%Y')

        # Limpeza de valores numéricos (converte 80% ou 0,8 para 80.0)
        colunas_valores = sensos + ['Nota']
        for col in colunas_valores:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.replace('%', '', regex=False)
                    .str.replace(',', '.', regex=False)
                )
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                if df[col].mean() <= 1.0 and df[col].mean() > 0:
                    df[col] = df[col] * 100

        for col in col_ncs_lista:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df = df.sort_values('Data')
        return df, sensos, col_ncs_lista

    try:
        df, sensos, col_ncs_lista = carregar_dados()

        # --- SIDEBAR ---
        st.sidebar.header("Filtros de Auditoria")
        setor_sel = st.sidebar.selectbox("Selecione o Setor", ['Todos'] + sorted(df['Setor'].unique()))
        mes_sel = st.sidebar.selectbox("Selecione o Mês/Ano", ['Todos'] + sorted(df['Mes_Ano_Ref'].unique(), reverse=True))

        df_plot = df.copy()
        if setor_sel != 'Todos':
            df_plot = df_plot[df_plot['Setor'] == setor_sel]
        if mes_sel != 'Todos':
            df_plot = df_plot[df_plot['Mes_Ano_Ref'] == mes_sel]

        if df_plot.empty:
            st.warning("Nenhum dado encontrado para os filtros selecionados.")
            st.stop()

        # --- MÉTRICAS GERAIS ---
        st.title("Dashboard de Auditoria 5S")
        nota_media = df_plot['Nota'].mean()
        total_ncs = df_plot[col_ncs_lista].sum().sum()

        m1, m2, m3 = st.columns(3)
        m1.metric("Nota Média (Geral)", f"{nota_media:.1f}%")
        m2.metric("Total de NCs", int(total_ncs))
        m3.metric("Qtd Auditorias", len(df_plot))
        st.divider()

        # --- CARDS POR SENSO (CORREÇÃO DA BUSCA DE NC) ---
        cols_sensos = st.columns(5)
        for i, s in enumerate(sensos):
            # Média da nota do senso
            media_s = df_plot[s].mean() if s in df_plot.columns else 0
            
            # Busca de NC usando normalização (resolve o problema de acentos)
            s_norm = normalizar(s)
            # Filtra colunas de NC que contenham o nome do senso (ex: "SAUDE" dentro de "QTD NC SAUDE E HIGIENE")
            col_nc_esp = [c for c in col_ncs_lista if s_norm in normalizar(c)]
            total_nc_s = df_plot[col_nc_esp].sum().sum() if col_nc_esp else 0

            cols_sensos[i].metric(
                label=s,
                value=f"{media_s:.1f}%",
                delta=f"{int(total_nc_s)} NCs",
                delta_color="inverse"
            )

        # --- EVOLUÇÃO MENSAL ---
        st.subheader("Evolução Mensal e Variação")
        df_hist = df if setor_sel == 'Todos' else df[df['Setor'] == setor_sel]
        df_mensal = df_hist.set_index('Data').resample('ME').agg({'Nota': 'mean'}).reset_index()

        if not df_mensal.empty:
            df_mensal['Mes_Ano'] = df_mensal['Data'].dt.strftime('%m/%Y')
            df_mensal['Variacao'] = df_mensal['Nota'].diff().fillna(0)
            fig_evol = go.Figure()
            fig_evol.add_bar(x=df_mensal['Mes_Ano'], y=df_mensal['Nota'], name='Nota (%)', text=df_mensal['Nota'].round(1), textposition='auto')
            fig_evol.add_scatter(x=df_mensal['Mes_Ano'], y=df_mensal['Variacao'], name='Dif (p.p.)', mode='lines+markers+text', text=df_mensal['Variacao'].round(1), textposition='top center', yaxis='y2')
            fig_evol.update_layout(yaxis=dict(title="Nota %", range=[0, 115]), yaxis2=dict(title="Variação", overlaying='y', side='right', showgrid=False), height=400, margin=dict(t=20, b=20))
            st.plotly_chart(fig_evol, use_container_width=True)

        # --- RADAR E GRÁFICO DE NCs ---
        c1, c2 = st.columns(2)
        with c1:
            valores_radar = [df_plot[s].mean() if s in df_plot.columns else 0 for s in sensos]
            if any(valores_radar):
                fig_rad = go.Figure(go.Scatterpolar(r=valores_radar + [valores_radar[0]], theta=sensos + [sensos[0]], fill='toself'))
                fig_rad.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), title="Radar de Performance (%)", height=400)
                st.plotly_chart(fig_rad, use_container_width=True)

        with c2:
            if col_ncs_lista:
                df_nc_bar = df_plot[col_ncs_lista].sum().reset_index()
                df_nc_bar.columns = ['Categoria', 'Total']
                # Limpa os nomes das categorias para o gráfico ficar mais bonito
                df_nc_bar['Categoria'] = df_nc_bar['Categoria'].str.replace('QTD NC ', '', case=False)
                fig_nc = px.bar(df_nc_bar, x='Categoria', y='Total', text_auto=True, title="Total de NCs por Categoria")
                fig_nc.update_layout(height=400)
                st.plotly_chart(fig_nc, use_container_width=True)

        # --- TABELA ---
        st.markdown("---")
        st.subheader("Detalhamento das Auditorias")
        with st.expander("Clique para visualizar os dados brutos"):
            df_display = df_plot.copy()
            df_display['Data'] = df_display['Data'].dt.strftime('%d/%m/%Y')
            st.dataframe(df_display, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")

    st.sidebar.markdown("---")

    st.sidebar.caption("Dashboard v3.1 | Schumann")
