import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Chroma Blindagens | Dashboard",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── ESTILO CHROMA ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Inter:wght@300;400;500&display=swap');

:root {
    --ciano: #00E5E5;
    --ciano-escuro: #00B8B8;
    --preto: #0A0A0F;
    --card: #111118;
    --card2: #1A1A24;
    --borda: #00E5E520;
    --texto: #E0E0E0;
    --texto2: #888;
}

* { font-family: 'Inter', sans-serif; }
h1, h2, h3, .titulo { font-family: 'Rajdhani', sans-serif !important; }

/* Fundo geral */
.stApp { background: var(--preto) !important; }
section[data-testid="stSidebar"] { background: var(--card) !important; border-right: 1px solid var(--borda); }

/* Remove padding padrão */
.block-container { padding: 1.5rem 2rem !important; }

/* Métricas */
[data-testid="metric-container"] {
    background: var(--card) !important;
    border: 1px solid var(--borda) !important;
    border-radius: 8px !important;
    padding: 1rem !important;
}
[data-testid="metric-container"] label { color: var(--texto2) !important; font-size: 0.75rem !important; letter-spacing: 2px; text-transform: uppercase; }
[data-testid="metric-container"] [data-testid="stMetricValue"] { color: var(--ciano) !important; font-family: 'Rajdhani', sans-serif !important; font-size: 2rem !important; font-weight: 700 !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { background: var(--card) !important; border-radius: 8px; gap: 4px; padding: 4px; }
.stTabs [data-baseweb="tab"] { background: transparent !important; color: var(--texto2) !important; border-radius: 6px !important; font-family: 'Rajdhani', sans-serif !important; font-size: 1rem !important; font-weight: 600 !important; letter-spacing: 1px; }
.stTabs [aria-selected="true"] { background: var(--ciano) !important; color: var(--preto) !important; }

/* Tabelas */
.stDataFrame { border: 1px solid var(--borda) !important; border-radius: 8px !important; }
thead th { background: var(--card2) !important; color: var(--ciano) !important; font-family: 'Rajdhani', sans-serif !important; font-size: 0.85rem !important; letter-spacing: 1px; text-transform: uppercase; }

/* Selectbox e inputs */
.stSelectbox > div > div { background: var(--card) !important; border: 1px solid var(--borda) !important; color: var(--texto) !important; }
.stMultiSelect > div > div { background: var(--card) !important; border: 1px solid var(--borda) !important; }

/* Scrollbar */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: var(--preto); }
::-webkit-scrollbar-thumb { background: var(--ciano-escuro); border-radius: 2px; }

/* Sidebar texto */
.css-1d391kg, [data-testid="stSidebar"] * { color: var(--texto) !important; }
</style>
""", unsafe_allow_html=True)

# ── CORES PLOTLY ──────────────────────────────────────────────────────────────
CORES = ["#00E5E5", "#00B8B8", "#008080", "#006666", "#004444",
         "#00FFCC", "#00CC99", "#009977", "#007755"]
LAYOUT_PLOTLY = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter", color="#E0E0E0", size=12),
    title_font=dict(family="Rajdhani", color="#00E5E5", size=18),
    legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor="rgba(0,229,229,0.12)"),
    xaxis=dict(gridcolor="rgba(255,255,255,0.03)", linecolor="rgba(255,255,255,0.06)"),
    yaxis=dict(gridcolor="rgba(255,255,255,0.03)", linecolor="rgba(255,255,255,0.06)"),
    margin=dict(l=20, r=20, t=40, b=20),
)


# ── CARREGAR DADOS ────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def carregar_dados(caminho):
    try:
        xls = pd.ExcelFile(caminho)
        dados = {}
        for aba in xls.sheet_names:
            df = pd.read_excel(caminho, sheet_name=aba)
            dados[aba] = df
        return dados, os.path.getmtime(caminho)
    except Exception as e:
        return None, None


def encontrar_excel():
    """Procura o Excel mais recente na pasta de relatórios."""
    pasta = r"C:\Users\gustavo\Downloads\Chroma_Relatorios"
    if not os.path.exists(pasta):
        return None
    arquivos = [f for f in os.listdir(pasta) if f.startswith("chroma_relatorio_completo") and f.endswith(".xlsx")]
    if not arquivos:
        return None
    return os.path.join(pasta, sorted(arquivos)[-1])


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 1rem 0 1.5rem'>
        <div style='font-family: Rajdhani; font-size: 2rem; font-weight: 700; color: #00E5E5; letter-spacing: 4px;'>CHROMA</div>
        <div style='font-size: 0.7rem; color: #888; letter-spacing: 6px; margin-top: -6px;'>BLINDAGENS</div>
        <div style='margin-top: 0.5rem; height: 2px; background: linear-gradient(90deg, transparent, #00E5E5, transparent);'></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📂 Fonte de Dados")

    caminho_manual = st.text_input(
        "Caminho do arquivo Excel",
        placeholder=r"C:\Users\...\chroma_relatorio.xlsx",
        help="Cole o caminho completo do arquivo Excel gerado pelo script"
    )

    excel_auto = encontrar_excel()
    caminho = caminho_manual if caminho_manual else excel_auto

    if caminho and os.path.exists(caminho):
        ts = datetime.fromtimestamp(os.path.getmtime(caminho))
        st.success(f"✓ Arquivo encontrado")
        st.caption(f"Atualizado: {ts.strftime('%d/%m/%Y %H:%M')}")
    elif not caminho:
        st.info("Nenhum arquivo encontrado. Cole o caminho acima.")

    st.markdown("---")
    st.markdown("### 🔄 Atualização")
    if st.button("↺  Recarregar dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.caption("Chroma Blindagens © 2026")


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='display:flex; align-items:center; gap:1rem; margin-bottom:1.5rem;'>
    <div style='flex:1'>
        <div style='font-family:Rajdhani; font-size:2rem; font-weight:700; color:#00E5E5; letter-spacing:3px;'>DASHBOARD OPERACIONAL</div>
        <div style='color:#888; font-size:0.85rem; letter-spacing:2px;'>CHROMA BLINDAGENS — VISÃO GERAL</div>
    </div>
</div>
<div style='height:1px; background:linear-gradient(90deg, #00E5E5, transparent); margin-bottom:1.5rem;'></div>
""", unsafe_allow_html=True)

# ── CONTEÚDO PRINCIPAL ────────────────────────────────────────────────────────
if not caminho or not os.path.exists(caminho):
    st.markdown("""
    <div style='text-align:center; padding:4rem; border:1px dashed #00E5E520; border-radius:12px;'>
        <div style='font-size:3rem;'>📊</div>
        <div style='font-family:Rajdhani; font-size:1.5rem; color:#00E5E5; margin:1rem 0;'>NENHUM ARQUIVO CARREGADO</div>
        <div style='color:#888;'>Execute o script <code>chroma_exportar.py</code> para gerar o Excel<br>ou informe o caminho do arquivo na barra lateral.</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

dados, _ = carregar_dados(caminho)

if not dados:
    st.error("Erro ao carregar o arquivo. Verifique se é um Excel válido.")
    st.stop()

abas = list(dados.keys())

# KPIs globais
total_registros = sum(len(df) for df in dados.values())
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total de Registros", f"{total_registros:,}".replace(",", "."))
with col2:
    st.metric("Relatórios", len(abas))
with col3:
    if "Contratos" in dados:
        st.metric("Contratos", f"{len(dados['Contratos']):,}".replace(",", "."))
with col4:
    if "Producao" in dados:
        st.metric("Em Produção", f"{len(dados['Producao']):,}".replace(",", "."))

st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)

# ── ABAS ──────────────────────────────────────────────────────────────────────
tabs = st.tabs([f"🛡️ {a}" for a in abas])

def tab_contratos(df):
    """Aba Contratos com KPIs e gráficos específicos."""
    df = df.copy()

    # Converte datas
    for col_data in ["Data Venda", "Data Confirmacao", "Data Entrada"]:
        if col_data in df.columns:
            df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")

    # Filtro de mês
    hoje = datetime.now()
    col_ref = "Data Venda" if "Data Venda" in df.columns else None

    st.markdown("#### 📅 Filtro de Período")
    fcol1, fcol2, fcol3 = st.columns(3)
    with fcol1:
        mes = st.selectbox("Mês", list(range(1, 13)),
            index=hoje.month - 1,
            format_func=lambda m: ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][m-1])
    with fcol2:
        ano = st.selectbox("Ano", list(range(2023, hoje.year + 1)), index=list(range(2023, hoje.year + 1)).index(hoje.year))
    with fcol3:
        filtro_mes = st.checkbox("Aplicar filtro de mês nos KPIs", value=True)

    if col_ref and filtro_mes:
        df_kpi = df[(df[col_ref].dt.month == mes) & (df[col_ref].dt.year == ano)]
    else:
        df_kpi = df

    # KPIs
    st.markdown("---")
    k1, k2, k3, k4 = st.columns(4)
    status_col = "Status" if "Status" in df.columns else None

    with k1:
        total = len(df_kpi)
        st.metric("Total do Período", f"{total:,}".replace(",", "."))
    with k2:
        if status_col:
            aguard = df_kpi[status_col].str.lower().str.contains("aguard", na=False).sum()
            st.metric("🕐 Aguardando Aprovação", aguard)
    with k3:
        if status_col:
            ativos = df_kpi[status_col].str.lower().str.contains("ativo", na=False).sum()
            st.metric("✅ Ativos", ativos)
    with k4:
        if status_col:
            cancelados = df_kpi[status_col].str.lower().str.contains("cancel", na=False).sum()
            st.metric("❌ Cancelados", cancelados)

    st.markdown("---")

    # Aplica filtro de mês também nos gráficos
    df_graf = df_kpi if filtro_mes else df

    # Gráficos
    g1, g2 = st.columns(2)

    with g1:
        if "Vendedor OS" in df_graf.columns:
            vendas = df_graf["Vendedor OS"].value_counts().nlargest(10).reset_index()
            vendas.columns = ["Vendedor", "Contratos"]
            fig = px.bar(vendas, x="Contratos", y="Vendedor", orientation="h",
                         color="Contratos", color_continuous_scale=["#004444", "#00E5E5"],
                         title="🏆 Vendas por Vendedor")
            fig.update_layout(**LAYOUT_PLOTLY)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)

    with g2:
        if "Modalidade" in df_graf.columns:
            modal = df_graf["Modalidade"].value_counts().reset_index()
            modal.columns = ["Modalidade", "Qtd"]
            fig2 = px.pie(modal, values="Qtd", names="Modalidade",
                          title="📊 Modalidade de Venda",
                          color_discrete_sequence=CORES, hole=0.5)
            fig2.update_layout(**LAYOUT_PLOTLY)
            fig2.update_traces(textfont_color="#E0E0E0")
            st.plotly_chart(fig2, use_container_width=True)

    if "Tipo Blindagem" in df_graf.columns:
        blind = df_graf["Tipo Blindagem"].value_counts().nlargest(10).reset_index()
        blind.columns = ["Tipo", "Qtd"]
        fig3 = px.bar(blind, x="Tipo", y="Qtd",
                      color="Qtd", color_continuous_scale=["#004444", "#00E5E5"],
                      title="🛡️ Tipo de Blindagem")
        fig3.update_layout(**LAYOUT_PLOTLY)
        fig3.update_coloraxes(showscale=False)
        st.plotly_chart(fig3, use_container_width=True)

    # Tabela
    st.markdown("#### 📋 Dados Completos")
    cols_show = [c for c in ["Num. OS", "Status", "Vendedor OS", "Nome Cliente",
                              "Fabricante", "Modelo", "Tipo Blindagem", "Modalidade",
                              "Valor Venda", "Data Venda", "Liberado Financeiro?"] if c in df.columns]
    st.dataframe(df[cols_show] if cols_show else df, use_container_width=True, height=350, hide_index=True)

    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("⬇️ Baixar Contratos CSV", data=csv,
        file_name=f"chroma_contratos_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv", key="dl_contratos")


def tab_generica(aba, df):
    """Aba genérica para os demais relatórios."""
    df = df.copy()
    cols_num = df.select_dtypes(include="number").columns.tolist()
    cols_cat = df.select_dtypes(include="object").columns.tolist()

    with st.expander("🔍 Filtros", expanded=False):
        fcols = st.columns(min(3, max(1, len(cols_cat))))
        filtros = {}
        for j, col in enumerate(cols_cat[:3]):
            with fcols[j]:
                opcoes = ["Todos"] + sorted(df[col].dropna().unique().astype(str).tolist())
                sel = st.selectbox(col, opcoes, key=f"{aba}_{col}")
                if sel != "Todos":
                    filtros[col] = sel
        for col, val in filtros.items():
            df = df[df[col].astype(str) == val]

    st.caption(f"{len(df):,} registros".replace(",", "."))

    if cols_cat and cols_num:
        gcol1, gcol2 = st.columns(2)
        with gcol1:
            top = df.groupby(cols_cat[0])[cols_num[0]].sum().nlargest(10).reset_index()
            fig = px.bar(top, x=cols_num[0], y=cols_cat[0], orientation="h",
                         color=cols_num[0], color_continuous_scale=["#004444", "#00E5E5"],
                         title=f"Top 10 — {cols_num[0]} por {cols_cat[0]}")
            fig.update_layout(**LAYOUT_PLOTLY)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with gcol2:
            distrib = df[cols_cat[0]].value_counts().head(8)
            fig2 = px.pie(values=distrib.values, names=distrib.index,
                          title=f"Distribuição — {cols_cat[0]}",
                          color_discrete_sequence=CORES, hole=0.5)
            fig2.update_layout(**LAYOUT_PLOTLY)
            fig2.update_traces(textfont_color="#E0E0E0")
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 📋 Dados")
    st.dataframe(df, use_container_width=True, height=350, hide_index=True)
    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(f"⬇️ Baixar {aba} CSV", data=csv,
        file_name=f"chroma_{aba.lower()}_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv", key=f"dl_{aba}")



def tab_producao(df):
    """Aba Produção com visão operacional completa."""
    df = df.copy()

    # Converte datas
    def formatar_data(val):
        if pd.isna(val) or val == "" or val is None:
            return ""
        try:
            # Número Excel (float/int)
            n = float(val)
            if n <= 0:
                return ""
            dt = pd.Timestamp("1899-12-30") + pd.Timedelta(days=n)
            return dt.strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            pass
        try:
            # String de data
            dt = pd.to_datetime(str(val), dayfirst=True, errors="coerce")
            if pd.isna(dt):
                return ""
            return dt.strftime("%d/%m/%Y")
        except Exception:
            return str(val)

    for col in ["Data Entrada", "Data Combinada Cliente", "Data de Chegada Vidro", "Data liberação"]:
        if col in df.columns:
            df[col] = df[col].apply(formatar_data)

    # Dias em produção como numérico
    if "Qtd. Dias em produção" in df.columns:
        df["Qtd. Dias em produção"] = pd.to_numeric(df["Qtd. Dias em produção"], errors="coerce")

    # KPIs
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Total em Produção", len(df))
    with k2:
        if "Qtd. Dias em produção" in df.columns:
            media = df["Qtd. Dias em produção"].mean()
            st.metric("Média Dias em Produção", f"{media:.0f} dias" if pd.notna(media) else "—")
    with k3:
        if "Autorização Blindagem" in df.columns:
            com_ab = df["Autorização Blindagem"].notna().sum()
            st.metric("✅ Com AB", com_ab)
    with k4:
        if "Declaração Blindagem" in df.columns:
            com_db = df["Declaração Blindagem"].notna().sum()
            st.metric("✅ Com DB", com_db)

    st.markdown("---")

    # Filtros
    with st.expander("🔍 Filtros", expanded=True):
        f1, f2, f3 = st.columns(3)
        with f1:
            if "Fabricante" in df.columns:
                fab_opts = ["Todos"] + sorted(df["Fabricante"].dropna().unique().tolist())
                fab_sel = st.selectbox("Fabricante", fab_opts, key="prod_fab")
                if fab_sel != "Todos":
                    df = df[df["Fabricante"] == fab_sel]
        with f2:
            if "Status Vidro" in df.columns:
                vid_opts = ["Todos"] + sorted(df["Status Vidro"].dropna().unique().tolist())
                vid_sel = st.selectbox("Status Vidro", vid_opts, key="prod_vid")
                if vid_sel != "Todos":
                    df = df[df["Status Vidro"] == vid_sel]
        with f3:
            if "Tipo Blindagem" in df.columns:
                blind_opts = ["Todos"] + sorted(df["Tipo Blindagem"].dropna().unique().tolist())
                blind_sel = st.selectbox("Tipo Blindagem", blind_opts, key="prod_blind")
                if blind_sel != "Todos":
                    df = df[df["Tipo Blindagem"] == blind_sel]

    st.caption(f"{len(df)} OS em produção")

    # Tabela operacional
    st.markdown("#### 📋 OS em Produção")

    cols_show = [c for c in [
        "Núm. OS", "Cliente", "Fabricante", "Modelo", "Tipo Blindagem",
        "Qtd. Dias em produção", "Status Vidro", "Data de Chegada Vidro",
        "Tipo de Opaco", "Data liberação", "Data Combinada Cliente",
        "Autorização Blindagem", "Declaração Blindagem", "Responsável"
    ] if c in df.columns]

    df_show = df[cols_show].copy() if cols_show else df.copy()

    # Destaca dias em produção altos
    def highlight_dias(val):
        if pd.isna(val):
            return ""
        if val > 60:
            return "background-color: #3d1010; color: #ff6b6b"
        elif val > 30:
            return "background-color: #3d2d10; color: #ffa94d"
        return ""

    if "Qtd. Dias em produção" in df_show.columns:
        st.dataframe(
            df_show.style.map(highlight_dias, subset=["Qtd. Dias em produção"]),
            use_container_width=True, height=400, hide_index=True
        )
    else:
        st.dataframe(df_show, use_container_width=True, height=400, hide_index=True)

    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("⬇️ Baixar Produção CSV", data=csv,
        file_name=f"chroma_producao_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv", key="dl_producao")


for i, (aba, tab) in enumerate(zip(abas, tabs)):
    df = dados[aba].copy()
    with tab:
        if df.empty:
            st.warning(f"Sem dados em {aba}.")
            continue
        if aba == "Contratos":
            tab_contratos(df)
        elif aba == "Producao":
            tab_producao(df)
        else:
            tab_generica(aba, df)
