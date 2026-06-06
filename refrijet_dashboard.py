import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO
import re

st.set_page_config(
    page_title="Refrijet — Desenvolvimento de Produtos",
    page_icon="❄️",
    layout="wide",
)

st.markdown("""
<style>
.block-container { padding-top: 1.4rem; padding-bottom: 1rem; }
.metric-card {
    background: #f8f9fa; border-radius: 10px;
    padding: 16px 20px; border-left: 4px solid;
    margin-bottom: 4px;
}
.metric-title { font-size: 11px; color: #6c757d; text-transform: uppercase;
    letter-spacing: .05em; margin-bottom: 4px; }
.metric-value { font-size: 28px; font-weight: 600; line-height: 1; }
.metric-sub   { font-size: 11px; color: #6c757d; margin-top: 4px; }
</style>
""", unsafe_allow_html=True)

FAMILIA_CORES = {
    "21 Compressores":           "#1f77b4",
    "22 Condensadores":          "#f59e0b",
    "44 Evaporadores":           "#10b981",
    "33 Comp. p/ Compressores":  "#ef4444",
    "30 Válvulas e Filtros":     "#8b5cf6",
    "35 Gases":                  "#ec4899",
    "29 Eletroventiladores":     "#14b8a6",
}
DEFAULT_COR = "#94a3b8"


@st.cache_data(show_spinner=False)
def carregar(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name="Solicitações", header=1)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["N° ORDEM"])
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df["QTD MENSAL"] = pd.to_numeric(df["QTD MENSAL"], errors="coerce").fillna(0).astype(int)
    df["PREÇO MERCADO"] = pd.to_numeric(df["PREÇO MERCADO"], errors="coerce")
    for col in ["FILIAL", "SOLICITANTE", "FAMILIA", "MONTADORA", "MODELO"]:
        if col in df.columns:
            df[col] = df[col].fillna("Não informado").str.strip()
    df["ANO"] = df["ANO"].fillna("—").astype(str).str.replace(r"\.0$", "", regex=True)
    for c in ["COD. ROYCE", "COD. HDS", "COD. OEM"]:
        if c in df.columns:
            df[c] = df[c].fillna("—").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    return df


def exportar_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Solicitações")
        wb = writer.book
        ws = writer.sheets["Solicitações"]
        hdr = wb.add_format({"bold": True, "bg_color": "#1f4e79", "font_color": "white",
                              "border": 1, "align": "center"})
        cel = wb.add_format({"border": 1})
        for i, col in enumerate(df.columns):
            ws.write(0, i, col, hdr)
            w = max(len(str(col)), df[col].astype(str).str.len().max())
            ws.set_column(i, i, min(w + 2, 40))
        for r, row in enumerate(df.itertuples(index=False), 1):
            for c, val in enumerate(row):
                ws.write(r, c, val if pd.notna(val) else "", cel)
        # aba resumo
        resumo = (df.groupby("FAMILIA")["QTD MENSAL"]
                    .agg(["sum", "count"])
                    .rename(columns={"sum": "Qtd. Total", "count": "Solicitações"})
                    .sort_values("Qtd. Total", ascending=False)
                    .reset_index())
        resumo.to_excel(writer, index=False, sheet_name="Resumo por Família")
    return buf.getvalue()


def kpi(titulo, valor, subtitulo, cor):
    st.markdown(f"""
    <div class="metric-card" style="border-color:{cor}">
        <div class="metric-title">{titulo}</div>
        <div class="metric-value" style="color:{cor}">{valor}</div>
        <div class="metric-sub">{subtitulo}</div>
    </div>""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## ❄️ Refrijet — Desenvolvimento de Produtos")
st.markdown("Faça o upload da planilha de solicitações para gerar o dashboard.")

uploaded = st.file_uploader(
    "Selecione o arquivo Excel (.xlsx)",
    type=["xlsx"],
    label_visibility="collapsed",
)

if not uploaded:
    st.info("Aguardando arquivo… Use o botão acima para selecionar o Excel de Solicitações.")
    st.stop()

with st.spinner("Carregando dados..."):
    df_raw = carregar(uploaded.read())

st.success(f"✅  {len(df_raw)} solicitações carregadas.")
st.divider()

# ── Filtros laterais ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Filtros")
    sel_filial     = st.selectbox("Filial",            ["Todas"]  + sorted(df_raw["FILIAL"].unique().tolist()))
    sel_familia    = st.selectbox("Família",           ["Todas"]  + sorted(df_raw["FAMILIA"].unique().tolist()))
    sel_montadora  = st.selectbox("Montadora",         ["Todas"]  + sorted(df_raw["MONTADORA"].unique().tolist()))
    sel_solicitante= st.selectbox("Solicitante",       ["Todos"]  + sorted(df_raw["SOLICITANTE"].unique().tolist()))

    datas = df_raw["DATA"].dropna()
    if not datas.empty:
        d_min, d_max = datas.min().date(), datas.max().date()
        intervalo = st.date_input("Período", value=(d_min, d_max),
                                  min_value=d_min, max_value=d_max)
    else:
        intervalo = None

    st.divider()
    st.markdown("### 📥 Exportar")
    st.download_button("⬇️ Baixar Excel completo",
                       data=exportar_excel(df_raw),
                       file_name="refrijet_solicitacoes.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── Filtrar ───────────────────────────────────────────────────────────────────
df = df_raw.copy()
if sel_filial      != "Todas": df = df[df["FILIAL"]      == sel_filial]
if sel_familia     != "Todas": df = df[df["FAMILIA"]     == sel_familia]
if sel_montadora   != "Todas": df = df[df["MONTADORA"]   == sel_montadora]
if sel_solicitante != "Todos": df = df[df["SOLICITANTE"] == sel_solicitante]
if intervalo and len(intervalo) == 2:
    df = df[df["DATA"].between(pd.Timestamp(intervalo[0]), pd.Timestamp(intervalo[1]))]

# ── KPIs ──────────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
with c1: kpi("Total de solicitações",     len(df),                         "registros no período",           "#1f77b4")
with c2: kpi("Possíveis vendas perdidas", f"{int(df['QTD MENSAL'].sum()):,}".replace(",", "."), "unidades/mês sem o produto", "#f59e0b")
with c3: kpi("Famílias solicitadas",      df["FAMILIA"].nunique(),          "categorias de produto",          "#10b981")
with c4: kpi("Solicitantes ativos",       df["SOLICITANTE"].nunique(),      "vendedores / filiais",           "#8b5cf6")

st.divider()

# ── Gráficos linha 1 ──────────────────────────────────────────────────────────
col_a, col_b = st.columns([3, 2])

with col_a:
    st.markdown("##### Qtd. mensal por família")
    fam_df = (df.groupby("FAMILIA")["QTD MENSAL"].sum()
                .reset_index().sort_values("QTD MENSAL"))
    fam_df["cor"] = fam_df["FAMILIA"].map(FAMILIA_CORES).fillna(DEFAULT_COR)
    chart = (alt.Chart(fam_df)
               .mark_bar()
               .encode(
                   x=alt.X("QTD MENSAL:Q", title="Unidades/mês"),
                   y=alt.Y("FAMILIA:N", sort="-x", title=""),
                   color=alt.Color("FAMILIA:N",
                                   scale=alt.Scale(domain=list(fam_df["FAMILIA"]),
                                                   range=list(fam_df["cor"])),
                                   legend=None),
                   tooltip=["FAMILIA", "QTD MENSAL"],
               ).properties(height=260))
    st.altair_chart(chart, use_container_width=True)

with col_b:
    st.markdown("##### Solicitações por filial")
    fil_df = df["FILIAL"].value_counts().reset_index()
    fil_df.columns = ["Filial", "Qtd"]
    chart2 = (alt.Chart(fil_df)
                .mark_arc(innerRadius=55)
                .encode(
                    theta=alt.Theta("Qtd:Q"),
                    color=alt.Color("Filial:N", legend=alt.Legend(orient="bottom")),
                    tooltip=["Filial", "Qtd"],
                ).properties(height=260))
    st.altair_chart(chart2, use_container_width=True)

# ── Gráficos linha 2 ──────────────────────────────────────────────────────────
col_c, col_d = st.columns([2, 3])

with col_c:
    st.markdown("##### Solicitantes — nº de solicitações")
    sol_df = df["SOLICITANTE"].value_counts().head(10).reset_index()
    sol_df.columns = ["Solicitante", "Qtd"]
    chart3 = (alt.Chart(sol_df)
                .mark_bar(color="#6366f1")
                .encode(
                    x=alt.X("Qtd:Q", title="Solicitações"),
                    y=alt.Y("Solicitante:N", sort="-x", title=""),
                    tooltip=["Solicitante", "Qtd"],
                ).properties(height=300))
    st.altair_chart(chart3, use_container_width=True)

with col_d:
    st.markdown("##### Top 15 montadoras — Qtd. mensal acumulada")
    mo_df = (df.groupby("MONTADORA")["QTD MENSAL"].sum()
               .reset_index().sort_values("QTD MENSAL", ascending=False).head(15))
    chart4 = (alt.Chart(mo_df)
                .mark_bar(color="#0ea5e9")
                .encode(
                    x=alt.X("QTD MENSAL:Q", title="Unidades/mês"),
                    y=alt.Y("MONTADORA:N", sort="-x", title=""),
                    tooltip=["MONTADORA", "QTD MENSAL"],
                ).properties(height=300))
    st.altair_chart(chart4, use_container_width=True)

# ── Evolução temporal ─────────────────────────────────────────────────────────
if df["DATA"].notna().any():
    st.markdown("##### Evolução diária de solicitações")
    daily = (df.dropna(subset=["DATA"])
               .groupby(df["DATA"].dt.date)["QTD MENSAL"].sum()
               .reset_index().rename(columns={"DATA": "Data", "QTD MENSAL": "Qtd"}))
    daily["Data"] = pd.to_datetime(daily["Data"])
    chart5 = (alt.Chart(daily)
                .mark_line(point=True, color="#1f77b4")
                .encode(
                    x=alt.X("Data:T", title=""),
                    y=alt.Y("Qtd:Q", title="Qtd/mês"),
                    tooltip=[alt.Tooltip("Data:T", format="%d/%m/%Y"), "Qtd"],
                ).properties(height=200))
    st.altair_chart(chart5, use_container_width=True)

st.divider()

# ── Top 10 ────────────────────────────────────────────────────────────────────
st.markdown("##### 🏆 Top 10 — Itens mais solicitados")
cols_top = [c for c in ["N° ORDEM","FAMILIA","MONTADORA","MODELO","ANO",
                         "COD. ROYCE","COD. HDS","QTD MENSAL","PREÇO MERCADO"] if c in df.columns]
st.dataframe(df.sort_values("QTD MENSAL", ascending=False).head(10)[cols_top],
             use_container_width=True, hide_index=True)

st.divider()

# ── Tabela completa ───────────────────────────────────────────────────────────
st.markdown(f"##### 📋 Todas as solicitações ({len(df)} registros)")
busca = st.text_input("🔎 Buscar em qualquer campo", placeholder="Ex: Renault, Compressor, RS…")
cols_show = [c for c in ["N° ORDEM","DATA","FILIAL","SOLICITANTE","FAMILIA",
                          "MONTADORA","MODELO","ANO","COD. ROYCE","COD. HDS",
                          "COD. OEM","QTD MENSAL","PREÇO MERCADO","OBS"] if c in df.columns]
df_show = df[cols_show].copy()
if busca:
    mask = df_show.apply(lambda col: col.astype(str).str.contains(busca, case=False, na=False)).any(axis=1)
    df_show = df_show[mask]

st.dataframe(df_show, use_container_width=True, hide_index=True, height=380)
st.download_button("⬇️ Exportar tabela atual",
                   data=exportar_excel(df[cols_show]),
                   file_name="refrijet_filtrado.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("MDL Tech · Gestão e Inteligência de Dados")
