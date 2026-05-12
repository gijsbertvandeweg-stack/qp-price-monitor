import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(
    page_title="QP Price Monitor",
    page_icon="📊",
    layout="wide",
)

EXCEL_PATH = Path(__file__).parent / "data.xlsx"

@st.cache_data(ttl=600)  # Herlaadt elke 10 minuten automatisch
def load_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name="Sheet1")
    df["datum"] = pd.to_datetime(df["datum"], errors="coerce")
    df = df.dropna(subset=["datum"])
    df["prijs"] = pd.to_numeric(df["prijs"], errors="coerce")
    df.loc[df["prijs"] == 0, "prijs"] = None
    return df

df = load_data()

SUCCESS_STATUSES = [
    "Succes (Schema.org)", "Succes (JSON-LD)", "Succes (Visual)",
    "Succes", "Succes (HTML price)", "Succes (Visueel HTML)",
]
df_ok = df[df["status"].isin(SUCCESS_STATUSES) & df["prijs"].notna()]

# ── Sidebar ────────────────────────────────────────────────────────────────────
st.sidebar.title("Filters")

leveranciers = ["Alle"] + sorted(df["Leverancier"].dropna().unique().tolist())
sel_leverancier = st.sidebar.selectbox("Leverancier", leveranciers)

retailers = sorted(df["retailer"].dropna().unique().tolist())
sel_retailers = st.sidebar.multiselect("Retailer(s)", retailers, default=retailers)

filtered = df_ok.copy()
if sel_leverancier != "Alle":
    filtered = filtered[filtered["Leverancier"] == sel_leverancier]
if sel_retailers:
    filtered = filtered[filtered["retailer"].isin(sel_retailers)]

producten = sorted(filtered["product_naam"].dropna().unique().tolist())
sel_product = st.sidebar.selectbox("Product", ["— selecteer —"] + producten)

datums = sorted(df["datum"].dt.date.unique())
datum_min, datum_max = datums[0], datums[-1]
sel_datum = st.sidebar.date_input(
    "Periode",
    value=(datum_min, datum_max),
    min_value=datum_min,
    max_value=datum_max,
)
if isinstance(sel_datum, tuple) and len(sel_datum) == 2:
    filtered = filtered[
        (filtered["datum"].dt.date >= sel_datum[0])
        & (filtered["datum"].dt.date <= sel_datum[1])
    ]

# ── Header ─────────────────────────────────────────────────────────────────────
st.title("📊 QP Price Monitor")
st.caption(f"Databron: {EXCEL_PATH.name}  ·  {len(df):,} meetpunten  ·  {df['datum'].dt.date.nunique()} meetmomenten")

# ── KPI cards ──────────────────────────────────────────────────────────────────
latest_date = df_ok["datum"].max()
prev_dates = sorted(df_ok[df_ok["datum"] < latest_date]["datum"].unique())
prev_date = prev_dates[-1] if prev_dates else None

col1, col2, col3, col4 = st.columns(4)

with col1:
    n_products = df_ok["product_naam"].nunique()
    st.metric("Producten gemonitord", n_products)

with col2:
    n_retailers = df_ok["retailer"].nunique()
    st.metric("Supermarkten", n_retailers)

with col3:
    success_rate = df[df["status"].isin(SUCCESS_STATUSES)].shape[0] / len(df) * 100
    st.metric("Scrape-succesrate", f"{success_rate:.1f}%")

with col4:
    st.metric("Laatste meting", latest_date.strftime("%d-%m-%Y"))

st.divider()

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 Prijsverloop", "⚖️ Queens vs Concurrent", "🏪 Retailervergelijking", "🔍 Datakwaliteit", "📋 Alle data"
])

# ── Tab 1: Prijsverloop ────────────────────────────────────────────────────────
with tab1:
    if sel_product == "— selecteer —":
        st.info("Selecteer een product in de sidebar om het prijsverloop te bekijken.")
    else:
        prod_df = filtered[filtered["product_naam"] == sel_product].copy()
        prod_df = prod_df.sort_values("datum")

        fig = px.line(
            prod_df,
            x="datum",
            y="prijs",
            color="retailer",
            markers=True,
            title=f"Prijsverloop — {sel_product}",
            labels={"datum": "Datum", "prijs": "Prijs (€)", "retailer": "Retailer"},
        )
        fig.update_layout(hovermode="x unified", yaxis_tickprefix="€")
        st.plotly_chart(fig, use_container_width=True)

        # Prijsmutaties tabel
        st.subheader("Prijsmutaties")
        if prev_date is not None:
            cur = prod_df[prod_df["datum"] == latest_date][["retailer", "prijs"]].rename(columns={"prijs": "Nu"})
            prv = prod_df[prod_df["datum"] == prev_date][["retailer", "prijs"]].rename(columns={"prijs": "Vorige"})
            mut = cur.merge(prv, on="retailer", how="outer").sort_values("retailer")
            mut["Verschil"] = mut["Nu"] - mut["Vorige"]
            mut["Δ%"] = (mut["Verschil"] / mut["Vorige"] * 100).round(1)
            mut["Nu"] = mut["Nu"].apply(lambda x: f"€{x:.2f}" if pd.notna(x) else "—")
            mut["Vorige"] = mut["Vorige"].apply(lambda x: f"€{x:.2f}" if pd.notna(x) else "—")
            mut["Verschil"] = mut["Verschil"].apply(lambda x: f"{x:+.2f}" if pd.notna(x) else "—")
            mut["Δ%"] = mut["Δ%"].apply(lambda x: f"{x:+.1f}%" if pd.notna(x) else "—")
            st.dataframe(mut.reset_index(drop=True), use_container_width=True, hide_index=True)
        else:
            st.write("Niet genoeg meetmomenten voor vergelijking.")

# ── Tab 2: Queens vs Concurrent ───────────────────────────────────────────────
with tab2:
    st.subheader("Meest recente prijzen — Queens vs. Concurrent")

    latest_all = df_ok[df_ok["datum"] == latest_date].copy()

    queens_latest = (
        latest_all[latest_all["Leverancier"] == "Queens"]
        .groupby(["product_naam", "retailer"])["prijs"]
        .mean()
        .reset_index()
        .rename(columns={"prijs": "Queens prijs"})
    )
    conc_latest = (
        latest_all[latest_all["Leverancier"] == "Concurrent"]
        .groupby(["product_naam", "retailer"])["prijs"]
        .mean()
        .reset_index()
        .rename(columns={"prijs": "Concurrent prijs"})
    )

    # Prijsverdeling boxplot per leverancier
    fig2 = px.box(
        latest_all,
        x="Leverancier",
        y="prijs",
        color="Leverancier",
        points="all",
        title=f"Prijsverdeling op {latest_date.strftime('%d-%m-%Y')}",
        labels={"prijs": "Prijs (€)"},
        color_discrete_map={"Queens": "#0068c9", "Concurrent": "#ff6b35"},
    )
    fig2.update_layout(yaxis_tickprefix="€", showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Queens producten**")
        q_tbl = queens_latest.copy()
        q_tbl["Queens prijs"] = q_tbl["Queens prijs"].apply(lambda x: f"€{x:.2f}")
        st.dataframe(q_tbl, use_container_width=True, hide_index=True)
    with col_b:
        st.markdown("**Concurrent producten**")
        c_tbl = conc_latest.copy()
        c_tbl["Concurrent prijs"] = c_tbl["Concurrent prijs"].apply(lambda x: f"€{x:.2f}")
        st.dataframe(c_tbl, use_container_width=True, hide_index=True)

# ── Tab 3: Retailervergelijking ────────────────────────────────────────────────
with tab3:
    if sel_product == "— selecteer —":
        st.info("Selecteer een product in de sidebar voor de retailervergelijking.")
    else:
        prod_latest = df_ok[
            (df_ok["product_naam"] == sel_product) & (df_ok["datum"] == latest_date)
        ].copy()

        if prod_latest.empty:
            st.warning("Geen recente data beschikbaar voor dit product.")
        else:
            prod_latest = prod_latest.sort_values("prijs")
            fig3 = px.bar(
                prod_latest,
                x="retailer",
                y="prijs",
                color="Leverancier",
                title=f"Prijsvergelijking per retailer — {sel_product}",
                labels={"prijs": "Prijs (€)", "retailer": "Retailer"},
                color_discrete_map={"Queens": "#0068c9", "Concurrent": "#ff6b35"},
                text_auto=".2f",
            )
            fig3.update_traces(texttemplate="€%{y:.2f}")
            fig3.update_layout(yaxis_tickprefix="€")
            st.plotly_chart(fig3, use_container_width=True)

            # Goedkoopste / duurste
            c1, c2 = st.columns(2)
            with c1:
                cheapest = prod_latest.loc[prod_latest["prijs"].idxmin()]
                st.success(f"**Goedkoopst:** {cheapest['retailer']} — €{cheapest['prijs']:.2f}")
            with c2:
                most_exp = prod_latest.loc[prod_latest["prijs"].idxmax()]
                st.error(f"**Duurste:** {most_exp['retailer']} — €{most_exp['prijs']:.2f}")

# ── Tab 4: Datakwaliteit ───────────────────────────────────────────────────────
with tab4:
    st.subheader("Scrape-status overzicht")

    status_counts = df.groupby(["datum", "status"]).size().reset_index(name="aantal")
    fig4 = px.bar(
        status_counts,
        x="datum",
        y="aantal",
        color="status",
        title="Scrape-resultaten per meetmoment",
        labels={"datum": "Datum", "aantal": "Aantal", "status": "Status"},
    )
    st.plotly_chart(fig4, use_container_width=True)

    st.subheader("Mislukte / ontbrekende metingen")
    failed = df[~df["status"].isin(SUCCESS_STATUSES)][
        ["datum", "Leverancier", "product_naam", "retailer", "status"]
    ].sort_values("datum", ascending=False)
    st.dataframe(failed.reset_index(drop=True), use_container_width=True, hide_index=True)

    st.subheader("Meest recente meting per product × retailer")
    coverage = (
        df_ok.groupby(["product_naam", "retailer"])["datum"]
        .max()
        .reset_index()
        .rename(columns={"datum": "Laatste meting"})
    )
    coverage["Laatste meting"] = coverage["Laatste meting"].dt.strftime("%d-%m-%Y")
    st.dataframe(coverage, use_container_width=True, hide_index=True)

# ── Tab 5: Alle data ───────────────────────────────────────────────────────────
with tab5:
    st.subheader("Alle meetdata")

    tbl_df = filtered.copy()

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        zoek = st.text_input("Zoek op productnaam", "")
    with col_f2:
        status_opties = ["Alle"] + sorted(tbl_df["status"].dropna().unique().tolist())
        sel_status = st.selectbox("Status", status_opties, key="tbl_status")
    with col_f3:
        sort_kolom = st.selectbox("Sorteren op", ["datum", "product_naam", "retailer", "prijs"], key="tbl_sort")

    if zoek:
        tbl_df = tbl_df[tbl_df["product_naam"].str.contains(zoek, case=False, na=False)]
    if sel_status != "Alle":
        tbl_df = tbl_df[tbl_df["status"] == sel_status]

    tbl_df = tbl_df.sort_values(sort_kolom, ascending=(sort_kolom != "datum"))
    tbl_display = tbl_df[["datum", "Leverancier", "Artikel_nummer", "product_naam", "prijs", "retailer", "status"]].copy()
    tbl_display["datum"] = tbl_display["datum"].dt.strftime("%d-%m-%Y")
    tbl_display["prijs"] = tbl_display["prijs"].apply(lambda x: f"€{x:.2f}" if pd.notna(x) else "—")

    st.caption(f"{len(tbl_display):,} rijen")
    st.dataframe(tbl_display.reset_index(drop=True), use_container_width=True, hide_index=True)

    csv = tbl_df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
    st.download_button(
        label="⬇️ Download als CSV",
        data=csv,
        file_name="price_monitor_export.csv",
        mime="text/csv",
    )
