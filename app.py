import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.colors import qualitative
from io import BytesIO
import numpy as np

# ==================================================
# Page config
# ==================================================
st.set_page_config(layout="wide")

# ==================================================
# Tabs
# ==================================================
tab_corr, tab_stress, tab_legenda = st.tabs(
    ["Correlation", "Stress Test", "Legend"]
)

# ==================================================
# DATA LOADING
# ==================================================
@st.cache_data
def load_corr_sheets(path):
    """Ritorna la lista dei sheet di un file Excel"""
    xls = pd.ExcelFile(path)
    return xls.sheet_names

@st.cache_data
def load_corr_data(path, sheet):
    """Carica i dati di un sheet specifico e imposta la prima colonna come index datetime"""
    df = pd.read_excel(path, sheet_name=sheet)
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0])
    return df.set_index(df.columns[0]).sort_index()

@st.cache_data
def load_stress_data(path):
    xls = pd.ExcelFile(path)
    records = []
    for sheet in xls.sheet_names:
        portfolio, scenario_name = sheet.split("&&", 1) if "&&" in sheet else (sheet, sheet)
        df = pd.read_excel(xls, sheet_name=sheet)
        df = df[df.iloc[:, 0] == "Total"]
        df = df.rename(columns={"Stress PnL": "StressPnL"})
        df["Date"] = pd.to_datetime(df["Date"])
        df["Portfolio"] = portfolio
        df["ScenarioName"] = scenario_name
        records.append(df[["Date", "Scenario", "StressPnL", "Portfolio", "ScenarioName"]])
    return pd.concat(records, ignore_index=True)

@st.cache_data
def load_stress_bystrat(path):
    xls = pd.ExcelFile(path)
    records = []
    for sheet in xls.sheet_names:
        portfolio, scenario = sheet.split("&&", 1) if "&&" in sheet else (sheet, sheet)
        df = pd.read_excel(xls, sheet_name=sheet)
        name_col = df.columns[0]
        date_col = df.columns[df.columns.str.contains("Date", case=False, regex=True)][0]
        pnl_col = df.columns[df.columns.str.contains("Stress PnL", case=False, regex=True)][0]
        df = df.rename(columns={name_col: "Name", date_col: "Date", pnl_col: "StressPnL"})
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Portfolio"] = portfolio
        df["ScenarioName"] = scenario
        records.append(df[["Name", "Date", "StressPnL", "Portfolio", "ScenarioName"]])
    combined = pd.concat(records, ignore_index=True)
    combined = combined.sort_values(["Date", "Portfolio", "ScenarioName", "Name"])
    return combined

@st.cache_data
def load_legenda(sheet, cols):
    return pd.read_excel("Legenda.xlsx", sheet_name=sheet, usecols=cols)

# ==================================================
# NAME MAP (Ticker → Name)
# ==================================================
@st.cache_data
def load_name_map():
    legenda = load_legenda("Portafogli", "A:C")
    return dict(zip(legenda["Ticker"], legenda["Name"]))

NAME_MAP = load_name_map()
def pretty_name(x):
    return NAME_MAP.get(x, x)

# ==================================================
# LOAD DATA
# ==================================================
corr_sheets = load_corr_sheets("corr_ptf.xlsx")
stress_data = load_stress_data("stress_test_bystrat.xlsx")
stress_bystrat = load_stress_bystrat("stress_test_bystrat.xlsx")

# ==================================================
# TAB — CORRELATION
# ==================================================
with tab_corr:
    st.title("Correlation Dashboard")

    col_ctrl, col_plot = st.columns([0.7, 4.3])

    with col_ctrl:
        st.subheader("Controls")

        # selezione sheet
        selected_sheet = st.selectbox("Select Sheet", corr_sheets)
        corr = load_corr_data("corr_ptf.xlsx", selected_sheet)

        # selezione date
        start, end = st.date_input(
            "Date range",
            (corr.index.min().date(), corr.index.max().date())
        )
        df = corr.loc[pd.to_datetime(start):pd.to_datetime(end)]

        # selezione serie
        selected = st.multiselect(
            "Select series",
            options=df.columns.tolist(),
            default=df.columns.tolist(),
            format_func=pretty_name
        )

    with col_plot:
        st.subheader("Correlation Time Series")
        fig = go.Figure()
        palette = qualitative.Plotly
        for i, c in enumerate(selected):
            fig.add_trace(go.Scatter(
                x=df.index,
                y=df[c] * 100,
                name=pretty_name(c),
                line=dict(color=palette[i % len(palette)])
            ))
        fig.update_layout(height=600, template="plotly_white", yaxis_title="Correlation (%)")
        st.plotly_chart(fig, use_container_width=True)

        # Download Excel
        output = BytesIO()
        (df[selected] * 100).to_excel(output)
        st.download_button(
            "📥 Download Correlation Time Series as Excel",
            output.getvalue(),
            "correlation_time_series.xlsx"
        )

        # Radar chart
        st.subheader("Correlation Radar")
        snapshot_date = df.index.max()
        snapshot = df.loc[snapshot_date, selected]
        mean_corr = df[selected].mean()
        theta = [pretty_name(c) for c in selected]

        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(
            r=snapshot.values * 100,
            theta=theta,
            name=f"End date ({snapshot_date.date()})",
            line=dict(width=3)
        ))
        fig_radar.add_trace(go.Scatterpolar(
            r=mean_corr.values * 100,
            theta=theta,
            name="Period mean",
            line=dict(dash="dot")
        ))
        fig_radar.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[-100,100], ticksuffix="%")),
            template="plotly_white", height=650
        )
        st.plotly_chart(fig_radar, use_container_width=True)

        # Summary Statistics
        st.subheader("Summary Statistics")
        stats_df = pd.DataFrame(index=selected)
        stats_df.insert(0, "Name", [pretty_name(s) for s in selected])
        stats_df["Mean (%)"] = df[selected].mean() * 100
        stats_df["Min (%)"] = df[selected].min() * 100
        stats_df["Min Date"] = [df[col][df[col] == df[col].min()].index.max() for col in selected]
        stats_df["Max (%)"] = df[selected].max() * 100
        stats_df["Max Date"] = [df[col][df[col] == df[col].max()].index.max() for col in selected]
        stats_df["Min Date"] = pd.to_datetime(stats_df["Min Date"]).dt.strftime("%d/%m/%Y")
        stats_df["Max Date"] = pd.to_datetime(stats_df["Max Date"]).dt.strftime("%d/%m/%Y")
        st.dataframe(stats_df.style.format({"Mean (%)": "{:.2f}%", "Min (%)": "{:.2f}%", "Max (%)": "{:.2f}%"}), use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            stats_df.to_excel(writer, sheet_name="Summary Statistics", index=False)
        st.download_button("📥 Download Summary Statistics as Excel", output.getvalue(), "summary_statistics.xlsx")

# ==================================================
# TAB — STRESS TEST
# ==================================================
with tab_stress:
    st.title("Dynamic Asset Allocation vs Funds")

    col_ctrl, col_plot = st.columns([0.7, 4.3])

    with col_ctrl:
        st.subheader("Controls")
        dates = sorted(stress_data["Date"].unique())
        date = pd.to_datetime(st.selectbox("Select date", [d.strftime("%Y/%m/%d") for d in dates], index=len(dates)-1))
        df = stress_data[stress_data["Date"] == date]

        portfolios = df["Portfolio"].unique().tolist()
        sel_ports = st.multiselect("Select portfolios", portfolios, default=portfolios, format_func=pretty_name)
        df = df[df["Portfolio"].isin(sel_ports)]

        scenarios = df["ScenarioName"].unique().tolist()
        sel_scen = st.multiselect("Select scenarios", scenarios, default=scenarios)
        df = df[df["ScenarioName"].isin(sel_scen)]

    with col_plot:
        st.subheader("Stress Test PnL")
        fig = go.Figure()
        for p in sel_ports:
            d = df[df["Portfolio"] == p]
            fig.add_trace(go.Bar(x=d["ScenarioName"], y=d["StressPnL"], name=pretty_name(p)))
        fig.update_layout(barmode="group", height=600, template="plotly_white", yaxis_title="Stress PnL (bps)")
        st.plotly_chart(fig, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Stress Test PnL", index=False)
        st.download_button("📥 Download Stress PnL as Excel", output.getvalue(), "stress_test_pnl.xlsx")

# ==================================================
# TAB — LEGENDA
# ==================================================
with tab_legenda:
    st.subheader("Series")
    st.dataframe(load_legenda("Portafogli", "A:C"), hide_index=True)
    st.subheader("Stress Scenarios")
    st.dataframe(load_legenda("Scenari", "A:C"), hide_index=True)
