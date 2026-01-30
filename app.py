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
        selected_sheet = st.selectbox(
            "Select Portfolio",
            corr_sheets,
            format_func=pretty_name
        )

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
        fig.update_layout(
            height=600,
            template="plotly_white",
            yaxis_title="Correlation",
            yaxis=dict(
                ticksuffix="%"  
            )
        )
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
    st.title("Stress Test Dashboard")

    col_ctrl, col_plot = st.columns([0.7, 4.3])

    with col_ctrl:
        st.subheader("Controls")

        # ------------------------------
        # Selezione data principale per Stress Test
        # ------------------------------
        dates = sorted(stress_data["Date"].unique())
        date = pd.to_datetime(
            st.selectbox(
                "Select date",  # label condivisa
                [d.strftime("%Y/%m/%d") for d in dates],
                index=len(dates) - 1
            )
        )

        df = stress_data[stress_data["Date"] == date]

        # Selezione Portfolios
        portfolios = df["Portfolio"].unique().tolist()
        sel_ports = st.multiselect(
            "Select portfolios",
            portfolios,
            default=portfolios,
            format_func=pretty_name
        )
        df = df[df["Portfolio"].isin(sel_ports)]

        # Selezione Scenarios
        scenarios = df["ScenarioName"].unique().tolist()
        sel_scen = st.multiselect(
            "Select scenarios",
            scenarios,
            default=scenarios
        )
        df = df[df["ScenarioName"].isin(sel_scen)]

    with col_plot:
        # ------------------------------
        # Bar chart principale
        # ------------------------------
        st.subheader("Stress Test PnL")

        fig = go.Figure()
        for p in sel_ports:
            d = df[df["Portfolio"] == p]
            fig.add_trace(go.Bar(
                x=d["ScenarioName"],
                y=d["StressPnL"],
                name=pretty_name(p)
            ))

        fig.update_layout(
            barmode="group",
            height=600,
            template="plotly_white",
            yaxis_title="Stress PnL (bps)"
        )

        st.plotly_chart(fig, use_container_width=True)

        # Download Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Stress Test PnL", index=False)
        st.download_button(
            label="📥 Download Stress PnL as Excel",
            data=output.getvalue(),
            file_name="stress_test_pnl.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_stress_pnl"
        )

        # ==============================
        # Expand for By Strategy Analysis
        # ==============================
        with st.expander("Expand for Stress Test analysis by strategy", expanded=False):
        
            # --------------------
            # Selezione data by strategy
            # --------------------
            dates_bystrat = sorted(stress_bystrat["Date"].dropna().unique())
            selected_date = pd.to_datetime(
                st.selectbox(
                    "Select date",
                    [d.strftime("%Y/%m/%d") for d in dates_bystrat],
                    index=len(dates_bystrat) - 1,
                    key="bystrat_date"
                )
            )
        
            # Selezione Portfolio
            clicked_portfolio = st.selectbox(
                "Analysis Portfolio",
                sel_ports,
                format_func=pretty_name
            )
        
            # Selezione Scenario
            clicked_scenario = st.selectbox(
                "Scenario",
                sel_scen
            )
        
            # --------------------
            # Filtra dati per strategia
            # --------------------
            df_detail = stress_bystrat[
                (stress_bystrat["Date"] == selected_date) &
                (stress_bystrat["Portfolio"] == clicked_portfolio) &
                (stress_bystrat["ScenarioName"] == clicked_scenario) &
                (stress_bystrat["Name"] != "Total")  # esclude la riga Total
            ]
        
            if not df_detail.empty:
                # --------------------
                # Prepara colori dinamici
                # --------------------
                vals = pd.to_numeric(df_detail["StressPnL"], errors="coerce").fillna(0).values
                max_abs = np.max(np.abs(vals)) if np.max(np.abs(vals)) != 0 else 1
                
                colors = []
                for v in vals:
                    v = float(v)  # assicuriamoci sia float
                    neg = int(np.clip(255 * abs(min(0, v) / max_abs), 0, 255))
                    pos = int(np.clip(255 * max(0, v) / max_abs, 0, 255))
                    colors.append(f"rgba({neg},{pos},0,0.8)")
        
                df_tm = df_detail.copy()
                df_tm["size"] = df_tm["StressPnL"].abs().clip(lower=0.01)
        
                # --------------------
                # Treemap
                # --------------------
                root_label = f"{pretty_name(clicked_portfolio)} - {clicked_scenario} ({selected_date.date()})"
        
                labels = [root_label] + df_tm["Name"].tolist()
                parents = [""] + [root_label] * len(df_tm)
                values = [df_tm["size"].sum()] + df_tm["size"].tolist()
                colors = ["white"] + colors
                texts = [""] + df_tm["StressPnL"].round(2).astype(str).tolist()
        
                vals = df_detail["StressPnL"].values
                max_abs = np.max(np.abs(vals)) if np.max(np.abs(vals)) != 0 else 1
                colors = [
                    f"rgba({int(255*abs(min(0,v)/max_abs))},{int(255*max(0,v)/max_abs)},0,0.8)"
                    for v in vals
                ]
                
                df_tm = df_detail.copy()
                df_tm["size"] = df_tm["StressPnL"].abs().clip(lower=0.01)                
                labels = [root_label] + df_tm.iloc[:, 0].tolist()
                parents = [""] + [root_label] * len(df_tm)
                values = [df_tm["size"].sum()] + df_tm["size"].tolist()
                colors = ["white"] + df_tm["StressPnL"].tolist()
                texts = [""] + df_tm["StressPnL"].round(2).astype(str).tolist()
                
                fig_detail = go.Figure(
                    go.Treemap(
                        labels=labels,
                        parents=parents,
                        values=values,
                        marker=dict(
                            colors=colors,
                            colorscale="RdYlGn",
                            cmid=0,
                            line=dict(color="white", width=2)
                        ),
                        text=texts,
                        texttemplate="%{label}<br><b>%{text} bps</b>",
                        textfont=dict(size=14, color="black"),
                        hovertemplate="<b>%{label}</b><br>Stress PnL: %{color:.2f} bps<extra></extra>",
                        branchvalues="total"
                    )
                )
                
                fig_detail.update_layout(
                    height=450,
                    template="plotly_white",
                    paper_bgcolor="white",
                    plot_bgcolor="white",
                    margin=dict(t=10, b=10, l=10, r=10)
                )

        
                st.plotly_chart(fig_detail, use_container_width=True)
        
                # Download Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_detail.to_excel(writer, sheet_name="StressPnL By Strategy", index=False)
                output.seek(0)
        
                st.download_button(
                    label="📥 Download StressPnL By Strategy as Excel",
                    data=output.getvalue(),
                    file_name="stress_by_strategy.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_stress_by_strategy"
                )
            else:
                st.info("No data available for the selected combination of date, portfolio, and scenario.")
        
        
        # ------------------------------
        # Comparison Analysis
        # ------------------------------
        st.markdown("---")
        st.subheader("Comparison Analysis")
        
        # stesso box di Exposure
        selected_portfolio = st.selectbox(
            "Analysis portfolio",
            sel_ports,
            index=sel_ports.index("E7X") if "E7X" in sel_ports else 0,
            format_func=pretty_name,
            key="stress_comp_portfolio"
        )
        
        # Portfolio analizzato
        df_p = df[df["Portfolio"] == selected_portfolio]
        
        # Bucket = tutti gli altri
        df_b = df[df["Portfolio"] != selected_portfolio]
        
        bucket = (
            df_b.groupby("ScenarioName")["StressPnL"]
            .agg(
                bucket_median="median",
                q25=lambda x: x.quantile(0.25),
                q75=lambda x: x.quantile(0.75)
            )
            .reset_index()
        )
        
        plot_df = df_p.merge(bucket, on="ScenarioName")
        
        fig = go.Figure()
        
        # bande 25–75%
        for _, r in plot_df.iterrows():
            fig.add_trace(go.Scatter(
                x=[r.q25, r.q75],
                y=[r.ScenarioName] * 2,
                mode="lines",
                line=dict(width=14, color="rgba(0,0,255,0.25)"),
                showlegend=False
            ))
        
        # mediana bucket
        fig.add_trace(go.Scatter(
            x=plot_df["bucket_median"],
            y=plot_df["ScenarioName"],
            mode="markers",
            name="Bucket median",
            marker=dict(color="blue", size=10)
        ))
        
        # portfolio selezionato
        fig.add_trace(go.Scatter(
            x=plot_df["StressPnL"],
            y=plot_df["ScenarioName"],
            mode="markers",
            name=pretty_name(selected_portfolio),
            marker=dict(symbol="star", size=14, color="gold")
        ))
        
        fig.update_layout(
            height=600,
            template="plotly_white",
            xaxis_title="Stress PnL (bps)"
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        st.markdown(
            """
            <div style="display: flex; align-items: center;">
                <sub style="margin-right: 4px;">Note: the shaded areas</sub>
                <div style="width: 20px; height: 14px; background-color: rgba(0,0,255,0.25); margin: 0 4px 0 0; border: 1px solid rgba(0,0,0,0.1);"></div>
                <sub>represent the dispersion between the 25th and 75th percentile of the Bucket.</sub>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # Download Excel
        output = BytesIO()
        plot_df.to_excel(output, index=False)
        output.seek(0)
        
        pretty_portfolio_name = pretty_name(selected_portfolio)
        
        st.download_button(
            label=f"📥 Download {pretty_portfolio_name} vs Bucket Stress Test as Excel",
            data=output.getvalue(),
            file_name=f"{pretty_portfolio_name.replace(' ', '_').lower()}_vs_bucket_stress_test.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_comparison_stress_test"
        )

# ==================================================
# TAB — LEGENDA
# ==================================================
with tab_legenda:
    st.subheader("Series")
    st.dataframe(load_legenda("Portafogli", "A:C"), hide_index=True)
    st.subheader("Stress Scenarios")
    st.dataframe(load_legenda("Scenari", "A:C"), hide_index=True)
