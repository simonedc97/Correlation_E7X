import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.colors import qualitative
from io import BytesIO
from streamlit_plotly_events import plotly_events
import numpy as np

# ==================================================
# Page config
# ==================================================
st.set_page_config(layout="wide")

# ==================================================
# Tabs
# ==================================================
tab_corr, tab_stress, tab_exposure, tab_legenda = st.tabs(
    ["Correlation", "Stress Test", "Exposure", "Legend"]
)

# ==================================================
# DATA LOADING
# ==================================================
@st.cache_data
def load_corr_data(path):
    df = pd.read_excel(path, sheet_name="Correlation Clean")
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0])
    return df.set_index(df.columns[0]).sort_index()


@st.cache_data
def load_stress_data(path):
    xls = pd.ExcelFile(path)
    records = []

    for sheet in xls.sheet_names:
        portfolio, scenario = sheet.split("&&", 1) if "&&" in sheet else (sheet, sheet)

        df = pd.read_excel(xls, sheet_name=sheet)
        df = df.rename(columns={
            df.columns[0]: "Date",
            df.columns[2]: "Scenario",
            df.columns[4]: "StressPnL"
        })

        df["Date"] = pd.to_datetime(df["Date"])
        df["Portfolio"] = portfolio
        df["ScenarioName"] = scenario

        records.append(df[["Date", "Scenario", "StressPnL", "Portfolio", "ScenarioName"]])

    return pd.concat(records, ignore_index=True)


@st.cache_data
def load_exposure_data(path):
    df = pd.read_excel(path, sheet_name="MeasuresSeries")
    df = df.rename(columns={
        df.columns[0]: "Date",
        df.columns[3]: "Portfolio",
        df.columns[4]: "Equity Exposure",
        df.columns[5]: "Duration",
        df.columns[6]: "Spread Duration"
    })
    df["Date"] = pd.to_datetime(df["Date"])
    return df


@st.cache_data
def load_legenda(sheet, cols):
    return pd.read_excel("Legenda.xlsx", sheet_name=sheet, usecols=cols)


# ==================================================
# NAME MAP (Ticker → Name)
# ==================================================
@st.cache_data
def load_name_map():
    legenda = load_legenda("E7X", "A:C")
    return dict(zip(legenda["Ticker"], legenda["Name"]))


NAME_MAP = load_name_map()


def pretty_name(x):
    return NAME_MAP.get(x, x)


# ==================================================
# LOAD DATA
# ==================================================
corr = load_corr_data("corrE7X.xlsx")
stress_data = load_stress_data("stress_test_totE7X.xlsx")
exposure_data = load_exposure_data("E7X_Exposure.xlsx")

# ==================================================
# TAB — CORRELATION
# ==================================================
with tab_corr:
    st.title("Dynamic Asset Allocation vs Funds")

    col_ctrl, col_plot = st.columns([0.7, 4.3])

    with col_ctrl:
        st.subheader("Controls")

        start, end = st.date_input(
            "Date range",
            (corr.index.min().date(), corr.index.max().date())
        )

        df = corr.loc[pd.to_datetime(start):pd.to_datetime(end)]

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
            yaxis_title="Correlation (%)"
        )

        st.plotly_chart(fig, use_container_width=True)

        # Download time series
        output = BytesIO()
        (df[selected] * 100).to_excel(output)
        st.download_button(
            "📥 Download Correlation Time Series as Excel",
            output.getvalue(),
            "correlation_time_series.xlsx"
        )

        
        # Radar
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
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[-100, 100],
                    ticksuffix="%"
                )
            ),
            template="plotly_white",
            height=650
        )

        st.plotly_chart(fig_radar, use_container_width=True)

        st.subheader("Summary Statistics")
        
        stats_df = pd.DataFrame(index=selected)
        
        # Colonna Name (pretty)
        stats_df.insert(
            0,
            "Name",
            [pretty_name(s) for s in selected]
        )
        
        # Statistiche
        stats_df["Mean (%)"] = df[selected].mean() * 100
        stats_df["Min (%)"] = df[selected].min() * 100
        stats_df["Min Date"] = [
            df[col][df[col] == df[col].min()].index.max()
            for col in selected
        ]
        stats_df["Max (%)"] = df[selected].max() * 100
        stats_df["Max Date"] = [
            df[col][df[col] == df[col].max()].index.max()
            for col in selected
        ]
        
        # Formattazione date
        stats_df["Min Date"] = pd.to_datetime(stats_df["Min Date"]).dt.strftime("%d/%m/%Y")
        stats_df["Max Date"] = pd.to_datetime(stats_df["Max Date"]).dt.strftime("%d/%m/%Y")
        
        # Visualizzazione
        st.dataframe(
            stats_df.style.format({
                "Mean (%)": "{:.2f}%",
                "Min (%)": "{:.2f}%",
                "Max (%)": "{:.2f}%"
            }),
            use_container_width=True
        )
      
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            stats_df.to_excel(
                writer,
                sheet_name="Summary Statistics",
                index=False
            )
        
        st.download_button(
            label="📥 Download Summary Statistics as Excel",
            data=output.getvalue(),
            file_name="summary_statistics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_summary_stats"
        )
# ==================================================
# TAB — STRESS TEST
# ==================================================
with tab_stress:
    st.title("Dynamic Asset Allocation vs Funds")

    col_ctrl, col_plot = st.columns([0.7, 4.3])

    with col_ctrl:
        st.subheader("Controls")

        dates = sorted(stress_data["Date"].unique())
        date = pd.to_datetime(
            st.selectbox(
                "Select date",
                [d.strftime("%Y/%m/%d") for d in dates],
                index=len(dates) - 1
            )
        )

        df = stress_data[stress_data["Date"] == date]

        portfolios = df["Portfolio"].unique().tolist()
        sel_ports = st.multiselect(
            "Select portfolios",
            portfolios,
            default=portfolios,
            format_func=pretty_name
        )

        df = df[df["Portfolio"].isin(sel_ports)]

        scenarios = df["ScenarioName"].unique().tolist()
        sel_scen = st.multiselect(
            "Select scenarios",
            scenarios,
            default=scenarios
        )

        df = df[df["ScenarioName"].isin(sel_scen)]

    with col_plot:
        # ------------------------------
        # Bar chart
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
        
        @st.cache_data
        def load_stress_bystrat(path):
            xls = pd.ExcelFile(path)
            records = []
        
            for sheet in xls.sheet_names:
                portfolio, scenario = sheet.split("&&", 1) if "&&" in sheet else (sheet, sheet)
        
                df = pd.read_excel(xls, sheet_name=sheet)
                df["Portfolio"] = portfolio
                df["ScenarioName"] = scenario
        
                records.append(df)
        
            return pd.concat(records, ignore_index=True)  # <-- qui finisce la funzione
        
        # === fuori dalla funzione ===
        stress_bystrat = load_stress_bystrat("stress_test_bystrat.xlsx")
        
        with st.expander("Expand for Stress Test analysis by strategy", expanded=False):
            
            # Select portfolio
            clicked_portfolio = st.selectbox(
                "Analysis Portfolio",
                sel_ports,
                format_func=pretty_name
            )
            
            # Select scenario
            clicked_scenario = st.selectbox(
                "Scenario",
                sel_scen
            )
            
            # Filtra dati
            df_detail = stress_bystrat[
                (stress_bystrat["Portfolio"] == clicked_portfolio) &
                (stress_bystrat["ScenarioName"] == clicked_scenario) &
                (stress_bystrat.iloc[:, 0] != "Total")  # esclude la riga Total
            ]
            
            if not df_detail.empty:
                # Colori tipo heatmap: rossi per negativi, verdi per positivi
                
                
                vals = df_detail["Stress PnL"].values
                # Normalizza tra 0 e 1 per il mapping
                max_abs = np.max(np.abs(vals)) if np.max(np.abs(vals)) != 0 else 1
                colors = [
                    f"rgba({int(255*abs(min(0,v)/max_abs))},{int(255*max(0,v)/max_abs)},0,0.8)"
                    for v in vals
                ]
                # In alternativa: rosso = [255,0,0], verde = [0,255,0] proporzionale al valore
                
                df_tm = df_detail.copy()
                df_tm["size"] = df_tm["Stress PnL"].abs().clip(lower=0.01)
                
                root_label = f"{pretty_name(clicked_portfolio)} - {clicked_scenario}"
                
                labels = [root_label] + df_tm.iloc[:, 0].tolist()
                parents = [""] + [root_label] * len(df_tm)
                values = [df_tm["size"].sum()] + df_tm["size"].tolist()
                colors = ["white"] + df_tm["Stress PnL"].tolist()
                
                # testo: vuoto per root, valori per le strategy
                texts = [""] + df_tm["Stress PnL"].round(2).astype(str).tolist()
                
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
                        # 🔥 il root (text="") non mostra nulla
                        texttemplate="%{label}<br><b>%{text} bps</b>",
                        textfont=dict(size=14, color="black"),
                        hovertemplate=(
                            "<b>%{label}</b><br>"
                            "Stress PnL: %{color:.2f} bps<extra></extra>"
                        ),
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

    
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_detail.to_excel(writer, sheet_name="Stress Test PnL By Strategy", index=False)
                
                st.download_button(
                    label="📥 Download Stress Test PnL By Strategy",
                    data=output.getvalue(),
                    file_name="stress_test_by_strat.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_stress_by_strategy"
                )

        # ------------------------------
        # Comparison Analysis
        # ------------------------------
        st.markdown("---")
        st.subheader("Comparison Analysis")

        selected_portfolio = st.selectbox(
            "Analysis portfolio",
            sel_ports,
            index=sel_ports.index("E7X") if "E7X" in sel_ports else 0,
            format_func=pretty_name
        )

        df_p = df[df["Portfolio"] == selected_portfolio]
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

        for _, r in plot_df.iterrows():
            fig.add_trace(go.Scatter(
                x=[r.q25, r.q75],
                y=[r.ScenarioName] * 2,
                mode="lines",
                line=dict(width=14, color="rgba(0,0,255,0.25)"),
                showlegend=False
            ))

        fig.add_trace(go.Scatter(
            x=plot_df["bucket_median"],
            y=plot_df["ScenarioName"],
            mode="markers",
            name="Bucket median",
            marker=dict(color="blue", size=10)
        ))

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

        output = BytesIO()
        plot_df.to_excel(output, index=False)

        pretty_portfolio_name = pretty_name(selected_portfolio)

        st.download_button(
            label=f"📥 Download {pretty_portfolio_name} vs Bucket Stress Test By Strategy as Excel",
            data=output,
            file_name=f"{pretty_portfolio_name.replace(' ', '_').lower()}_vs_bucket_stress test.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==================================================
# TAB — EXPOSURE
# ==================================================
with tab_exposure:
    st.title("Dynamic Asset Allocation vs Funds")

    col_ctrl, col_plot = st.columns([0.7, 4.3])

    with col_ctrl:
        st.subheader("Controls")

        dates = sorted(exposure_data["Date"].unique())
        date = pd.to_datetime(
            st.selectbox(
                "Select date",
                [d.strftime("%Y/%m/%d") for d in dates],
                index=len(dates) - 1
            )
        )

        df = exposure_data[exposure_data["Date"] == date]

        ports = df["Portfolio"].unique().tolist()
        sel_ports = st.multiselect(
            "Select portfolios",
            ports,
            default=ports,
            format_func=pretty_name
        )

        df = df[df["Portfolio"].isin(sel_ports)]

        metrics = ["Equity Exposure", "Duration", "Spread Duration"]

    with col_plot:
        st.subheader("Exposure")
    
        fig = go.Figure()
        df_plot = df.melt("Portfolio", metrics, "Metric", "Value")
    
        for p in sel_ports:
            d = df_plot[df_plot["Portfolio"] == p]
            fig.add_trace(go.Bar(
                x=d["Metric"],
                y=d["Value"],
                name=pretty_name(p)
            ))
    
        fig.update_layout(
            barmode="group",
            height=600,
            template="plotly_white"
        )
    
        st.plotly_chart(fig, use_container_width=True)
    
        # ------------------------------
        # Download Excel (stesso pattern)
        # ------------------------------
        df_download = df_plot[df_plot["Portfolio"].isin(sel_ports)].copy()
    
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_download.to_excel(
                writer,
                sheet_name="Exposure",
                index=False
            )
    
        st.download_button(
            label="📥 Download Exposure as Excel",
            data=output.getvalue(),
            file_name="exposure.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_exposure"
        )

        # ------------------------------
        # Comparison Analysis
        # ------------------------------
        st.markdown("---")
        st.subheader("Comparison Analysis")

        selected_portfolio = st.selectbox(
            "Analysis portfolio",
            sel_ports,
            index=sel_ports.index("E7X") if "E7X" in sel_ports else 0,
            format_func=pretty_name
        )

        df_p = df[df["Portfolio"] == selected_portfolio][metrics]
        df_b = df[df["Portfolio"] != selected_portfolio][metrics]

        bucket = df_b.agg(
            ["median", lambda x: x.quantile(0.25), lambda x: x.quantile(0.75)]
        ).T

        bucket.columns = ["bucket_median", "q25", "q75"]

        comp = (
            df_p.T
            .rename(columns={df_p.index[0]: selected_portfolio})
            .reset_index()
            .rename(columns={"index": "Metric"})
            .merge(bucket.reset_index().rename(columns={"index": "Metric"}))
        )

        fig = go.Figure()

        for _, r in comp.iterrows():
            fig.add_trace(go.Scatter(
                x=[r.q25, r.q75],
                y=[r.Metric] * 2,
                mode="lines",
                line=dict(width=14, color="rgba(0,0,255,0.3)"),
                showlegend=False
            ))

        fig.add_trace(go.Scatter(
            x=comp["bucket_median"],
            y=comp["Metric"],
            mode="markers",
            name="Bucket median",
            marker=dict(color="blue", size=10)
        ))

        fig.add_trace(go.Scatter(
            x=comp[selected_portfolio],
            y=comp["Metric"],
            mode="markers",
            name=pretty_name(selected_portfolio),
            marker=dict(symbol="star", size=14, color="gold")
        ))

        fig.update_layout(
            height=600,
            template="plotly_white"
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

        output = BytesIO()
        comp.to_excel(output, index=False)
        output.seek(0)   # <-- QUESTO MANCAVA
        
        pretty_portfolio_name = pretty_name(selected_portfolio)
        
        st.download_button(
            label=f"📥 Download {pretty_portfolio_name} vs Bucket Exposure as Excel",
            data=output,
            file_name=f"{pretty_portfolio_name.replace(' ', '_').lower()}_vs_bucket_exposure.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# ==================================================
# TAB — LEGENDA
# ==================================================
with tab_legenda:
    st.subheader("Series")
    st.dataframe(load_legenda("E7X", "A:B"), hide_index=True)

    st.subheader("Stress Scenarios")
    st.dataframe(load_legenda("Scenari", "A:C"), hide_index=True)
