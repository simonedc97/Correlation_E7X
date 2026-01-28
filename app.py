import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.colors import qualitative
from io import BytesIO

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

st.sidebar.title("Controls")

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
# LOAD DATA (E7X)
# ==================================================
corr = load_corr_data("corrE7X.xlsx")
stress_data = load_stress_data("stress_test_totE7X.xlsx")
exposure_data = load_exposure_data("E7X_Exposure.xlsx")

# ==================================================
# TAB — CORRELATION
# ==================================================
with tab_corr:
    st.title("E7X Dynamic Asset Allocation vs Funds")

    # Sidebar
    st.sidebar.subheader("Date range (Correlation)")
    start, end = st.sidebar.date_input(
        "Select range",
        (corr.index.min().date(), corr.index.max().date())
    )

    df = corr.loc[pd.to_datetime(start):pd.to_datetime(end)]

    st.sidebar.subheader("Series (Correlation)")
    selected = st.sidebar.multiselect(
        "Select series",
        df.columns.tolist(),
        default=df.columns.tolist()
    )

    # Time series
    fig = go.Figure()
    palette = qualitative.Plotly

    for i, c in enumerate(selected):
        fig.add_trace(go.Scatter(
            x=df.index,
            y=df[c] * 100,
            name=c,
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
        "📥 Download time series data",
        output.getvalue(),
        "correlation_time_series.xlsx"
    )

    # Radar
    snapshot = df.loc[df.index.max(), selected]
    mean_corr = df[selected].mean()

    fig_radar = go.Figure()
    fig_radar.add_trace(go.Scatterpolar(
        r=snapshot * 100,
        theta=snapshot.index,
        name="End date"
    ))
    fig_radar.add_trace(go.Scatterpolar(
        r=mean_corr * 100,
        theta=mean_corr.index,
        name="Mean",
        line=dict(dash="dot")
    ))

    fig_radar.update_layout(
        polar=dict(radialaxis=dict(range=[-100, 100])),
        height=600
    )
    st.plotly_chart(fig_radar, use_container_width=True)

    # Summary stats
    legenda = load_legenda("E7X", "A:C")
    name_map = dict(zip(legenda["Ticker"], legenda["Name"]))

    stats = pd.DataFrame(index=selected)
    stats.insert(0, "Name", [name_map.get(s, "") for s in selected])
    stats["Mean (%)"] = df[selected].mean() * 100
    stats["Min (%)"] = df[selected].min() * 100
    stats["Max (%)"] = df[selected].max() * 100

    st.dataframe(
        stats.style.format({
            "Mean (%)": "{:.2f}%",
            "Min (%)": "{:.2f}%",
            "Max (%)": "{:.2f}%"
        }),
        use_container_width=True
    )

    output = BytesIO()
    stats.to_excel(output)
    st.download_button(
        "📥 Download summary statistics",
        output.getvalue(),
        "correlation_summary.xlsx"
    )

# ==================================================
# TAB — STRESS TEST (WITH COMPARISON)
# ==================================================
with tab_stress:
    st.title("E7X Dynamic Asset Allocation vs Funds")

    st.sidebar.subheader("Date (Stress Test)")
    dates = sorted(stress_data["Date"].unique())
    date = pd.to_datetime(st.sidebar.selectbox(
        "Select date",
        [d.strftime("%Y/%m/%d") for d in dates],
        index=len(dates)-1
    ))

    df = stress_data[stress_data["Date"] == date]

    st.sidebar.subheader("Series (Stress Test)")
    portfolios = df["Portfolio"].unique().tolist()
    sel_ports = st.sidebar.multiselect(
        "Select portfolios",
        portfolios,
        default=portfolios
    )
    df = df[df["Portfolio"].isin(sel_ports)]

    st.sidebar.subheader("Scenarios")
    scenarios = df["ScenarioName"].unique().tolist()
    sel_scen = st.sidebar.multiselect(
        "Select scenarios",
        scenarios,
        default=scenarios
    )
    df = df[df["ScenarioName"].isin(sel_scen)]

    # Bar chart
    fig = go.Figure()
    for i, p in enumerate(sel_ports):
        d = df[df["Portfolio"] == p]
        fig.add_trace(go.Bar(
            x=d["ScenarioName"],
            y=d["StressPnL"],
            name=p
        ))

    fig.update_layout(
        barmode="group",
        height=600,
        template="plotly_white",
        yaxis_title="Stress PnL (bps)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # Download
    output = BytesIO()
    df[["Portfolio", "ScenarioName", "StressPnL"]].to_excel(output, index=False)
    st.download_button(
        "📥 Download Stress Test data",
        output.getvalue(),
        "stress_test.xlsx"
    )

    # -----------------------------
    # Comparison Analysis
    # -----------------------------
    st.markdown("---")
    st.subheader("Comparison Analysis")

    selected_portfolio = st.selectbox(
        "Analysis portfolio",
        sel_ports,
        index=sel_ports.index("E7X") if "E7X" in sel_ports else 0
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
            y=[r.ScenarioName]*2,
            mode="lines",
            line=dict(width=14, color="rgba(255,0,0,0.3)"),
            showlegend=False
        ))

    fig.add_trace(go.Scatter(
        x=plot_df["bucket_median"],
        y=plot_df["ScenarioName"],
        mode="markers",
        name="Bucket median",
        marker=dict(color="red")
    ))

    fig.add_trace(go.Scatter(
        x=plot_df["StressPnL"],
        y=plot_df["ScenarioName"],
        mode="markers",
        name=selected_portfolio,
        marker=dict(symbol="star", size=14)
    ))

    fig.update_layout(
        height=600,
        template="plotly_white",
        xaxis_title="Stress PnL (bps)"
    )

    st.plotly_chart(fig, use_container_width=True)

    output = BytesIO()
    plot_df.to_excel(output, index=False)
    st.download_button(
        f"📥 Download {selected_portfolio} vs Bucket",
        output.getvalue(),
        f"{selected_portfolio}_vs_bucket_stress.xlsx"
    )

# ==================================================
# TAB — EXPOSURE (WITH COMPARISON)
# ==================================================
with tab_exposure:
    st.title("E7X Dynamic Asset Allocation vs Funds")

    st.sidebar.subheader("Date (Exposure)")
    dates = sorted(exposure_data["Date"].unique())
    date = pd.to_datetime(st.sidebar.selectbox(
        "Select date",
        [d.strftime("%Y/%m/%d") for d in dates],
        index=len(dates)-1
    ))

    df = exposure_data[exposure_data["Date"] == date]

    st.sidebar.subheader("Series (Exposure)")
    ports = df["Portfolio"].unique().tolist()
    sel_ports = st.sidebar.multiselect("Select portfolios", ports, default=ports)
    df = df[df["Portfolio"].isin(sel_ports)]

    metrics = ["Equity Exposure", "Duration", "Spread Duration"]

    # Bar
    df_plot = df.melt("Portfolio", metrics, "Metric", "Value")
    fig = go.Figure()

    for i, p in enumerate(sel_ports):
        d = df_plot[df_plot["Portfolio"] == p]
        fig.add_trace(go.Bar(
            x=d["Metric"],
            y=d["Value"],
            name=p
        ))

    fig.update_layout(
        barmode="group",
        height=600,
        template="plotly_white"
    )
    st.plotly_chart(fig, use_container_width=True)

    output = BytesIO()
    df.to_excel(output, index=False)
    st.download_button(
        "📥 Download Exposure data",
        output.getvalue(),
        "exposure.xlsx"
    )

    # Comparison
    st.markdown("---")
    st.subheader("Comparison Analysis")

    selected_portfolio = st.selectbox(
        "Analysis portfolio",
        sel_ports,
        index=sel_ports.index("E7X") if "E7X" in sel_ports else 0
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
            y=[r.Metric]*2,
            mode="lines",
            line=dict(width=14, color="rgba(0,0,255,0.3)"),
            showlegend=False
        ))

    fig.add_trace(go.Scatter(
        x=comp["bucket_median"],
        y=comp["Metric"],
        mode="markers",
        name="Bucket median"
    ))

    fig.add_trace(go.Scatter(
        x=comp[selected_portfolio],
        y=comp["Metric"],
        mode="markers",
        marker=dict(symbol="star", size=14),
        name=selected_portfolio
    ))

    fig.update_layout(
        height=600,
        template="plotly_white"
    )
    st.plotly_chart(fig, use_container_width=True)

    output = BytesIO()
    comp.to_excel(output, index=False)
    st.download_button(
        f"📥 Download {selected_portfolio} vs Bucket Exposure",
        output.getvalue(),
        f"{selected_portfolio}_vs_bucket_exposure.xlsx"
    )

# ==================================================
# TAB — LEGENDA
# ==================================================
with tab_legenda:
    st.title("E7X Dynamic Asset Allocation vs Funds")

    st.subheader("Series")
    st.dataframe(load_legenda("E7X", "A:C"), hide_index=True)

    st.subheader("Stress Scenarios")
    st.dataframe(load_legenda("Scenari", "A:B"), hide_index=True)
