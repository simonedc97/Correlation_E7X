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

# ==================================================
# Sidebar
# ==================================================
st.sidebar.title("Controls")

# ==================================================
# DATA LOADING FUNCTIONS
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

    for sheet_name in xls.sheet_names:
        if "&&" in sheet_name:
            portfolio, scenario_name = sheet_name.split("&&", 1)
        else:
            portfolio, scenario_name = sheet_name, sheet_name

        df = pd.read_excel(xls, sheet_name=sheet_name)
        df = df.rename(columns={
            df.columns[0]: "Date",
            df.columns[2]: "Scenario",
            df.columns[4]: "StressPnL"
        })
        df["Date"] = pd.to_datetime(df["Date"])
        df["Portfolio"] = portfolio
        df["ScenarioName"] = scenario_name

        records.append(
            df[["Date", "Scenario", "StressPnL", "Portfolio", "ScenarioName"]]
        )

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
    df.columns = df.columns.str.strip()
    return df


@st.cache_data
def load_legenda_sheet(sheet_name, usecols):
    return pd.read_excel(
        "Legenda.xlsx",
        sheet_name=sheet_name,
        usecols=usecols
    )

# ==================================================
# LOAD DATA (ONLY E7X)
# ==================================================
corr = load_corr_data("corrE7X.xlsx")
stress_data = load_stress_data("stress_test_totE7X.xlsx")
exposure_data = load_exposure_data("E7X_Exposure.xlsx")

# ==================================================
# TAB — CORRELATION
# ==================================================
with tab_corr:
    st.session_state.current_tab = "Correlation"
    st.title("E7X Dynamic Asset Allocation vs Funds")

    # -----------------------------
    # Sidebar controls
    # -----------------------------
    st.sidebar.subheader("Date range (Correlation)")
    start_date, end_date = st.sidebar.date_input(
        "Select start and end date",
        value=(corr.index.min().date(), corr.index.max().date()),
        min_value=corr.index.min().date(),
        max_value=corr.index.max().date()
    )

    df = corr.loc[pd.to_datetime(start_date):pd.to_datetime(end_date)]

    st.sidebar.subheader("Series (Correlation)")
    selected_series = st.sidebar.multiselect(
        "Select series",
        options=df.columns.tolist(),
        default=df.columns.tolist()
    )

    if not selected_series:
        st.warning("Please select at least one series.")
        st.stop()

    # -----------------------------
    # Time series
    # -----------------------------
    fig_ts = go.Figure()
    palette = qualitative.Plotly
    color_map = {s: palette[i % len(palette)] for i, s in enumerate(selected_series)}

    for col in selected_series:
        fig_ts.add_trace(
            go.Scatter(
                x=df.index,
                y=df[col] * 100,
                mode="lines",
                name=col,
                line=dict(color=color_map[col]),
                hovertemplate="%{y:.2f}%<extra></extra>"
            )
        )

    fig_ts.update_layout(
        height=600,
        hovermode="x unified",
        template="plotly_white",
        yaxis=dict(ticksuffix="%"),
        xaxis_title="Date",
        yaxis_title="Correlation"
    )

    st.plotly_chart(fig_ts, use_container_width=True)

    # -----------------------------
    # Radar
    # -----------------------------
    st.subheader("Correlation Radar")

    snapshot = df.loc[df.index.max(), selected_series]
    mean_corr = df[selected_series].mean()

    fig_radar = go.Figure()
    fig_radar.add_trace(
        go.Scatterpolar(
            r=snapshot.values * 100,
            theta=snapshot.index,
            name="End date",
            line=dict(width=3)
        )
    )
    fig_radar.add_trace(
        go.Scatterpolar(
            r=mean_corr.values * 100,
            theta=mean_corr.index,
            name="Period mean",
            line=dict(dash="dot")
        )
    )

    fig_radar.update_layout(
        polar=dict(radialaxis=dict(range=[-100, 100], ticksuffix="%")),
        template="plotly_white",
        height=650
    )

    st.plotly_chart(fig_radar, use_container_width=True)

    # -----------------------------
    # Summary statistics
    # -----------------------------
    legenda = load_legenda_sheet("E7X", "A:C")
    ticker_to_name = dict(zip(legenda["Ticker"], legenda["Name"]))

    stats_df = pd.DataFrame(index=selected_series)
    stats_df.insert(0, "Name", [ticker_to_name.get(t, "") for t in selected_series])
    stats_df["Mean (%)"] = df[selected_series].mean() * 100
    stats_df["Min (%)"] = df[selected_series].min() * 100
    stats_df["Max (%)"] = df[selected_series].max() * 100

    st.dataframe(
        stats_df.style.format({
            "Mean (%)": "{:.2f}%",
            "Min (%)": "{:.2f}%",
            "Max (%)": "{:.2f}%"
        }),
        use_container_width=True
    )


# ==================================================
# TAB — STRESS TEST
# ==================================================
with tab_stress:
    st.session_state.current_tab = "StressTest"
    st.title("E7X Dynamic Asset Allocation vs Funds")

    st.sidebar.subheader("Date (Stress Test)")
    dates = sorted(stress_data["Date"].dropna().unique())
    selected_date = st.sidebar.selectbox(
        "Select date",
        [d.strftime("%Y/%m/%d") for d in dates],
        index=len(dates) - 1
    )
    selected_date = pd.to_datetime(selected_date)

    df = stress_data[stress_data["Date"] == selected_date]

    st.sidebar.subheader("Series (Stress Test)")
    portfolios = sorted(df["Portfolio"].unique())
    selected_portfolios = st.sidebar.multiselect(
        "Select portfolios",
        portfolios,
        default=portfolios
    )

    df = df[df["Portfolio"].isin(selected_portfolios)]

    st.sidebar.subheader("Scenarios (Stress Test)")
    scenarios = sorted(df["ScenarioName"].unique())
    selected_scenarios = st.sidebar.multiselect(
        "Select scenarios",
        scenarios,
        default=scenarios
    )

    df = df[df["ScenarioName"].isin(selected_scenarios)]

    # -----------------------------
    # Grouped bar
    # -----------------------------
    fig = go.Figure()
    palette = qualitative.Plotly

    for i, p in enumerate(selected_portfolios):
        d = df[df["Portfolio"] == p]
        fig.add_trace(
            go.Bar(
                x=d["ScenarioName"],
                y=d["StressPnL"],
                name=p,
                marker_color=palette[i % len(palette)]
            )
        )

    fig.update_layout(
        barmode="group",
        template="plotly_white",
        yaxis_title="Stress PnL (bps)",
        height=600
    )

    st.plotly_chart(fig, use_container_width=True)

# ==================================================
# TAB — EXPOSURE
# ==================================================
with tab_exposure:
    st.session_state.current_tab = "Exposure"
    st.title("E7X Dynamic Asset Allocation vs Funds")

    st.sidebar.subheader("Date (Exposure)")
    dates = sorted(exposure_data["Date"].unique())
    selected_date = st.sidebar.selectbox(
        "Select date",
        [d.strftime("%Y/%m/%d") for d in dates],
        index=len(dates) - 1
    )
    selected_date = pd.to_datetime(selected_date)

    df = exposure_data[exposure_data["Date"] == selected_date]

    st.sidebar.subheader("Series (Exposure)")
    portfolios = sorted(df["Portfolio"].unique())
    selected_portfolios = st.sidebar.multiselect(
        "Select portfolios",
        portfolios,
        default=portfolios
    )

    df = df[df["Portfolio"].isin(selected_portfolios)]

    metrics = ["Equity Exposure", "Duration", "Spread Duration"]

    df_plot = df.melt(
        id_vars=["Portfolio"],
        value_vars=metrics,
        var_name="Metric",
        value_name="Value"
    )

    fig = go.Figure()
    palette = qualitative.Plotly

    for i, p in enumerate(selected_portfolios):
        d = df_plot[df_plot["Portfolio"] == p]
        fig.add_trace(
            go.Bar(
                x=d["Metric"],
                y=d["Value"],
                name=p,
                marker_color=palette[i % len(palette)],
                text=d["Value"].round(1),
                textposition="auto"
            )
        )

    fig.update_layout(
        barmode="group",
        template="plotly_white",
        height=600
    )

    st.plotly_chart(fig, use_container_width=True)

# ==================================================
# TAB — LEGENDA
# ==================================================
with tab_legenda:
    st.session_state.current_tab = "Legenda"
    st.title("E7X Dynamic Asset Allocation vs Funds")

    legenda_main = load_legenda_sheet("E7X", "A:C")
    legenda_scenari = load_legenda_sheet("Scenari", "A:B")

    st.subheader("Series")
    st.dataframe(legenda_main, use_container_width=True, hide_index=True)

    st.markdown("---")

    st.subheader("Stress Test Scenarios")
    st.dataframe(legenda_scenari, use_container_width=True, hide_index=True)
