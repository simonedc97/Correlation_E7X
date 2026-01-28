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
# NAME MAP (TICKER → NAME)
# ==================================================
@st.cache_data
def load_name_map():
    legenda = load_legenda("E7X", "A:C")
    return dict(zip(legenda["Ticker"], legenda["Name"]))

NAME_MAP = load_name_map()


def pretty_name(ticker):
    """Fallback automatico: se non esiste in legenda → ticker"""
    return NAME_MAP.get(ticker, ticker)


# ==================================================
# LOAD DATA
# ==================================================
corr = load_corr_data("corrE7X_test.xlsx")
stress_data = load_stress_data("stress_test_totE7X.xlsx")
exposure_data = load_exposure_data("E7X_Exposure.xlsx")

# ==================================================
# TAB — CORRELATION
# ==================================================
with tab_corr:
    st.title("Dynamic Asset Allocation vs Funds")

    col_ctrl, col_plot = st.columns([1, 4])

    with col_ctrl:
        st.subheader("Controls")
        start, end = st.date_input(
            "Date range",
            (corr.index.min().date(), corr.index.max().date())
        )
        df = corr.loc[pd.to_datetime(start):pd.to_datetime(end)]

        selected = st.multiselect(
            "Select series",
            df.columns.tolist(),
            default=df.columns.tolist()
        )

    with col_plot:
        # ------------------------------
        # Time series
        # ------------------------------
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
            "📥 Download time series data",
            output.getvalue(),
            "correlation_time_series.xlsx"
        )

        # ------------------------------
        # Radar
        # ------------------------------
        st.subheader("Correlation Radar")

        snapshot_date = df.index.max()
        snapshot = df.loc[snapshot_date, selected]
        mean_corr = df[selected].mean()

        theta_names = [pretty_name(c) for c in snapshot.index]

        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(
            r=snapshot.values * 100,
            theta=theta_names,
            name=f"End date ({snapshot_date.date()})",
            line=dict(width=3)
        ))

        fig_radar.add_trace(go.Scatterpolar(
            r=mean_corr.values * 100,
            theta=theta_names,
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

        # ------------------------------
        # Summary stats
        # ------------------------------
        stats = pd.DataFrame(
            {
                "Name": [pretty_name(s) for s in selected],
                "Mean (%)": df[selected].mean() * 100,
                "Min (%)": df[selected].min() * 100,
                "Max (%)": df[selected].max() * 100,
            },
            index=selected
        )

        stats.index.name = "Ticker"

        st.dataframe(
            stats.reset_index(drop=True).style.format({
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
# TAB — STRESS TEST
# ==================================================
with tab_stress:
    st.title("Dynamic Asset Allocation vs Funds")

    col_ctrl, col_plot = st.columns([1, 4])

    with col_ctrl:
        st.subheader("Controls")
        dates = sorted(stress_data["Date"].unique())
        date = pd.to_datetime(st.selectbox(
            "Select date",
            [d.strftime("%Y/%m/%d") for d in dates],
            index=len(dates)-1
        ))

        df = stress_data[stress_data["Date"] == date]

        portfolios = df["Portfolio"].unique().tolist()
        sel_ports = st.multiselect(
            "Select portfolios",
            portfolios,
            default=portfolios
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

# ==================================================
# TAB — EXPOSURE
# ==================================================
with tab_exposure:
    st.title("Dynamic Asset Allocation vs Funds")

    col_ctrl, col_plot = st.columns([1, 4])

    with col_ctrl:
        st.subheader("Controls")
        dates = sorted(exposure_data["Date"].unique())
        date = pd.to_datetime(st.selectbox(
            "Select date",
            [d.strftime("%Y/%m/%d") for d in dates],
            index=len(dates)-1
        ))

        df = exposure_data[exposure_data["Date"] == date]

        ports = df["Portfolio"].unique().tolist()
        sel_ports = st.multiselect("Select portfolios", ports, default=ports)
        df = df[df["Portfolio"].isin(sel_ports)]

        metrics = ["Equity Exposure", "Duration", "Spread Duration"]

    with col_plot:
        df_plot = df.melt("Portfolio", metrics, "Metric", "Value")
        fig = go.Figure()

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

# ==================================================
# TAB — LEGENDA
# ==================================================
with tab_legenda:
    st.subheader("Series")
    st.dataframe(load_legenda("E7X", "A:B"), hide_index=True)
    st.subheader("Stress Scenarios")
    st.dataframe(load_legenda("Scenari", "A:C"), hide_index=True)
