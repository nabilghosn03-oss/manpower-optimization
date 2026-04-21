import streamlit as st
import pandas as pd
import pulp
import io
import plotly.graph_objects as go

st.set_page_config(page_title="Manpower Optimization", layout="wide")

# ===== CUSTOM CSS =====
st.markdown("""
<style>
    .stApp { background-color: #f7f7f5; }

    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e8e8e4;
    }

    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e8e8e4;
        border-radius: 10px;
        padding: 16px;
    }

    [data-testid="stMetricValue"] {
        font-size: 20px !important;
    }

    [data-testid="stPlotlyChart"] {
        background: #ffffff;
        border: 1px solid #e8e8e4;
        border-radius: 12px;
        padding: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
st.markdown("## 🎯 Manpower Optimization System")
st.markdown("Optimize workforce allocation while minimizing costs")
st.markdown("---")

# ===== SIDEBAR =====
st.sidebar.markdown("## ⚙️ Settings")
uploaded_file = st.sidebar.file_uploader("Upload Manpower_input.xlsx", type=["xlsx"])

if uploaded_file is None:
    st.info("Please upload your Excel file.")
    st.stop()

# ===== LOAD DATA =====
df = pd.read_excel(uploaded_file, header=None)
data = df.iloc[2:].reset_index(drop=True)

COL_JOB_FAMILY = 1
COL_TOTAL_EMPLOYEES = 2
COL_CURRENT_SAUDI = 3
COL_COST_SAUDI = 4
COL_COST_NON_SAUDI = 5
COL_OUTSOURCE_CURRENT_COST = 6
COL_OUTSOURCE_IMPROVED_COST = 7
COL_MAX_OUTSOURCE_RATIO = 8

for col in [COL_TOTAL_EMPLOYEES, COL_CURRENT_SAUDI, COL_COST_SAUDI,
            COL_COST_NON_SAUDI, COL_OUTSOURCE_CURRENT_COST,
            COL_OUTSOURCE_IMPROVED_COST, COL_MAX_OUTSOURCE_RATIO]:
    data[col] = pd.to_numeric(data[col], errors='coerce')

# ===== SETTINGS =====
outsource_choice = st.sidebar.radio(
    "Outsourced Cost Type",
    ["current", "improved"]
)

COL_COST_OUTSOURCED = (
    COL_OUTSOURCE_IMPROVED_COST if outsource_choice == "improved"
    else COL_OUTSOURCE_CURRENT_COST
)

enforce_saudization = st.sidebar.checkbox("Enforce Saudization", value=True)

SAUDIZATION_RATE = None
if enforce_saudization:
    SAUDIZATION_RATE = st.sidebar.number_input(
        "Saudization Rate",
        min_value=0.0,
        max_value=1.0,
        value=0.30,
        step=0.01
    )

can_fire_saudi = st.sidebar.checkbox("Allow reducing Saudi headcount", value=False)

run = st.sidebar.button("🚀 Run Optimization")

# ===== OPTIMIZATION =====
if run:
    prob = pulp.LpProblem("Manpower_Optimization", pulp.LpMinimize)

    S, N, O = [], [], []

    for i in range(len(data)):
        current_saudi = data.iloc[i][COL_CURRENT_SAUDI]
        total_employees = data.iloc[i][COL_TOTAL_EMPLOYEES]
        max_outsource_ratio = data.iloc[i][COL_MAX_OUTSOURCE_RATIO]

        saudi_lower_bound = 0 if can_fire_saudi else current_saudi

        s = pulp.LpVariable(f"S_{i}", lowBound=saudi_lower_bound, cat="Integer")
        n = pulp.LpVariable(f"N_{i}", lowBound=0, cat="Integer")
        o = pulp.LpVariable(f"O_{i}", lowBound=0, cat="Integer")

        S.append(s)
        N.append(n)
        O.append(o)

        prob += s + n + o == total_employees
        prob += o <= max_outsource_ratio * total_employees

    prob += pulp.lpSum(
        data.iloc[i][COL_COST_SAUDI] * S[i] +
        data.iloc[i][COL_COST_NON_SAUDI] * N[i] +
        data.iloc[i][COL_COST_OUTSOURCED] * O[i]
        for i in range(len(data))
    )

    if enforce_saudization:
        total_saudi = pulp.lpSum(S)
        total_inhouse = pulp.lpSum(S[i] + N[i] for i in range(len(data)))
        prob += total_saudi - SAUDIZATION_RATE * total_inhouse >= 0

    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    results = []

    for i in range(len(data)):
        saudi = int(S[i].varValue)
        non_saudi = int(N[i].varValue)
        outsourced = int(O[i].varValue)

        cost_saudi = saudi * data.iloc[i][COL_COST_SAUDI]
        cost_non_saudi = non_saudi * data.iloc[i][COL_COST_NON_SAUDI]
        cost_outsourced = outsourced * data.iloc[i][COL_COST_OUTSOURCED]

        results.append({
            "Job Family": data.iloc[i][COL_JOB_FAMILY],
            "In-House Saudi": saudi,
            "In-House Non-Saudi": non_saudi,
            "Outsourced": outsourced,
            "Total Employees": saudi + non_saudi + outsourced,
            "Cost - Saudi": cost_saudi,
            "Cost - Non-Saudi": cost_non_saudi,
            "Cost - Outsourced": cost_outsourced,
            "Total Cost": cost_saudi + cost_non_saudi + cost_outsourced
        })

    results_df = pd.DataFrame(results)

    total_cost = results_df["Total Cost"].sum()
    total_saudi = results_df["In-House Saudi"].sum()
    total_non_saudi = results_df["In-House Non-Saudi"].sum()
    total_outsourced = results_df["Outsourced"].sum()
    total_employees = results_df["Total Employees"].sum()
    saudization = total_saudi / (total_saudi + total_non_saudi) * 100

    # ===== KPI SECTION =====
    row1 = st.columns(3)
    row2 = st.columns(3)

    row1[0].metric("💰 Total Cost", f"SAR {total_cost:,.0f}")
    row1[1].metric("👥 Total Employees", f"{total_employees:,}")
    row1[2].metric("🇸🇦 Saudi", f"{total_saudi:,}")

    row2[0].metric("🌍 Non-Saudi", f"{total_non_saudi:,}")
    row2[1].metric("🤝 Outsourced", f"{total_outsourced:,}")
    row2[2].metric("📊 Saudization", f"{saudization:.1f}%")

    st.markdown("---")

    # ===== COST ANALYSIS =====
    col1, col2 = st.columns(2)

    with col1:
        fig1 = go.Figure(data=[go.Pie(
            labels=["Saudi", "Non-Saudi", "Outsourced"],
            values=[
                results_df["Cost - Saudi"].sum(),
                results_df["Cost - Non-Saudi"].sum(),
                results_df["Cost - Outsourced"].sum()
            ],
            hole=0.45,
            textinfo="percent+label",
            textposition="inside"
        )])

        fig1.update_layout(title="Cost by sourcing method")
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        top_costs = results_df.sort_values("Total Cost", ascending=False).head(10)

        fig2 = go.Figure(data=[go.Pie(
            labels=top_costs["Job Family"],
            values=top_costs["Total Cost"],
            hole=0.45,
            textinfo="percent+label",
            textposition="inside"
        )])

        fig2.update_layout(title="Cost by job family")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")

    # ===== IMPROVED HEADCOUNT CHART =====
    st.markdown("### 👥 Headcount Distribution by Job Family")

    df_plot = results_df.sort_values("Total Employees", ascending=True)

    fig_bar = go.Figure()

    fig_bar.add_trace(go.Bar(
        y=df_plot["Job Family"],
        x=df_plot["In-House Saudi"],
        name="Saudi",
        orientation="h"
    ))

    fig_bar.add_trace(go.Bar(
        y=df_plot["Job Family"],
        x=df_plot["In-House Non-Saudi"],
        name="Non-Saudi",
        orientation="h"
    ))

    fig_bar.add_trace(go.Bar(
        y=df_plot["Job Family"],
        x=df_plot["Outsourced"],
        name="Outsourced",
        orientation="h"
    ))

    fig_bar.update_layout(
        barmode="stack",
        height=max(500, len(df_plot) * 35),
        title="Headcount by Job Family",
        xaxis_title="Employees"
    )

    st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---")

    # ===== TABLE =====
    st.dataframe(results_df, use_container_width=True)

    # ===== DOWNLOAD =====
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        results_df.to_excel(writer, index=False)

    st.download_button(
        "📥 Download Excel",
        output.getvalue(),
        "Manpower_Results.xlsx"
    )