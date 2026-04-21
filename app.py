import streamlit as st
import pandas as pd
import pulp
import io
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Manpower Optimization", layout="wide")

# ===== CUSTOM CSS (UNCHANGED FROM YOUR CODE) =====
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
        padding: 16px 20px;
    }

    [data-testid="stMetricValue"] {
        font-size: 22px !important;
        font-weight: 600 !important;
    }

    [data-testid="stExpander"] {
        background-color: #ffffff;
        border: 1px solid #e8e8e4 !important;
        border-radius: 10px !important;
    }

    .stButton > button {
        background-color: #1D9E75 !important;
        color: white !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
st.markdown("## 🎯 Manpower Optimization System")
st.markdown("<p style='color:#888;margin-top:-10px;'>Optimize workforce allocation while minimizing costs</p>", unsafe_allow_html=True)
st.markdown("---")

# ===== SIDEBAR (UNCHANGED — AS YOU REQUESTED) =====
st.sidebar.markdown("## ⚙️ Settings")
st.sidebar.markdown("---")

st.sidebar.markdown("### 📁 Input Data")
uploaded_file = st.sidebar.file_uploader("Upload Manpower_input.xlsx", type=['xlsx'])

if uploaded_file is None:
    st.info("👈 Please upload your Excel file in the sidebar to get started.")
    st.stop()

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

st.sidebar.markdown("---")
st.sidebar.markdown("### 1️⃣ Outsourced Cost Type")
outsource_choice = st.sidebar.radio(
    "Which outsourced cost to use?",
    options=['current', 'improved'],
    format_func=lambda x: "Current outsourced cost" if x == 'current' else "Improved outsourced cost",
    index=0
)

COL_COST_OUTSOURCED = COL_OUTSOURCE_IMPROVED_COST if outsource_choice == 'improved' else COL_OUTSOURCE_CURRENT_COST
outsource_type = "Improved" if outsource_choice == 'improved' else "Current"

st.sidebar.markdown("---")
st.sidebar.markdown("### 2️⃣ Saudization Rate")
enforce_saudization = st.sidebar.checkbox("Enforce overall Saudization Rate?", value=True)

SAUDIZATION_RATE = None
if enforce_saudization:
    SAUDIZATION_RATE = st.sidebar.number_input(
        "Saudization Rate (decimal)",
        min_value=0.0, max_value=1.0, value=0.3, step=0.01, format="%.2f"
    )
    st.sidebar.success(f"✓ Target: {SAUDIZATION_RATE*100:.1f}%")
else:
    st.sidebar.info("No Saudization constraint")

st.sidebar.markdown("---")
st.sidebar.markdown("### 3️⃣ Saudi Labor Policy")
can_fire_saudi = st.sidebar.checkbox("Allow reducing current Saudi headcount?", value=False)

run = st.sidebar.button("🚀 Run Optimization", type="primary", use_container_width=True)

# ===== OPTIMIZATION =====
if run:
    prob = pulp.LpProblem("Manpower_Optimization", pulp.LpMinimize)

    S, N, O = [], [], []

    for i in range(len(data)):
        current_saudi = data.iloc[i][COL_CURRENT_SAUDI]
        total_employees = data.iloc[i][COL_TOTAL_EMPLOYEES]
        max_outsource_ratio = data.iloc[i][COL_MAX_OUTSOURCE_RATIO]

        saudi_lb = 0 if can_fire_saudi else current_saudi

        s = pulp.LpVariable(f"S_{i}", lowBound=saudi_lb, cat="Integer")
        n = pulp.LpVariable(f"N_{i}", lowBound=0, cat="Integer")
        o = pulp.LpVariable(f"O_{i}", lowBound=0, cat="Integer")

        S.append(s); N.append(n); O.append(o)

        prob += s + n + o == total_employees
        prob += o <= max_outsource_ratio * total_employees

    prob += pulp.lpSum(
        data.iloc[i][COL_COST_SAUDI] * S[i] +
        data.iloc[i][COL_COST_NON_SAUDI] * N[i] +
        data.iloc[i][COL_COST_OUTSOURCED] * O[i]
        for i in range(len(data))
    )

    if enforce_saudization:
        prob += pulp.lpSum(S) >= SAUDIZATION_RATE * pulp.lpSum(S[i] + N[i] for i in range(len(data)))

    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    results = []

    for i in range(len(data)):
        saudi = int(S[i].varValue)
        non_saudi = int(N[i].varValue)
        outsourced = int(O[i].varValue)

        results.append({
            "Job Family": data.iloc[i][COL_JOB_FAMILY],
            "In-House Saudi": saudi,
            "In-House Non-Saudi": non_saudi,
            "Outsourced": outsourced,
            "Total Employees": saudi + non_saudi + outsourced,
            "Cost - Saudi": saudi * data.iloc[i][COL_COST_SAUDI],
            "Cost - Non-Saudi": non_saudi * data.iloc[i][COL_COST_NON_SAUDI],
            "Cost - Outsourced": outsourced * data.iloc[i][COL_COST_OUTSOURCED],
        })

    results_df = pd.DataFrame(results)
    results_df["Total Cost"] = results_df[
        ["Cost - Saudi", "Cost - Non-Saudi", "Cost - Outsourced"]
    ].sum(axis=1)

    # ===== SESSION STORAGE (YOUR ORIGINAL LOGIC KEPT) =====
    st.session_state.results_df = results_df
    st.session_state.total_cost = results_df["Total Cost"].sum()
    st.session_state.total_saudi_final = results_df["In-House Saudi"].sum()
    st.session_state.total_non_saudi_final = results_df["In-House Non-Saudi"].sum()
    st.session_state.total_outsourced_final = results_df["Outsourced"].sum()
    st.session_state.total_employees_final = results_df["Total Employees"].sum()
    st.session_state.saudization_achieved = (
        st.session_state.total_saudi_final /
        (st.session_state.total_saudi_final + st.session_state.total_non_saudi_final)
    ) * 100

    st.success("Optimization completed!")

# ===== RESULTS =====
if "results_df" in st.session_state:

    # ===== KPI (FIXED OVERFLOW ISSUE) =====
    c1, c2, c3 = st.columns(3)
    c4, c5, c6 = st.columns(3)

    c1.metric("💰 Total Cost", f"SAR {st.session_state.total_cost:,.0f}")
    c2.metric("👥 Employees", f"{st.session_state.total_employees_final:,}")
    c3.metric("🇸🇦 Saudi", f"{st.session_state.total_saudi_final:,}")

    c4.metric("🌍 Non-Saudi", f"{st.session_state.total_non_saudi_final:,}")
    c5.metric("🤝 Outsourced", f"{st.session_state.total_outsourced_final:,}")
    c6.metric("📊 Saudization", f"{st.session_state.saudization_achieved:.1f}%")

    st.markdown("---")

    # ===== COST PIE (FIXED % DISPLAY) =====
    col1, col2 = st.columns(2)

    with col1:
        fig = go.Figure(data=[go.Pie(
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

        st.plotly_chart(fig, use_container_width=True)

    with col2:
        top = results_df.sort_values("Total Cost", ascending=False).head(10)

        fig2 = go.Figure(data=[go.Pie(
            labels=top["Job Family"],
            values=top["Total Cost"],
            hole=0.45,
            textinfo="percent+label",
            textposition="inside"
        )])

        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")

    # ===== IMPROVED HEADCOUNT CHART =====
    st.markdown("### 👥 Headcount Distribution")

    dfp = results_df.sort_values("Total Employees")

    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=dfp["Job Family"],
        x=dfp["In-House Saudi"],
        name="Saudi",
        orientation="h"
    ))

    fig.add_trace(go.Bar(
        y=dfp["Job Family"],
        x=dfp["In-House Non-Saudi"],
        name="Non-Saudi",
        orientation="h"
    ))

    fig.add_trace(go.Bar(
        y=dfp["Job Family"],
        x=dfp["Outsourced"],
        name="Outsourced",
        orientation="h"
    ))

    fig.update_layout(
        barmode="stack",
        height=max(500, len(dfp) * 35),
        xaxis_title="Headcount",
        yaxis_title=""
    )

    st.plotly_chart(fig, use_container_width=True)

    # ===== TABLE (UNCHANGED) =====
    st.markdown("### 📋 Detailed Breakdown")
    st.dataframe(results_df, use_container_width=True)