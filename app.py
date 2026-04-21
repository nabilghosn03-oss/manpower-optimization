import streamlit as st
import pandas as pd
import pulp
import io
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
 
st.set_page_config(page_title="Manpower Optimization", layout="wide")
 
# ===== CUSTOM CSS =====
st.markdown("""
<style>
    /* Main background */
    .stApp { background-color: #f7f7f5; }
 
    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e8e8e4;
    }
    [data-testid="stSidebar"] .stMarkdown p { font-size: 13px; color: #666; }
 
    /* Metric cards */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e8e8e4;
        border-radius: 10px;
        padding: 16px 20px;
    }
    [data-testid="stMetricLabel"] { font-size: 12px !important; color: #888 !important; font-weight: 500 !important; text-transform: uppercase; letter-spacing: 0.05em; }
    [data-testid="stMetricValue"] { font-size: 22px !important; font-weight: 600 !important; color: #1a1a1a !important; }
 
    /* Expanders */
    [data-testid="stExpander"] {
        background-color: #ffffff;
        border: 1px solid #e8e8e4 !important;
        border-radius: 10px !important;
        margin-bottom: 6px;
    }
 
    /* Buttons */
    .stButton > button {
        background-color: #1D9E75 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        padding: 10px 20px !important;
        transition: background 0.2s !important;
    }
    .stButton > button:hover { background-color: #0F6E56 !important; }
 
    /* Download button */
    [data-testid="stDownloadButton"] > button {
        background-color: #ffffff !important;
        color: #1D9E75 !important;
        border: 1.5px solid #1D9E75 !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
    }
    [data-testid="stDownloadButton"] > button:hover {
        background-color: #E1F5EE !important;
    }
 
    hr { border: none; border-top: 1px solid #e8e8e4; margin: 24px 0; }
    h1 { color: #1a1a1a !important; font-weight: 600 !important; }
    h2, h3 { color: #1a1a1a !important; font-weight: 500 !important; }
 
    [data-testid="stAlert"] { border-radius: 8px !important; }
 
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
st.markdown("<p style='color:#888;margin-top:-12px;margin-bottom:8px;'>Optimize workforce allocation while minimizing costs</p>", unsafe_allow_html=True)
st.markdown("---")
 
# ===== SIDEBAR =====
st.sidebar.markdown("## ⚙️ Settings")
st.sidebar.markdown("---")
 
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
outsource_choice = st.sidebar.radio(
    "Which outsourced cost to use?",
    options=['current', 'improved'],
    index=0
)
COL_COST_OUTSOURCED = COL_OUTSOURCE_IMPROVED_COST if outsource_choice == 'improved' else COL_OUTSOURCE_CURRENT_COST
outsource_type = "Improved" if outsource_choice == 'improved' else "Current"
 
st.sidebar.markdown("---")
enforce_saudization = st.sidebar.checkbox("Enforce overall Saudization Rate?", value=True)
SAUDIZATION_RATE = None
if enforce_saudization:
    SAUDIZATION_RATE = st.sidebar.number_input(
        "Saudization Rate (decimal)",
        min_value=0.0, max_value=1.0, value=0.3, step=0.01
    )
 
st.sidebar.markdown("---")
can_fire_saudi = st.sidebar.checkbox("Allow reducing current Saudi headcount?", value=False)
 
run = st.sidebar.button("🚀 Run Optimization", type="primary", use_container_width=True)
 
# ===== RUN OPTIMIZATION =====
if run:
    prob = pulp.LpProblem("Manpower_Optimization", pulp.LpMinimize)
    S, N, O = [], [], []
 
    for i in range(len(data)):
        s = pulp.LpVariable(f'S_{i}', lowBound=0, cat='Integer')
        n = pulp.LpVariable(f'N_{i}', lowBound=0, cat='Integer')
        o = pulp.LpVariable(f'O_{i}', lowBound=0, cat='Integer')
        S.append(s); N.append(n); O.append(o)
 
        prob += s + n + o == data.iloc[i][COL_TOTAL_EMPLOYEES]
 
    prob += pulp.lpSum(
        data.iloc[i][COL_COST_SAUDI] * S[i] +
        data.iloc[i][COL_COST_NON_SAUDI] * N[i] +
        data.iloc[i][COL_COST_OUTSOURCED] * O[i]
        for i in range(len(data))
    )
 
    prob.solve(pulp.PULP_CBC_CMD(msg=0))
 
    results_data = []
    for i in range(len(data)):
        saudi = int(S[i].varValue)
        non_saudi = int(N[i].varValue)
        outsourced = int(O[i].varValue)
 
        results_data.append({
            'Job Family': data.iloc[i][COL_JOB_FAMILY],
            'In-House Saudi': saudi,
            'In-House Non-Saudi': non_saudi,
            'Outsourced': outsourced,
            'Total Employees': saudi + non_saudi + outsourced,
            'Cost - Saudi (SAR)': data.iloc[i][COL_COST_SAUDI] * saudi,
            'Cost - Non-Saudi (SAR)': data.iloc[i][COL_COST_NON_SAUDI] * non_saudi,
            'Cost - Outsourced (SAR)': data.iloc[i][COL_COST_OUTSOURCED] * outsourced,
        })
 
    results_df = pd.DataFrame(results_data)
 
    st.session_state.results_df = results_df
    st.session_state.total_cost = pulp.value(prob.objective)
    st.session_state.total_saudi_final = results_df['In-House Saudi'].sum()
    st.session_state.total_non_saudi_final = results_df['In-House Non-Saudi'].sum()
    st.session_state.total_outsourced_final = results_df['Outsourced'].sum()
    st.session_state.total_employees_final = results_df['Total Employees'].sum()
    st.session_state.saudization_achieved = (
        st.session_state.total_saudi_final /
        (st.session_state.total_saudi_final + st.session_state.total_non_saudi_final)
    ) * 100 if (st.session_state.total_saudi_final + st.session_state.total_non_saudi_final) > 0 else 0
 
# ===== DISPLAY RESULTS =====
if hasattr(st.session_state, 'results_df'):

    # ===== KPI (FIXED LAYOUT) =====
    c1, c2, c3 = st.columns(3)
    c4, c5, c6 = st.columns(3)

    with c1:
        st.metric("💰 Total Cost", f"SAR {st.session_state.total_cost:,.0f}")
    with c2:
        st.metric("👥 Total Employees", f"{st.session_state.total_employees_final:,}")
    with c3:
        st.metric("📊 Saudization Rate", f"{st.session_state.saudization_achieved:.1f}%")

    with c4:
        st.metric("🇸🇦 In-House Saudi", f"{st.session_state.total_saudi_final:,}")
    with c5:
        st.metric("🌍 In-House Non-Saudi", f"{st.session_state.total_non_saudi_final:,}")
    with c6:
        st.metric("🤝 Outsourced", f"{st.session_state.total_outsourced_final:,}")
 
    st.markdown("---")
 
    # ===== COST ANALYSIS =====
    st.markdown("### 📈 Cost Analysis")
 
    viz_col1, viz_col2 = st.columns(2)
 
    COLORS_METHOD = ['#1D9E75', '#378ADD', '#F0993B']
 
    with viz_col1:
        values_m = [
            st.session_state.total_saudi_final,
            st.session_state.total_non_saudi_final,
            st.session_state.total_outsourced_final
        ]
 
        fig_method = go.Figure(data=[go.Pie(
            labels=['Saudi', 'Non-Saudi', 'Outsourced'],
            values=values_m,
            textinfo='label+percent',
            hole=0.5,
            marker=dict(colors=COLORS_METHOD),
        )])
 
        fig_method.update_layout(
            title="Cost by sourcing method",
            annotations=[dict(
                text=f"Total<br>SAR {st.session_state.total_cost:,.0f}",
                x=0.5, y=0.5,
                showarrow=False,
                font=dict(size=14)
            )]
        )
 
        st.plotly_chart(fig_method, use_container_width=True)
 
    st.markdown("---")
 
    # ===== DETAILED TABLE =====
    st.markdown("### 📋 Detailed Allocation by Job Family")
 
    for _, row in st.session_state.results_df.iterrows():
        with st.expander(f"{row['Job Family']}"):
            st.write(row)
 
    st.markdown("---")
 
    # ===== DOWNLOAD =====
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.results_df.to_excel(writer, index=False)
 
    st.download_button(
        "Download Excel",
        data=output.getvalue(),
        file_name="results.xlsx"
    )