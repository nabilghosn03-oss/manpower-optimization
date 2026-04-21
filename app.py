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
 
    /* Section dividers */
    hr { border: none; border-top: 1px solid #e8e8e4; margin: 24px 0; }
 
    /* Headers */
    h1 { color: #1a1a1a !important; font-weight: 600 !important; }
    h2, h3 { color: #1a1a1a !important; font-weight: 500 !important; }
 
    /* Info / success / warning boxes */
    [data-testid="stAlert"] { border-radius: 8px !important; }
 
    /* Chart containers */
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
 
st.sidebar.markdown("### 📁 Input Data")
uploaded_file = st.sidebar.file_uploader("Upload Manpower_input.xlsx", type=['xlsx'])
 
if uploaded_file is None:
    st.info("👈 Please upload your Excel file in the sidebar to get started.")
    st.stop()
 
# Read the Excel file
try:
    df = pd.read_excel(uploaded_file, header=None)
except Exception as e:
    st.error(f"Error reading Excel file: {e}")
    st.stop()
 
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
    st.sidebar.info("No Saudization Rate constraint")
 
st.sidebar.markdown("---")
st.sidebar.markdown("### 3️⃣ Saudi Labor Policy")
can_fire_saudi = st.sidebar.checkbox("Allow reducing current Saudi headcount?", value=False)
if can_fire_saudi:
    st.sidebar.warning("Saudi labor can be reduced below current levels")
else:
    st.sidebar.info("Saudi labor minimum constraint applied")
 
st.sidebar.markdown("---")
run = st.sidebar.button("🚀 Run Optimization", type="primary", use_container_width=True)
 
# ===== RUN OPTIMIZATION =====
if run:
    with st.spinner("Optimizing workforce allocation..."):
        try:
            prob = pulp.LpProblem("Manpower_Optimization", pulp.LpMinimize)
            S, N, O = [], [], []
 
            for i in range(len(data)):
                current_saudi = data.iloc[i][COL_CURRENT_SAUDI]
                total_employees = data.iloc[i][COL_TOTAL_EMPLOYEES]
                max_outsource_ratio = data.iloc[i][COL_MAX_OUTSOURCE_RATIO]
                saudi_lower_bound = 0 if can_fire_saudi else current_saudi
 
                s = pulp.LpVariable(f'S_{i}', lowBound=saudi_lower_bound, cat='Integer')
                n = pulp.LpVariable(f'N_{i}', lowBound=0, cat='Integer')
                o = pulp.LpVariable(f'O_{i}', lowBound=0, cat='Integer')
                S.append(s); N.append(n); O.append(o)
 
                prob += s + n + o == total_employees, f"Total_Employees_{i}"
                prob += o <= max_outsource_ratio * total_employees, f"Max_Outsource_Ratio_{i}"
 
            prob += pulp.lpSum(
                data.iloc[i][COL_COST_SAUDI] * S[i] +
                data.iloc[i][COL_COST_NON_SAUDI] * N[i] +
                data.iloc[i][COL_COST_OUTSOURCED] * O[i]
                for i in range(len(data))
            )
 
            if enforce_saudization:
                prob += pulp.lpSum(S) >= SAUDIZATION_RATE * pulp.lpSum(S[i] + N[i] for i in range(len(data))), "Saudization_Rate"
 
            prob.solve(pulp.PULP_CBC_CMD(msg=0))
 
            if pulp.LpStatus[prob.status] == 'Optimal':
                results_data = []
                for i in range(len(data)):
                    saudi = int(S[i].varValue)
                    non_saudi = int(N[i].varValue)
                    outsourced = int(O[i].varValue)
                    cost_saudi = data.iloc[i][COL_COST_SAUDI] * saudi
                    cost_non_saudi = data.iloc[i][COL_COST_NON_SAUDI] * non_saudi
                    cost_outsourced = data.iloc[i][COL_COST_OUTSOURCED] * outsourced
                    results_data.append({
                        'Job Family': data.iloc[i][COL_JOB_FAMILY],
                        'In-House Saudi': saudi,
                        'In-House Non-Saudi': non_saudi,
                        'Outsourced': outsourced,
                        'Total Employees': saudi + non_saudi + outsourced,
                        'Cost - Saudi (SAR)': cost_saudi,
                        'Cost - Non-Saudi (SAR)': cost_non_saudi,
                        'Cost - Outsourced (SAR)': cost_outsourced,
                        'Total Cost (SAR)': cost_saudi + cost_non_saudi + cost_outsourced
                    })
 
                results_df = pd.DataFrame(results_data)
                total_saudi_final = sum(int(S[i].varValue) for i in range(len(data)))
                total_non_saudi_final = sum(int(N[i].varValue) for i in range(len(data)))
                total_outsourced_final = sum(int(O[i].varValue) for i in range(len(data)))
                total_employees_final = total_saudi_final + total_non_saudi_final + total_outsourced_final
                total_inhouse_final = total_saudi_final + total_non_saudi_final
                saudization_achieved = (total_saudi_final / total_inhouse_final * 100) if total_inhouse_final > 0 else 0
 
                st.session_state.results_df = results_df
                st.session_state.total_cost = pulp.value(prob.objective)
                st.session_state.total_saudi_final = total_saudi_final
                st.session_state.total_non_saudi_final = total_non_saudi_final
                st.session_state.total_outsourced_final = total_outsourced_final
                st.session_state.total_employees_final = total_employees_final
                st.session_state.saudization_achieved = saudization_achieved
                st.session_state.optimization_status = "Optimal"
                st.session_state.total_cost_saudi = results_df['Cost - Saudi (SAR)'].sum()
                st.session_state.total_cost_non_saudi = results_df['Cost - Non-Saudi (SAR)'].sum()
                st.session_state.total_cost_outsourced = results_df['Cost - Outsourced (SAR)'].sum()
                st.success("✅ Optimization completed successfully!")
            else:
                st.error(f"❌ Optimization failed: {pulp.LpStatus[prob.status]}")
 
        except Exception as e:
            st.error(f"Error during optimization: {str(e)}")
 
# ===== DISPLAY RESULTS =====
if hasattr(st.session_state, 'optimization_status'):
 
    # ===== KPI METRICS (ONLY EDITED SECTION) =====
    row1 = st.columns(3)
    row2 = st.columns(3)

    with row1[0]:
        st.metric("💰 Total Cost", f"SAR {st.session_state.total_cost:,.0f}")

    with row1[1]:
        st.metric("👥 Total Employees", f"{st.session_state.total_employees_final:,}")

    with row1[2]:
        st.metric("📊 Saudization Rate", f"{st.session_state.saudization_achieved:.1f}%")

    with row2[0]:
        st.metric("🇸🇦 In-House Saudi", f"{st.session_state.total_saudi_final:,}")

    with row2[1]:
        st.metric("🌍 In-House Non-Saudi", f"{st.session_state.total_non_saudi_final:,}")

    with row2[2]:
        st.metric("🤝 Outsourced", f"{st.session_state.total_outsourced_final:,}")
 
    st.markdown("---")
 
    # ===== EVERYTHING BELOW UNCHANGED =====