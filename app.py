import streamlit as st
import pandas as pd
import pulp
import io
import plotly.graph_objects as go
import plotly.express as px
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Manpower Optimization", layout="wide")

st.title("🎯 Manpower Optimization System")
st.markdown("Optimize your workforce allocation while minimizing costs")

# Sidebar for file upload
st.sidebar.header("📁 Upload Input Data")
uploaded_file = st.sidebar.file_uploader("Upload your Manpower_input.xlsx", type=['xlsx'])

if uploaded_file is None:
    st.info("👈 Please upload your Excel file in the sidebar to get started")
    st.stop()

# Read the Excel file
try:
    df = pd.read_excel(uploaded_file, header=None)
except Exception as e:
    st.error(f"Error reading Excel file: {e}")
    st.stop()

# Column mapping
headers = df.iloc[1].tolist()
data = df.iloc[2:].reset_index(drop=True)

COL_JOB_FAMILY = 1
COL_TOTAL_EMPLOYEES = 2
COL_CURRENT_SAUDI = 3
COL_COST_SAUDI = 4
COL_COST_NON_SAUDI = 5
COL_OUTSOURCE_CURRENT_COST = 6
COL_OUTSOURCE_IMPROVED_COST = 7
COL_MAX_OUTSOURCE_RATIO = 8

# Convert to numeric
data[COL_TOTAL_EMPLOYEES] = pd.to_numeric(data[COL_TOTAL_EMPLOYEES], errors='coerce')
data[COL_CURRENT_SAUDI] = pd.to_numeric(data[COL_CURRENT_SAUDI], errors='coerce')
data[COL_COST_SAUDI] = pd.to_numeric(data[COL_COST_SAUDI], errors='coerce')
data[COL_COST_NON_SAUDI] = pd.to_numeric(data[COL_COST_NON_SAUDI], errors='coerce')
data[COL_OUTSOURCE_CURRENT_COST] = pd.to_numeric(data[COL_OUTSOURCE_CURRENT_COST], errors='coerce')
data[COL_OUTSOURCE_IMPROVED_COST] = pd.to_numeric(data[COL_OUTSOURCE_IMPROVED_COST], errors='coerce')
data[COL_MAX_OUTSOURCE_RATIO] = pd.to_numeric(data[COL_MAX_OUTSOURCE_RATIO], errors='coerce')

# ===== USER INPUTS IN SIDEBAR =====
st.sidebar.header("⚙️ Optimization Settings")

# Question 1: Outsourcing costs
st.sidebar.subheader("1️⃣ Outsourced Cost Type")
outsource_choice = st.sidebar.radio(
    "Which outsourced cost would you like to use?",
    options=['current', 'improved'],
    format_func=lambda x: "Current Outsourced Cost" if x == 'current' else "Improved Outsourced Cost",
    index=0
)

if outsource_choice == 'improved':
    COL_COST_OUTSOURCED = COL_OUTSOURCE_IMPROVED_COST
    outsource_type = "Improved"
else:
    COL_COST_OUTSOURCED = COL_OUTSOURCE_CURRENT_COST
    outsource_type = "Current"

st.sidebar.success(f"✓ Using {outsource_type} Outsourced Costs")

# Question 2: Saudization rate
st.sidebar.subheader("2️⃣ Saudization Rate")
enforce_saudization = st.sidebar.checkbox("Enforce overall Saudization Rate?", value=True)

SAUDIZATION_RATE = None
if enforce_saudization:
    SAUDIZATION_RATE = st.sidebar.number_input(
        "Enter Saudization Rate (decimal format)",
        min_value=0.0,
        max_value=1.0,
        value=0.3,
        step=0.01,
        format="%.2f"
    )
    st.sidebar.success(f"✓ Saudization Rate: {SAUDIZATION_RATE*100:.2f}%")
else:
    st.sidebar.info("✓ No Saudization Rate constraint")

# Question 3: Can fire Saudi labor
st.sidebar.subheader("3️⃣ Saudi Labor")
can_fire_saudi = st.sidebar.checkbox("Can fire current Saudi labor?", value=False)

if can_fire_saudi:
    st.sidebar.warning("⚠️ Saudi labor can be reduced below current levels")
else:
    st.sidebar.info("✓ Current Saudi labor minimum constraint applied")

# ===== RUN OPTIMIZATION =====
if st.sidebar.button("🚀 Run Optimization", key="optimize_btn", type="primary"):
    with st.spinner("⏳ Optimizing your workforce allocation..."):
        try:
            # Create problem
            prob = pulp.LpProblem("Manpower_Optimization", pulp.LpMinimize)

            # Variables
            S = []  # Saudi in-house
            N = []  # Non-Saudi in-house
            O = []  # Outsourced

            for i in range(len(data)):
                current_saudi = data.iloc[i][COL_CURRENT_SAUDI]
                total_employees = data.iloc[i][COL_TOTAL_EMPLOYEES]
                max_outsource_ratio = data.iloc[i][COL_MAX_OUTSOURCE_RATIO]
                
                # Saudi in-house: set lower bound based on whether we can fire them
                if can_fire_saudi:
                    saudi_lower_bound = 0
                else:
                    saudi_lower_bound = current_saudi
                
                s = pulp.LpVariable(f'S_{i}', lowBound=saudi_lower_bound, cat='Integer')
                n = pulp.LpVariable(f'N_{i}', lowBound=0, cat='Integer')
                o = pulp.LpVariable(f'O_{i}', lowBound=0, cat='Integer')
                
                S.append(s)
                N.append(n)
                O.append(o)
                
                # Constraint 1: Total employees must equal input
                prob += s + n + o == total_employees, f"Total_Employees_{i}"
                
                # Constraint 2: Outsourcing ratio constraint
                prob += o <= max_outsource_ratio * total_employees, f"Max_Outsource_Ratio_{i}"

            # Objective function: minimize total cost
            prob += pulp.lpSum(data.iloc[i][COL_COST_SAUDI] * S[i] + 
                               data.iloc[i][COL_COST_NON_SAUDI] * N[i] + 
                               data.iloc[i][COL_COST_OUTSOURCED] * O[i] 
                               for i in range(len(data)))

            # Constraint 3: Saudization constraint (only if enforced)
            if enforce_saudization:
                total_saudi = pulp.lpSum(S)
                total_inhouse = pulp.lpSum(S[i] + N[i] for i in range(len(data)))
                prob += total_saudi >= SAUDIZATION_RATE * total_inhouse, "Saudization_Rate"

            # Solve the problem
            prob.solve(pulp.PULP_CBC_CMD(msg=0))

            # Process results
            if pulp.LpStatus[prob.status] == 'Optimal':
                total_cost = pulp.value(prob.objective)
                
                # Create results dataframe
                results_data = []
                for i in range(len(data)):
                    job_family = data.iloc[i][COL_JOB_FAMILY]
                    saudi = int(S[i].varValue)
                    non_saudi = int(N[i].varValue)
                    outsourced = int(O[i].varValue)
                    total = saudi + non_saudi + outsourced
                    
                    cost_saudi = data.iloc[i][COL_COST_SAUDI] * saudi
                    cost_non_saudi = data.iloc[i][COL_COST_NON_SAUDI] * non_saudi
                    cost_outsourced = data.iloc[i][COL_COST_OUTSOURCED] * outsourced
                    total_cost_job = cost_saudi + cost_non_saudi + cost_outsourced
                    
                    results_data.append({
                        'Job Family': job_family,
                        'In-House Saudi': saudi,
                        'In-House Non-Saudi': non_saudi,
                        'Outsourced': outsourced,
                        'Total Employees': total,
                        'Cost - Saudi': cost_saudi,
                        'Cost - Non-Saudi': cost_non_saudi,
                        'Cost - Outsourced': cost_outsourced,
                        'Total Cost': total_cost_job
                    })
                
                results_df = pd.DataFrame(results_data)
                
                # Calculate totals
                total_saudi_final = sum(int(S[i].varValue) for i in range(len(data)))
                total_non_saudi_final = sum(int(N[i].varValue) for i in range(len(data)))
                total_outsourced_final = sum(int(O[i].varValue) for i in range(len(data)))
                total_employees_final = total_saudi_final + total_non_saudi_final + total_outsourced_final
                total_inhouse_final = total_saudi_final + total_non_saudi_final
                saudization_achieved = (total_saudi_final / total_inhouse_final * 100) if total_inhouse_final > 0 else 0
                
                # Cost totals
                total_cost_saudi = results_df['Cost - Saudi'].sum()
                total_cost_non_saudi = results_df['Cost - Non-Saudi'].sum()
                total_cost_outsourced = results_df['Cost - Outsourced'].sum()
                
                # Store in session state
                st.session_state.results_df = results_df
                st.session_state.total_cost = total_cost
                st.session_state.total_saudi_final = total_saudi_final
                st.session_state.total_non_saudi_final = total_non_saudi_final
                st.session_state.total_outsourced_final = total_outsourced_final
                st.session_state.total_employees_final = total_employees_final
                st.session_state.saudization_achieved = saudization_achieved
                st.session_state.optimization_status = "Optimal"
                st.session_state.total_cost_saudi = total_cost_saudi
                st.session_state.total_cost_non_saudi = total_cost_non_saudi
                st.session_state.total_cost_outsourced = total_cost_outsourced
                
                st.success("✅ Optimization completed successfully!")
            else:
                st.error(f"❌ Optimization failed: {pulp.LpStatus[prob.status]}")
        
        except Exception as e:
            st.error(f"Error during optimization: {str(e)}")

# ===== DISPLAY RESULTS =====
if hasattr(st.session_state, 'optimization_status'):
    st.markdown("---")
    st.header("📊 Results")
    
    # Key metrics in columns
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            "💰 Total Cost",
            f"${st.session_state.total_cost:,.2f}"
        )
    
    with col2:
        st.metric(
            "👥 Total Employees",
            f"{st.session_state.total_employees_final:,}"
        )
    
    with col3:
        st.metric(
            "🇸🇦 In-House Saudi",
            f"{st.session_state.total_saudi_final:,}"
        )
    
    with col4:
        st.metric(
            "🌍 In-House Non-Saudi",
            f"{st.session_state.total_non_saudi_final:,}"
        )
    
    with col5:
        st.metric(
            "🤝 Outsourced",
            f"{st.session_state.total_outsourced_final:,}"
        )
    
    # Saudization Rate
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric(
            "Saudization Rate Achieved",
            f"{st.session_state.saudization_achieved:.2f}%"
        )
    
    with col2:
        if enforce_saudization:
            st.metric(
                "Saudization Rate Required",
                f"{SAUDIZATION_RATE*100:.2f}%"
            )
    
    # ===== VISUALIZATIONS =====
    st.markdown("---")
    st.subheader("📈 Cost Analysis")
    
    # Create two columns for pie charts
    viz_col1, viz_col2 = st.columns(2)
    
    # Pie chart 1: Cost by Sourcing Method
    with viz_col1:
        st.markdown("**Cost Split by Sourcing Method**")
        cost_by_method = {
            'In-House Saudi': st.session_state.total_cost_saudi,
            'In-House Non-Saudi': st.session_state.total_cost_non_saudi,
            'Outsourced': st.session_state.total_cost_outsourced
        }
        
        fig_method = go.Figure(data=[go.Pie(
            labels=list(cost_by_method.keys()),
            values=list(cost_by_method.values()),
            marker=dict(colors=['#2E7D32', '#1976D2', '#F57C00'])
        )])
        fig_method.update_layout(height=400, showlegend=True)
        st.plotly_chart(fig_method, use_container_width=True)
    
    # Pie chart 2: Cost by Job Family (with "Other" category)
    with viz_col2:
        st.markdown("**Cost Split by Job Family**")
        
        # Sort by cost descending and group smaller ones as "Other"
        results_sorted = st.session_state.results_df.copy()
        results_sorted['Total Cost'] = pd.to_numeric(
            results_sorted['Total Cost'].astype(str).str.replace(',', ''), 
            errors='coerce'
        )
        results_sorted = results_sorted.sort_values('Total Cost', ascending=False)
        
        total_cost_all = results_sorted['Total Cost'].sum()
        threshold = total_cost_all * 0.05  # Top 5% threshold
        
        top_families = results_sorted[results_sorted['Total Cost'] >= threshold]
        other_cost = results_sorted[results_sorted['Total Cost'] < threshold]['Total Cost'].sum()
        
        job_families = list(top_families['Job Family'])
        costs = list(top_families['Total Cost'])
        
        if other_cost > 0:
            job_families.append("Other")
            costs.append(other_cost)
            other_list = list(results_sorted[results_sorted['Total Cost'] < threshold]['Job Family'])
            other_text = f"Other ({len(other_list)} job families)"
            job_families[-1] = other_text
        
        fig_family = go.Figure(data=[go.Pie(
            labels=job_families,
            values=costs,
            textposition='inside',
            textinfo='label+percent'
        )])
        fig_family.update_layout(height=400, showlegend=True)
        st.plotly_chart(fig_family, use_container_width=True)
    
    # Detailed results table with expandable rows
    st.markdown("---")
    st.subheader("📋 Detailed Allocation by Job Family")
    st.markdown("*Click the expanders below to see cost breakdown per job family*")
    
    # Display as expandable sections
    for idx, row in st.session_state.results_df.iterrows():
        with st.expander(f"**{row['Job Family']}** - ${row['Total Cost']:,.2f}"):
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Saudi", int(row['In-House Saudi']))
            with col2:
                st.metric("Non-Saudi", int(row['In-House Non-Saudi']))
            with col3:
                st.metric("Outsourced", int(row['Outsourced']))
            with col4:
                st.metric("Total", int(row['Total Employees']))
            
            # Cost breakdown bar chart for this job family
            costs_breakdown = {
                'In-House Saudi': float(row['Cost - Saudi']),
                'In-House Non-Saudi': float(row['Cost - Non-Saudi']),
                'Outsourced': float(row['Cost - Outsourced'])
            }
            
            fig_breakdown = go.Figure(data=[
                go.Bar(
                    x=list(costs_breakdown.keys()),
                    y=list(costs_breakdown.values()),
                    marker_color=['#2E7D32', '#1976D2', '#F57C00']
                )
            ])
            fig_breakdown.update_layout(
                height=300,
                yaxis_title="Cost ($)",
                xaxis_title="Sourcing Type",
                showlegend=False
            )
            st.plotly_chart(fig_breakdown, use_container_width=True)
    
    # Download Excel
    st.markdown("---")
    st.subheader("📥 Download Results")
    
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        st.session_state.results_df.to_excel(writer, sheet_name='Optimization Results', index=False)
        
        # Summary sheet
        summary_data = {
            'Metric': ['Total Cost', 'Total Employees', 'In-House Saudi', 'In-House Non-Saudi', 'Outsourced', 
                      'Saudization Rate Achieved (%)', 'Optimization Status', 'Outsourced Cost Type', 'Can Fire Saudi', 'Saudization Enforced'],
            'Value': [f'{st.session_state.total_cost:,.2f}', st.session_state.total_employees_final, 
                     st.session_state.total_saudi_final, st.session_state.total_non_saudi_final, 
                     st.session_state.total_outsourced_final, f'{st.session_state.saudization_achieved:.2f}', 
                     st.session_state.optimization_status, outsource_type, 'Yes' if can_fire_saudi else 'No', 
                     'Yes' if enforce_saudization else 'No']
        }
        
        if enforce_saudization:
            summary_data['Metric'].insert(6, 'Saudization Rate Required (%)')
            summary_data['Value'].insert(6, f'{SAUDIZATION_RATE*100:.2f}')
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    output_buffer.seek(0)
    st.download_button(
        label="📊 Download Results as Excel",
        data=output_buffer.getvalue(),
        file_name="Manpower_Optimization_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ===== SIDEBAR INFO =====
st.sidebar.markdown("---")
st.sidebar.markdown("### 📋 About")
st.sidebar.markdown("""
This tool optimizes your workforce allocation to minimize costs while respecting:
- Maximum outsourcing ratios
- Saudi labor retention policies
- Saudization rate requirements
""")

st.sidebar.markdown("---")
st.sidebar.markdown("### 🌐 Access from Any Laptop")
st.sidebar.markdown("""
**Current Status:** Local only (localhost:8501)

**To make it accessible from anywhere:**
1. Deploy to Streamlit Cloud (free)
2. Push code to GitHub
3. Visit share.streamlit.io

**Alternatively:**
- Use ngrok for temporary sharing
- Deploy to AWS/Heroku/Azure
""")

