import streamlit as st
import pandas as pd
import pulp
import io
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Manpower Optimization", layout="wide")

# ===== EMBEDDED MAPPING DICTIONARIES =====
# These mappings are embedded in the code and will be used to standardize employee data

ACTIVITY_MAPPING = {
    'Production': 'Factory',
    'Mgic Store': 'Showroom',
    'Logistics': 'Installation',
    'Maintenance': 'Factory',
    'Idle Saudi Labor': 'Idle Saudi Labor',
    'Accounts': 'Head Office',
    'Administration': 'Head Office',
    'Head Office': 'Head Office',
    'Procurement': 'Head Office',
    'Marketing': 'Head Office',
    'Cost controll': 'Head Office',
    'Planning': 'Head Office',
    'Contracting': 'Installation',
    'Quarries': 'Quarries',
    'AD.F - VIP South Land': 'Installation',
    'Shamiya Makkah Haram Extension -Piazza P1150': 'Installation',
    'Shamiya Service Buildings Deck': 'Installation',
    'OBEROI \nHOTEL\nRiyad': 'Installation',
    'Shopdrawing': 'Head Office',
    'MATAF EXTENSION': 'Installation',
    'KSAU - HS - Mental Health Hospital': 'Installation',
    'Shura Central Hotel 01 – Stone Surfaces': 'Installation',
    'CUC Tunnels': 'Installation',
    'King Faisal Specialist Hospital Project - Jeddah': 'Installation',
    'Al-Khaldiyah Palace': 'Installation',
    'KSAU - HS - Administration Building': 'Installation',
    'HARAM HIGH SPEED RAIL (Makkah Station)': 'Installation',
    'Al Murjan Villa': 'Installation',
    'Installation Supply': 'Installation',
    'Jeddah Showroom': 'Showroom',
    'MAC - Laheq Sales Center': 'Installation',
    'INTERGRATIVE WELLNESS core-30859030301': 'Installation',
    'MGIC - Laheq Sales Center': 'Installation',
    'Showroom': 'Showroom',
    'Export': 'Head Office',
    'Installation': 'Installation',
    'Factory': 'Factory',
}

PROFESSION_MAPPING = {
    'Executive Management': 'Executive Management',
    'Operator': 'Operator',
    'Officer': 'Officer',
    'Foreman': 'Foreman',
    'Super-Intendant': 'Superintendent',
    'Labor': 'Labor',
    'Inspector': 'Inspector',
    'Skilled Labor': 'Skilled Labor',
    'Coordinator': 'Coordinator',
    'Controller': 'Controller',
    'Supervisor': 'Supervisor',
    'Purchaser': 'Purchaser',
    'Clerk': 'Clerk',
    'Document Controller': 'Controller',
    'Technician': 'Technician',
    'Driver': 'Driver',
    'Manager': 'Manager',
    'Assistant': 'Assistant',
    'Engineer': 'Engineer',
    'Analyst': 'Analyst',
    'Store Keeper': 'Store Keeper',
    'Messenger': 'Messenger',
    'Surveyor': 'Surveyor',
    'Representative': 'Representative',
    'Draftsman': 'Draftsman',
    'Time Keeper': 'Time Keeper',
    'Treasurer': 'Treasurer',
    'Specialist': 'Administration',
    'Accountant': 'Administration',
    'Administrator': 'Administration',
    'Cashier': 'Administration',
    'Security Guard': 'Security Guard',
    'Senior Geologist': 'Geologist',
    'Geologist': 'Geologist',
    'Junior Technician': 'Technician',
    'Marble Finishing Technician': 'Technician',
    'Marble Installation Technician': 'Technician',
    'Semi Skilled Labor': 'Semi Skilled Labor',
    'Junior Foreman': 'Foreman',
    'Installation Foreman': 'Foreman',
    'Junior Surveyor': 'Surveyor',
    'Quality Inspector': 'Inspector',
    'Maintenance Technician': 'Technician',
    'Machine Operator': 'Operator',
    'Junior Operator': 'Operator',
    'Senior Technician': 'Technician',
    'Junior Inspector': 'Inspector',
    'Junior Clerk': 'Clerk',
    'Junior Engineer': 'Engineer',
    'Senior Operator': 'Operator',
    'Drill Technician': 'Technician',
    'Finisher': 'Finisher',
    'Fixer': 'Semi Skilled Labor',
    'Labor Polishing': 'Skilled Labor',
    'Support': 'Labor',
    'Scissor Operator': 'Operator',
    'Labor Grouting': 'Skilled Labor',
    'Labor Safety': 'Labor',
    'Night Security Guard': 'Security Guard',
    'Polisher': 'Skilled Labor',
    'Camp Labor': 'Labor',
    'Forklift Operator': 'Operator',
    'Store Labor': 'Labor',
    'SAFETY FOREMAN': 'Foreman',
    'Electrician Labor': 'Labor',
    'Safety': 'Safety Officer',
    'Welder Operator': 'Technical Operator',
    'MARBLE FIXER': 'Semi Skilled Labor',
    'Asst.Timekeeper': 'Time Keeper',
    'Motor Rewinding': 'Skilled Labor',
    'Electrician': 'Skilled Labor',
    'MECHANIC': 'Skilled Labor',
    'LATHEMACHINE OPERATOR': 'Technical Operator',
    'WELDER': 'Skilled Labor',
    'INDUSTRIAL MECHANIC': 'Skilled Labor',
    'PRODUCTION FOREMAN': 'Foreman',
    'Driller': 'Skilled Labor',
    'EQUIPMENT OPERATOR': 'Operator',
    'DATA ENTRY': 'Administration',
    'MECHANICAL ENGINEER': 'Engineer',
    'SPARE PARTS FOREMAN': 'Foreman',
    'TECHNICAL DATA OPERATOR': 'Technical Operator',
    'Lathe machine': 'Technical Operator',
    'Heavy equipment operator': 'Technical Operator',
    'SALES REPRESENTATIVE': 'Representative',
    'Erector': 'Semi Skilled Labor',
    'Help': 'Labor',
    'MARBLE FINISHER': 'Finisher',
    'Support Foreman': 'Foreman',
    'Safety/Security': 'Security Guard',
    'Camp/Security': 'Security Guard',
    'Grouting Labor': 'Labor',
    'Office Boy': 'Labor',
    'Timekeeper': 'Time Keeper',
    'Finishing Foreman': 'Foreman',
    'ASST. TIMEKEEPER': 'Time Keeper',
    'QUARRIES SUPVERSIOR': 'Supervisor',
    'SAFETY OFFICER': 'Safety Officer',
    'Forklift opp.': 'Operator',
    'PAINTER': 'Skilled Labor',
    'MATERIAL PLANNING MANAGER': 'Management',
    'Procurement Agent': 'Administration',
    'PROJECT ENGINEER (Planing)': 'Engineer',
    'QUALITY CONTROL COORDINATOR': 'Administration',
    'Gang Leader': 'Foreman',
    'Quarry Worker': 'Skilled Labor',
    'Water Treatment  Operator': 'Technical Operator',
    'INDUSTRIAL TECHNICAL': 'Technician',
    'Site/Security': 'Security Guard',
    'Mechanical': 'Semi Skilled Labor',
    'Bobcat Operator': 'Operator',
    'Store Foreman': 'Foreman',
    'Logistic Foreman': 'Foreman',
    'Polishing Foreman': 'Foreman',
    'Project Engineer': 'Engineer',
    'ESTIMATION BIM ENGINEER': 'Engineer',
    'Finish': 'Finisher',
    'Supp': 'Superintendent',
    'Technical Manager': 'Management',
    'Carpenter': 'Skilled Labor',
    'Quantity Surveyor': 'Surveyor',
    'S Labor': 'Skilled Labor',
}

JOB_FAMILY_MAPPING = {
    'Factory - Executive Management': 'Executive Management',
    'Factory - Operator': 'Factory Operator',
    'Showroom - Officer': 'Safety Officer',
    'Installation - Operator': 'Installation Operator',
    'Factory - Foreman': 'Production Foreman',
    'Factory - Superintendent': 'Administration',
    'Installation - Officer': 'Safety Officer',
    'Factory - Labor': 'Labor',
    'Factory - Inspector': 'Factory Inspector',
    'Factory - Skilled Labor': 'Factory Skilled Labor',
    'Factory - Coordinator': 'Coordinator',
    'Showroom - Controller': 'Controller',
    'Factory - Supervisor': 'Factory Supervisor',
    'Factory - Purchaser': 'Administration',
    'Factory - Clerk': 'Clerk',
    'Factory - Controller': 'Controller',
    'Factory - Technician': 'Factory Technician',
    'Installation - Driver': 'Driver',
    'Factory - Officer': 'Safety Officer',
    'Idle Saudi Labor - Operator': 'Installation Operator',
    'Showroom - Operator': 'Factory Operator',
    'Installation - Labor': 'Labor',
    'Showroom - Labor': 'Labor',
    'Factory - Manager': 'Management',
    'Factory - Assistant': 'Assistant',
    'Factory - Engineer': 'Engineer',
    'Factory - Analyst': 'Analyst',
    'Installation - Supervisor': 'Installation Supervisor',
    'Showroom - Store Keeper': 'Store Keeper',
    'Idle Saudi Labor - Messenger': 'Labor',
    'Idle Saudi Labor - Technician': 'Installation Technician',
    'Idle Saudi Labor - Officer': 'Safety Officer',
    'Showroom - Clerk': 'Clerk',
    'Idle Saudi Labor - Skilled Labor': 'Idle Saudi Skilled Labor',
    'Factory - Surveyor': 'Surveyor',
    'Idle Saudi Labor - Representative': 'Representative',
    'Factory - Draftsman': 'Draftsman',
    'Showroom - Coordinator': 'Coordinator',
    'Installation - Controller': 'Controller',
    'Idle Saudi Labor - Purchaser': 'Administration',
    'Showroom - Foreman': 'Production Foreman',
    'Idle Saudi Labor - Clerk': 'Clerk',
    'Factory - Driver': 'Driver',
    'Factory - Time Keeper': 'Time Keeper',
    'Installation - Engineer': 'Engineer',
    'Showroom - Manager': 'Management',
    'Head Office - Manager': 'Management',
    'Head Office - Clerk': 'Clerk',
    'Head Office - Treasurer': 'Administration',
    'Head Office - Skilled Labor': 'Installation Skilled Labor',
    'Head Office - Labor': 'Labor',
    'Head Office - Foreman': 'Installation Foreman',
    'Head Office - Coordinator': 'Coordinator',
    'Head Office - Messenger': 'Labor',
    'Head Office - Purchaser': 'Administration',
    'Head Office - Analyst': 'Analyst',
    'Head Office - Controller': 'Controller',
    'Head Office - Engineer': 'Engineer',
    'Head Office - Supervisor': 'Installation Supervisor',
    'Head Office - Surveyor': 'Surveyor',
    'Head Office - Executive Management': 'Executive Management',
    'Head Office - Officer': 'Safety Officer',
    'Head Office - Assistant': 'Assistant',
    'Installation - Manager': 'Management',
    'Head Office - Driver': 'Driver',
    'Head Office - Draftsman': 'Draftsman',
    'Quarries - Supervisor': 'Quarries Supervisor',
    'Quarries - Technician': 'Quarries Technician',
    'Quarries - Foreman': 'Quarries Foreman',
    'Quarries - Security Guard': 'Security Guard',
    'Quarries - Clerk': 'Clerk',
    'Quarries - Messenger': 'Labor',
    'Quarries - Store Keeper': 'Store Keeper',
    'Quarries - Operator': 'Quarries Operator',
    'Quarries - Driver': 'Driver',
    'Quarries - Geologist': 'Geologist',
    'Quarries - Officer': 'Safety Officer',
    'Quarries - Manager': 'Management',
    'Installation - Foreman': 'Installation Foreman',
    'Installation - Executive Management': 'Executive Management',
    'Installation - Purchaser': 'Administration',
    'Installation - Technician': 'Installation Technician',
    'Installation - Clerk': 'Clerk',
    'Installation - Skilled Labor': 'Installation Skilled Labor',
    'Installation - Messenger': 'Labor',
    'Installation - Store Keeper': 'Store Keeper',
    'Installation - Coordinator': 'Coordinator',
    'Installation - Surveyor': 'Surveyor',
    'Installation - Analyst': 'Analyst',
    'Installation - Draftsman': 'Draftsman',
    'Showroom - Driver': 'Driver',
    'Showroom - Technician': 'Factory Technician',
    'Installation - Time Keeper': 'Time Keeper',
    'Installation - Assistant': 'Assistant',
    'Installation - Inspector': 'Installation Inspector',
    'Installation - Superintendent': 'Administration',
    'Head Office - Inspector': 'Installation Inspector',
    'Showroom - Supervisor': 'Showroom Supervisor',
    'Showroom - Inspector': 'Factory Inspector',
    'Showroom - Engineer': 'Engineer',
    'Showroom - Draftsman': 'Draftsman',
    'Showroom - Surveyor': 'Surveyor',
    'Showroom - Representative': 'Representative',
    'Showroom - Skilled Labor': 'Factory Skilled Labor',
    'Installation - Semi Skilled Labor': 'Semi Skilled Labor',
    'Factory - Semi Skilled Labor': 'Semi Skilled Labor',
    'Factory - Finisher': 'Factory Finisher',
    'Factory - Store Keeper': 'Store Keeper',
    'Factory - Technical Operator': 'Factory Technical Operator',
    'Quarries - Skilled Labor': 'Quarries Skilled Labor',
    'Head Office - Administration': 'Administration',
    'Installation - Administration': 'Administration',
    'Showroom - Administration': 'Administration',
    'Quarries - Technical Operator': 'Quarries Technical Operator',
    'Installation - Security Guard': 'Security Guard',
    'Installation - Safety Officer': 'Safety Officer',
    'Installation - Technical Operator': 'Installation Technical Operator',
    'Installation - Finisher': 'Installation Finisher',
    'Quarries - Coordinator': 'Coordinator',
}

# Initialize session state for workflow stages
if 'stage' not in st.session_state:
    st.session_state.stage = 'upload_raw'

# ===== HELPER FUNCTION FOR SMART JOB FAMILY MAPPING =====
def get_job_family_with_fallback(activity_profession, exact_mapping):
    """
    Try to map Activity-Profession combination to a Job Family.
    If exact match not found, try to match by profession only.
    """
    # First, try exact match
    if activity_profession in exact_mapping:
        return exact_mapping[activity_profession]
    
    # Extract profession part (after the " - ")
    if ' - ' in activity_profession:
        parts = activity_profession.split(' - ')
        activity = parts[0]
        profession = parts[1]
        
        # Try to find a matching job family with same profession
        for key, value in exact_mapping.items():
            # Match by profession: if the job family name ends with the same profession
            if profession.lower() in key.lower() and value.lower() == profession.lower():
                return value
            # Or if the mapped value is just the profession
            if value.lower() == profession.lower():
                return value
        
        # Last resort: try to match with different activity but same profession
        for key, value in exact_mapping.items():
            if profession.lower() in value.lower():
                return value
    
    # If all else fails, return the profession itself
    return activity_profession.split(' - ')[-1] if ' - ' in activity_profession else 'Other'

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
st.markdown("## 🎯 Manpower Optimization Tool")
st.markdown("<p style='color:#888;margin-top:-12px;margin-bottom:8px;'>Upload your staff data and optimize workforce allocation while minimizing total costs</p>", unsafe_allow_html=True)
st.markdown("---")

# ===== DETERMINE WORKFLOW STAGE =====
stage = st.session_state.get('stage', 'upload_raw')

# ===== STAGE 1: UPLOAD & PROCESS RAW DATA =====
if stage == 'upload_raw':
    st.markdown("### 📋 Stage 1: Data Processing")
    st.markdown("Upload your Manpower.xlsx file (with Inhouse, Subcontractor, and Profession Mapping sheets)")
    
    uploaded_file = st.file_uploader("Choose Manpower.xlsx file", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Load all sheets - only Inhouse and Subcontractor sheets are needed
            inhouse_df = pd.read_excel(uploaded_file, sheet_name='Inhouse')
            subcontractor_df = pd.read_excel(uploaded_file, sheet_name='Subcontractor')
            
            # Strip whitespace from column names to handle trailing spaces
            inhouse_df.columns = inhouse_df.columns.str.strip()
            subcontractor_df.columns = subcontractor_df.columns.str.strip()
            
            st.success("✅ File loaded successfully!")
            st.info("📝 Using embedded 3-step hierarchical mapping: Activity → Profession → Job Family")
            
            # ===== PROCESS INHOUSE DATA =====
            st.markdown("---")
            st.markdown("#### Processing In-House Staff...")
            
            # Step 1: Standardize Activity (Location → Standardized Activity)
            inhouse_df['Activity_Standardized'] = inhouse_df['Location'].map(ACTIVITY_MAPPING).fillna(inhouse_df['Location'])
            
            # Step 2: Standardize Profession
            inhouse_df['Profession_Standardized'] = inhouse_df['Profession'].map(PROFESSION_MAPPING).fillna(inhouse_df['Profession'])
            
            # Step 3: Merge Activity + Profession and map to Job Family
            inhouse_df['Activity_Profession'] = inhouse_df['Activity_Standardized'] + ' - ' + inhouse_df['Profession_Standardized']
            inhouse_df['Job_Family'] = inhouse_df['Activity_Profession'].apply(lambda x: get_job_family_with_fallback(x, JOB_FAMILY_MAPPING))
            
            inhouse_df['Is_Saudi'] = (inhouse_df['Nationality'] == 'SAUDI').astype(int)
            inhouse_df['Cost_Per_Employee'] = inhouse_df['Total Paid'] + inhouse_df['Total Unpaid']
            
            # Add overtime cost (O.T Hrs) - assume 50 SAR/hr
            if 'O.T Hrs' in inhouse_df.columns:
                inhouse_df['Cost_Per_Employee'] += inhouse_df['O.T Hrs'].fillna(0) * 50
            
            inhouse_summary = inhouse_df.groupby('Job_Family').agg({
                'No': 'count',
                'Is_Saudi': 'sum',
                'Cost_Per_Employee': 'sum'
            }).rename(columns={'No': 'Total_Inhouse', 'Is_Saudi': 'Saudi_Inhouse'})
            
            inhouse_summary['Cost_NonSaudi_Inhouse'] = inhouse_summary['Cost_Per_Employee']
            inhouse_summary['Avg_Cost_Per_Employee'] = inhouse_summary['Cost_Per_Employee'] / inhouse_summary['Total_Inhouse']
            
            st.write(f"✅ Processed {len(inhouse_df)} in-house employees across {len(inhouse_summary)} job families")
            
            # ===== PROCESS SUBCONTRACTOR DATA =====
            st.markdown("#### Processing Subcontracted Staff...")
            
            # Step 1: Standardize Activity (Working in → Standardized Activity)
            subcontractor_df['Activity_Standardized'] = subcontractor_df['Working in'].map(ACTIVITY_MAPPING).fillna(subcontractor_df['Working in'])
            
            # Step 2: Standardize Profession
            subcontractor_df['Profession_Standardized'] = subcontractor_df['Profession'].map(PROFESSION_MAPPING).fillna(subcontractor_df['Profession'])
            
            # Step 3: Merge Activity + Profession and map to Job Family
            subcontractor_df['Activity_Profession'] = subcontractor_df['Activity_Standardized'] + ' - ' + subcontractor_df['Profession_Standardized']
            subcontractor_df['Job_Family'] = subcontractor_df['Activity_Profession'].apply(lambda x: get_job_family_with_fallback(x, JOB_FAMILY_MAPPING))
            
            subcontractor_df['Is_Saudi'] = (subcontractor_df['Nationality'] == 'SAUDI').astype(int)
            
            # Calculate outsourced cost with sponsor-dependent insurance
            def calc_outsource_cost(row):
                cost = 0
                cost += row['Basic'] if pd.notna(row['Basic']) else 0
                cost += row['Housing Paid'] if pd.notna(row['Housing Paid']) else 0
                cost += row['Trans Paid'] if pd.notna(row['Trans Paid']) else 0
                cost += row['Food'] if pd.notna(row['Food']) else 0
                cost += row['Gosi'] if pd.notna(row['Gosi']) else 0
                
                # Insurance based on sponsor
                sponsor = row['Sponser'] if pd.notna(row['Sponser']) else 'Unknown'
                if sponsor in ['Ewan', 'Mahara', 'Tatweer']:
                    cost += 38
                elif sponsor in ['Saed Azka']:
                    cost += 46
                elif sponsor == 'ARCO':
                    cost += 50
                else:
                    cost += 38  # Default
                
                cost += row['Value O.T (SAR)'] if pd.notna(row['Value O.T (SAR)']) else 0
                cost += row['Government fees'] if pd.notna(row['Government fees']) else 0
                cost += row['Service margin'] if pd.notna(row['Service margin']) else 0
                cost += row['E.O.S monthly'] if pd.notna(row['E.O.S monthly']) else 0
                
                return cost
            
            subcontractor_df['Cost_Per_Employee'] = subcontractor_df.apply(calc_outsource_cost, axis=1)
            
            subcontractor_summary = subcontractor_df.groupby('Job_Family').agg({
                'No': 'count',
                'Is_Saudi': 'sum',
                'Cost_Per_Employee': 'sum'
            }).rename(columns={'No': 'Total_Outsourced', 'Is_Saudi': 'Saudi_Outsourced'})
            
            subcontractor_summary['Cost_Outsourced'] = subcontractor_summary['Cost_Per_Employee']
            subcontractor_summary['Avg_Cost_Per_Employee'] = subcontractor_summary['Cost_Per_Employee'] / subcontractor_summary['Total_Outsourced']
            
            st.write(f"✅ Processed {len(subcontractor_df)} subcontracted employees across {len(subcontractor_summary)} job families")
            
            # ===== MERGE & CREATE OPTIMIZATION INPUT =====
            st.markdown("#### Generating Optimization Input...")
            
            # Get all job families from both inhouse and subcontractor data
            all_job_families = set(inhouse_summary.index) | set(subcontractor_summary.index)
            optimization_data = []
            
            for job_family in sorted(all_job_families):
                inhouse_row = inhouse_summary.loc[job_family] if job_family in inhouse_summary.index else None
                outsource_row = subcontractor_summary.loc[job_family] if job_family in subcontractor_summary.index else None
                
                # Calculate totals
                total_inhouse = inhouse_row['Total_Inhouse'] if inhouse_row is not None else 0
                total_outsourced = outsource_row['Total_Outsourced'] if outsource_row is not None else 0
                total_employees = total_inhouse + total_outsourced
                
                total_inhouse_saudi = int(inhouse_row['Saudi_Inhouse']) if inhouse_row is not None else 0
                total_inhouse_non_saudi = int(total_inhouse - total_inhouse_saudi)
                
                # Calculate average costs per employee
                avg_cost_inhouse_saudi = (inhouse_row['Cost_Per_Employee'] / total_inhouse) if (inhouse_row is not None and total_inhouse > 0) else 0
                avg_cost_inhouse_non_saudi = (inhouse_row['Cost_Per_Employee'] / total_inhouse_non_saudi) if (inhouse_row is not None and total_inhouse_non_saudi > 0) else 0
                avg_cost_outsourced = (outsource_row['Cost_Outsourced'] / total_outsourced) if (outsource_row is not None and total_outsourced > 0) else 0
                
                max_outsource_ratio = 0.5  # Default 50%
                
                if total_employees > 0:
                    optimization_data.append({
                        'Job Family': job_family,
                        'Avg Cost Non-Saudi Inhouse': avg_cost_inhouse_non_saudi,
                        'Avg Cost Saudi Inhouse': avg_cost_inhouse_saudi,
                        'Avg Cost Outsourced': avg_cost_outsourced,
                        'Total Employees': int(total_employees),
                        'Total Inhouse Saudi': total_inhouse_saudi,
                        'Max Outsource Ratio': max_outsource_ratio
                    })
            
            optimization_df = pd.DataFrame(optimization_data)
            
            st.write(f"✅ Created optimization input with {len(optimization_df)} job families")
            
            # Format numbers for display (thousands separator, no decimals)
            display_df = optimization_df.copy()
            display_df['Total Employees'] = display_df['Total Employees'].apply(lambda x: f"{int(x):,}")
            display_df['Total Inhouse Saudi'] = display_df['Total Inhouse Saudi'].apply(lambda x: f"{int(x):,}")
            display_df['Avg Cost Non-Saudi Inhouse'] = display_df['Avg Cost Non-Saudi Inhouse'].apply(lambda x: f"{x:,.0f}")
            display_df['Avg Cost Saudi Inhouse'] = display_df['Avg Cost Saudi Inhouse'].apply(lambda x: f"{x:,.0f}")
            display_df['Avg Cost Outsourced'] = display_df['Avg Cost Outsourced'].apply(lambda x: f"{x:,.0f}")
            display_df['Max Outsource Ratio'] = display_df['Max Outsource Ratio'].apply(lambda x: f"{x:.1%}")
            
            st.dataframe(display_df, use_container_width=True)
            
            # Store in session
            st.session_state.optimization_df = optimization_df
            st.session_state.inhouse_cleaned = inhouse_df
            st.session_state.subcontractor_cleaned = subcontractor_df
            
            # ===== DOWNLOAD CLEANED DATA =====
            st.markdown("---")
            st.markdown("### 📥 Download Cleaned Data")
            st.markdown("Download the processed and cleaned data before optimization")
            
            # Create Excel file with cleaned data
            cleaned_buffer = io.BytesIO()
            with pd.ExcelWriter(cleaned_buffer, engine='openpyxl') as writer:
                # Select relevant columns for download - Inhouse
                inhouse_cols = ['No', 'Nationality', 'Location', 'Activity_Standardized', 'Profession', 
                               'Profession_Standardized', 'Job_Family', 'Is_Saudi', 
                               'Total Paid', 'Total Unpaid', 'Cost_Per_Employee']
                inhouse_export = inhouse_df[[col for col in inhouse_cols if col in inhouse_df.columns]].copy()
                inhouse_export.to_excel(writer, sheet_name='Inhouse Cleaned', index=False)
                
                # Select relevant columns for download - Subcontractor
                subcontractor_cols = ['No', 'Nationality', 'Working in', 'Activity_Standardized', 'Profession', 
                                     'Profession_Standardized', 'Job_Family', 'Is_Saudi',
                                     'Basic', 'Housing Paid', 'Trans Paid', 'Food', 'Gosi', 'Sponser',
                                     'Cost_Per_Employee']
                subcontractor_export = subcontractor_df[[col for col in subcontractor_cols if col in subcontractor_df.columns]].copy()
                subcontractor_export.to_excel(writer, sheet_name='Subcontractor Cleaned', index=False)
                
                # Optimization input
                optimization_df.to_excel(writer, sheet_name='Optimization Input', index=False)
            
            cleaned_buffer.seek(0)
            st.download_button(
                label="📊 Download Cleaned Data as Excel",
                data=cleaned_buffer.getvalue(),
                file_name="Manpower_Cleaned_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_cleaned_data"
            )
            
            st.session_state.stage = 'optimize'
            
            st.markdown("---")
            st.markdown("### ✅ Ready for Optimization!")
            st.info("Click the button below to proceed to the optimization stage")
            
            if st.button("Proceed to Optimization →", type='primary'):
                st.rerun()
        
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            import traceback
            st.write(traceback.format_exc())

# ===== STAGE 2: OPTIMIZATION =====
elif stage == 'optimize':
    
    # Back button at top
    if st.button("← Back to Data Processing"):
        st.session_state.stage = 'upload_raw'
        st.session_state.pop('optimization_df', None)
        st.rerun()
    
    st.markdown("---")
    st.markdown("### 📊 Stage 2: Optimization Settings")
    
    # ===== SIDEBAR SETTINGS =====
    st.sidebar.markdown("## ⚙️ Optimization Settings")
    st.sidebar.markdown("---")
    
    st.sidebar.markdown("### 1️⃣ Saudization Rate")
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
    st.sidebar.markdown("### 2️⃣ Saudi Labor Policy")
    can_fire_saudi = st.sidebar.checkbox("Allow reducing current Saudi headcount?", value=False)
    if can_fire_saudi:
        st.sidebar.warning("Saudi labor can be reduced below current levels")
    else:
        st.sidebar.info("Saudi labor minimum constraint applied")
    
    st.sidebar.markdown("---")
    run = st.sidebar.button("🚀 Run Optimization", type="primary", use_container_width=True)
    
    # ===== PREPARE DATA =====
    optimization_df = st.session_state.optimization_df
    
    # Convert to the format expected by optimization
    data = optimization_df.copy()
    
    COL_JOB_FAMILY = 0
    COL_TOTAL_EMPLOYEES = 4
    COL_TOTAL_INHOUSE_SAUDI = 5
    COL_AVG_COST_SAUDI = 2
    COL_AVG_COST_NON_SAUDI = 1
    COL_AVG_COST_OUTSOURCED = 3
    COL_MAX_OUTSOURCE_RATIO = 6
    
    outsource_type = "Current"
    
    # ===== RUN OPTIMIZATION =====
    if run:
        with st.spinner("Optimizing workforce allocation..."):
            try:
                prob = pulp.LpProblem("Manpower_Optimization", pulp.LpMinimize)
                S, N, O = [], [], []
    
                for i in range(len(data)):
                    current_saudi = data.iloc[i]['Total Inhouse Saudi']
                    total_employees = data.iloc[i]['Total Employees']
                    max_outsource_ratio = data.iloc[i]['Max Outsource Ratio']
                    saudi_lower_bound = 0 if can_fire_saudi else current_saudi
    
                    s = pulp.LpVariable(f'S_{i}', lowBound=saudi_lower_bound, cat='Integer')
                    n = pulp.LpVariable(f'N_{i}', lowBound=0, cat='Integer')
                    o = pulp.LpVariable(f'O_{i}', lowBound=0, cat='Integer')
                    S.append(s); N.append(n); O.append(o)
    
                    prob += s + n + o == total_employees, f"Total_Employees_{i}"
                    prob += o <= max_outsource_ratio * total_employees, f"Max_Outsource_Ratio_{i}"
    
                # Objective: minimize total cost using average costs per employee
                prob += pulp.lpSum(
                    data.iloc[i]['Avg Cost Saudi Inhouse'] * S[i] +
                    data.iloc[i]['Avg Cost Non-Saudi Inhouse'] * N[i] +
                    data.iloc[i]['Avg Cost Outsourced'] * O[i]
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
                        
                        # Calculate costs using average costs
                        cost_saudi = data.iloc[i]['Avg Cost Saudi Inhouse'] * saudi
                        cost_non_saudi = data.iloc[i]['Avg Cost Non-Saudi Inhouse'] * non_saudi
                        cost_outsourced = data.iloc[i]['Avg Cost Outsourced'] * outsourced
                        
                        results_data.append({
                            'Job Family': data.iloc[i]['Job Family'],
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
                    st.session_state.outsource_type = outsource_type
                    st.success("✅ Optimization completed successfully!")
                else:
                    st.error(f"❌ Optimization failed: {pulp.LpStatus[prob.status]}")
    
            except Exception as e:
                st.error(f"Error during optimization: {str(e)}")
                import traceback
                st.write(traceback.format_exc())
    
    # ===== DISPLAY RESULTS =====
    if hasattr(st.session_state, 'optimization_status'):
        st.markdown("---")
        st.markdown("### 📊 Optimization Results")
        
        # ----- KPI METRICS -----
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("💰 Total Cost", f"SAR {st.session_state.total_cost:,.0f}")
        with col2:
            st.metric("👥 Total Employees", f"{st.session_state.total_employees_final:,}")
        with col3:
            st.metric("📊 Saudization Rate", f"{st.session_state.saudization_achieved:.1f}%")
        
        col4, col5, col6 = st.columns(3)
        with col4:
            st.metric("🇸🇦 In-House Saudi", f"{st.session_state.total_saudi_final:,}")
        with col5:
            st.metric("🌍 In-House Non-Saudi", f"{st.session_state.total_non_saudi_final:,}")
        with col6:
            st.metric("🤝 Outsourced", f"{st.session_state.total_outsourced_final:,}")
    
        st.markdown("---")
    
        # ----- COST ANALYSIS -----
        st.markdown("### 📈 Cost Analysis")
    
        viz_col1, viz_col2 = st.columns(2)
    
        COLORS_METHOD = ['#1D9E75', '#378ADD', '#F0993B']
        COLORS_FAMILY = [
            '#1D9E75', '#378ADD', '#F0993B', '#7F77DD', '#D85A30',
            '#D4537E', '#639922', '#BA7517', '#888780', '#185FA5', '#5F5E5A'
        ]
    
        # Pie 1: Cost by sourcing method
        with viz_col1:
            total_all = (st.session_state.total_cost_saudi +
                         st.session_state.total_cost_non_saudi +
                         st.session_state.total_cost_outsourced)
    
            labels_m = ['In-House Saudi', 'In-House Non-Saudi', 'Outsourced']
            values_m = [
                st.session_state.total_cost_saudi,
                st.session_state.total_cost_non_saudi,
                st.session_state.total_cost_outsourced
            ]
    
            fig_method = go.Figure(data=[go.Pie(
                labels=labels_m,
                values=values_m,
                marker=dict(colors=COLORS_METHOD, line=dict(color='#ffffff', width=2)),
                hovertemplate='<b>%{label}</b><br>SAR %{value:,.0f}<br>%{percent}<extra></extra>',
                textinfo='percent',
                textposition='auto',
                hole=0.45
            )])
            fig_method.update_layout(
                title=dict(text='Cost by sourcing method', font=dict(size=14, color='#1a1a1a'), x=0, xanchor='left'),
                height=360,
                margin=dict(t=40, b=20, l=20, r=20),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                legend=dict(
                    orientation='v',
                    x=1.02, y=0.5,
                    xanchor='left', yanchor='middle',
                    font=dict(size=12, color='#444'),
                    bgcolor='rgba(0,0,0,0)'
                ),
                annotations=[dict(
                    text=f"SAR<br>{total_all/1e6:.1f}M" if total_all >= 1e6 else f"SAR<br>{total_all:,.0f}",
                    x=0.5, y=0.5, font_size=13, showarrow=False,
                    font=dict(color='#1a1a1a', family='sans-serif')
                )]
            )
            st.plotly_chart(fig_method, use_container_width=True)
    
        # Pie 2: Cost by Job Family — top 10 + Other
        with viz_col2:
            results_sorted = st.session_state.results_df.copy()
            results_sorted = results_sorted.sort_values('Total Cost (SAR)', ascending=False).reset_index(drop=True)
    
            top10 = results_sorted.head(10)
            other_rows = results_sorted.iloc[10:]
            other_cost = other_rows['Total Cost (SAR)'].sum()
            other_count = len(other_rows)
    
            job_families = list(top10['Job Family'])
            costs = list(top10['Total Cost (SAR)'])
    
            if other_cost > 0:
                job_families.append(f"Other ({other_count} families)")
                costs.append(other_cost)
    
            fig_family = go.Figure(data=[go.Pie(
                labels=job_families,
                values=costs,
                marker=dict(colors=COLORS_FAMILY[:len(job_families)], line=dict(color='#ffffff', width=2)),
                hovertemplate='<b>%{label}</b><br>SAR %{value:,.0f}<br>%{percent}<extra></extra>',
                textinfo='percent', 
                textposition='auto',
                hole=0.45
            )])
            fig_family.update_layout(
                title=dict(text='Cost by job family (top 10)', font=dict(size=14, color='#1a1a1a'), x=0, xanchor='left'),
                height=360,
                margin=dict(t=40, b=20, l=20, r=20),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                legend=dict(
                    orientation='v',
                    x=1.02, y=0.5,
                    xanchor='left', yanchor='middle',
                    font=dict(size=11, color='#444'),
                    bgcolor='rgba(0,0,0,0)'
                )
            )
            st.plotly_chart(fig_family, use_container_width=True)
    
        st.markdown("---")
    
        # ----- DETAILED TABLE -----
        st.markdown("### 📋 Detailed Allocation by Job Family")
        st.markdown("<p style='color:#888;font-size:13px;margin-top:-10px;margin-bottom:12px;'>Click each row to expand the cost breakdown</p>", unsafe_allow_html=True)
    
        for idx, row in st.session_state.results_df.iterrows():
            with st.expander(f"**{row['Job Family']}** — SAR {row['Total Cost (SAR)']:,.0f}"):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Saudi", f"{int(row['In-House Saudi']):,}")
                c2.metric("Non-Saudi", f"{int(row['In-House Non-Saudi']):,}")
                c3.metric("Outsourced", f"{int(row['Outsourced']):,}")
                c4.metric("Total", f"{int(row['Total Employees']):,}")
    
                fig_bk = go.Figure(data=[go.Bar(
                    x=['In-House Saudi', 'In-House Non-Saudi', 'Outsourced'],
                    y=[float(row['Cost - Saudi (SAR)']),
                       float(row['Cost - Non-Saudi (SAR)']),
                       float(row['Cost - Outsourced (SAR)'])],
                    marker_color=['#1D9E75', '#378ADD', '#F0993B'],
                    hovertemplate='%{x}<br>SAR %{y:,.0f}<extra></extra>'
                )])
                fig_bk.update_layout(
                    height=260,
                    margin=dict(t=10, b=10, l=10, r=10),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    yaxis=dict(title='Cost (SAR)', gridcolor='#eeeeea', tickfont=dict(size=11)),
                    xaxis=dict(gridcolor='rgba(0,0,0,0)'),
                    showlegend=False
                )
                st.plotly_chart(fig_bk, use_container_width=True, key=f"breakdown_chart_{idx}")
    
        st.markdown("---")
    
        # ----- DOWNLOAD -----
        st.markdown("### 📥 Download Results")
    
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            st.session_state.results_df.to_excel(writer, sheet_name='Optimization Results', index=False)
    
            summary_data = {
                'Metric': ['Total Cost (SAR)', 'Total Employees', 'In-House Saudi',
                           'In-House Non-Saudi', 'Outsourced', 'Saudization Rate Achieved (%)',
                           'Optimization Status', 'Outsourced Cost Type',
                           'Can Reduce Saudi', 'Saudization Enforced'],
                'Value': [f'{st.session_state.total_cost:,.0f}',
                          st.session_state.total_employees_final,
                          st.session_state.total_saudi_final,
                          st.session_state.total_non_saudi_final,
                          st.session_state.total_outsourced_final,
                          f'{st.session_state.saudization_achieved:.2f}',
                          st.session_state.optimization_status,
                          st.session_state.outsource_type,
                          'Yes' if can_fire_saudi else 'No',
                          'Yes' if enforce_saudization else 'No']
            }
            if enforce_saudization:
                summary_data['Metric'].insert(6, 'Saudization Rate Required (%)')
                summary_data['Value'].insert(6, f'{SAUDIZATION_RATE*100:.2f}')
    
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
    
        output_buffer.seek(0)
        st.download_button(
            label="📊 Download Results as Excel",
            data=output_buffer.getvalue(),
            file_name="Manpower_Optimization_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )