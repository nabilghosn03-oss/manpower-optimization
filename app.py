import pandas as pd
import pulp

# Constants
EXCEL_FILE = r"C:\Users\USER\Desktop\Manpower_input.xlsx"

# Read data
try:
    df = pd.read_excel(EXCEL_FILE, header=None)
    print("Excel file loaded successfully.")
    print("First few rows:")
    print(df.head())
    print("\nShape:", df.shape)
except FileNotFoundError:
    print(f"Error: Excel file not found at {EXCEL_FILE}")
    exit(1)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit(1)

# Column mapping (based on Excel column order)
# The headers are in row 1
headers = df.iloc[1].tolist()
print("\nColumn Headers:")
for i, header in enumerate(headers):
    print(f"Column {i}: {header}")

# Extract data starting from row 2 (row 1 is headers, row 0 is empty)
data = df.iloc[2:].reset_index(drop=True)

# Define column indices
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

# ===== USER INPUTS =====
print("\n" + "="*80)
print("OPTIMIZATION PARAMETERS")
print("="*80)

# Question 1: Outsourcing costs
print("\n1. Which outsourced cost would you like to use?")
print("   - Enter 'current' for Current Outsourced Cost")
print("   - Enter 'improved' for Improved Outsourced Cost")
outsource_choice = input("Your choice (current/improved): ").strip().lower()

if outsource_choice == 'improved':
    COL_COST_OUTSOURCED = COL_OUTSOURCE_IMPROVED_COST
    outsource_type = "Improved"
else:
    COL_COST_OUTSOURCED = COL_OUTSOURCE_CURRENT_COST
    outsource_type = "Current"

print(f"✓ Using {outsource_type} Outsourced Costs")

# Question 2: Saudization rate
print("\n2. Do you want to enforce an overall Saudization Rate?")
saud_choice = input("Enter 'yes' or 'no': ").strip().lower()

if saud_choice == 'yes':
    try:
        SAUDIZATION_RATE = float(input("Enter the Saudization Rate (e.g., 0.3 for 30%): ").strip())
        print(f"✓ Saudization Rate constraint: {SAUDIZATION_RATE*100:.2f}%")
    except ValueError:
        print("Invalid input. Using default rate of 0.3 (30%)")
        SAUDIZATION_RATE = 0.3
    enforce_saudization = True
else:
    SAUDIZATION_RATE = None
    enforce_saudization = False
    print("✓ No overall Saudization Rate constraint")

# Question 3: Can fire Saudi labor
print("\n3. Can you fire current Saudi labor?")
fire_choice = input("Enter 'yes' or 'no': ").strip().lower()

can_fire_saudi = fire_choice == 'yes'
if can_fire_saudi:
    print("✓ Saudi labor can be reduced below current levels")
else:
    print("✓ Current Saudi labor cannot be fired (minimum constraint applied)")

# ===== OPTIMIZATION =====
print("\n" + "="*80)
print("RUNNING OPTIMIZATION...")
print("="*80 + "\n")

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
    # Non-Saudi in-house: can be 0 or more
    n = pulp.LpVariable(f'N_{i}', lowBound=0, cat='Integer')
    # Outsourced: can be 0 or more
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
prob.solve()

# Print and save results
print("\n" + "="*80)
print("OPTIMIZATION RESULTS")
print("="*80)
print(f"Status: {pulp.LpStatus[prob.status]}")
if pulp.LpStatus[prob.status] == 'Optimal':
    total_cost = pulp.value(prob.objective)
    print(f"Total Cost: ${total_cost:,.2f}")
    
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
    total_saudi_final = results_df['In-House Saudi'].sum()
    total_non_saudi_final = results_df['In-House Non-Saudi'].sum()
    total_outsourced_final = results_df['Outsourced'].sum()
    total_employees_final = results_df['Total Employees'].sum()
    total_inhouse_final = total_saudi_final + total_non_saudi_final
    saudization_achieved = (total_saudi_final / total_inhouse_final * 100) if total_inhouse_final > 0 else 0
    total_cost_final = results_df['Total Cost'].sum()
    
    # Print to console
    print(f"\nOptimal Allocation per Job Family:")
    print("-"*80)
    print(f"{'Job Family':<20} {'Saudi':<10} {'Non-Saudi':<10} {'Outsourced':<12} {'Total':<8}")
    print("-"*80)
    for idx, row in results_df.iterrows():
        print(f"{row['Job Family']:<20} {int(row['In-House Saudi']):<10} {int(row['In-House Non-Saudi']):<10} {int(row['Outsourced']):<12} {int(row['Total Employees']):<8}")
    
    print("-"*80)
    print(f"{'TOTAL':<20} {total_saudi_final:<10} {total_non_saudi_final:<10} {total_outsourced_final:<12} {total_employees_final:<8}")
    print(f"\nSaudization Rate Achieved: {saudization_achieved:.2f}%")
    if enforce_saudization:
        print(f"Saudization Rate Required: {SAUDIZATION_RATE*100:.2f}%")
    print(f"Total Cost: ${total_cost_final:,.2f}")
    
    # Save to Excel
    output_file = r"C:\Users\USER\Desktop\Manpower_Optimization_Results.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write results summary
        results_df.to_excel(writer, sheet_name='Optimization Results', index=False)
        
        # Write summary sheet
        summary_values = {
            'Metric': ['Total Cost', 'Total Employees', 'In-House Saudi', 'In-House Non-Saudi', 'Outsourced', 
                      'Saudization Rate Achieved (%)', 'Optimization Status', 'Outsourced Cost Type', 'Can Fire Saudi', 'Saudization Enforced'],
            'Value': [f'{total_cost_final:,.2f}', total_employees_final, total_saudi_final, total_non_saudi_final, 
                     total_outsourced_final, f'{saudization_achieved:.2f}', pulp.LpStatus[prob.status],
                     outsource_type, 'Yes' if can_fire_saudi else 'No', 'Yes' if enforce_saudization else 'No']
        }
        
        if enforce_saudization:
            summary_values['Metric'].insert(6, 'Saudization Rate Required (%)')
            summary_values['Value'].insert(6, f'{SAUDIZATION_RATE*100:.2f}')
        
        summary_df = pd.DataFrame(summary_values)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f"\n✓ Results saved to: {output_file}")
else:
    print("No optimal solution found")