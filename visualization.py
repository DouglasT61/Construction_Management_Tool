import dash
from dash import html, dcc
from dash.dependencies import Input, Output
import plotly.graph_objs as go
import pandas as pd
import openpyxl

# Initialize the Dash app
app = dash.Dash(__name__)

# Read the Excel file
workbook = openpyxl.load_workbook(r'C:\Users\dthom\OneDrive\Personal\Hart Advisors Group\GitHub\Construction_Management_Tool\(2021.08.04) Coliseum Storage Development Tracker June 2021.xlsx', data_only=True)
worksheet = workbook['Variance']
data_sheet = pd.DataFrame([[cell.value for cell in row] for row in worksheet.rows])

print("\nAvailable Column Labels:")
print(data_sheet[1].tolist())  # Print all values in column B

def find_row_index(df, value, column_start=1, column_end=2):
    try:
        print(f"\nSearching for: {value}")
        # Check both columns B and C (indices 1 and 2)
        for col in range(column_start, column_end + 1):
            matches = df[df[col].str.contains(value, na=False, case=False, regex=False)]  # Added regex=False for exact matching
            if not matches.empty:
                print(f"Found matches in column {col}: {matches[col].tolist()}")
                return matches.index[0]
        print(f"No matches found for: {value}")
        return None
    except (AttributeError, IndexError) as e:
        print(f"Error searching for {value}: {str(e)}")
        return None

# Update the data extraction with exact labels from the Excel file
revenue_idx = find_row_index(data_sheet, "Actual Operating Revenues Total", column_start=1, column_end=1)
expenses_idx = find_row_index(data_sheet, "Actual Operating Expenses", column_start=1, column_end=1)
debt_service_idx = find_row_index(data_sheet, "Debt Service", column_start=1, column_end=1)

# Add debug prints to verify the exact row contents
print("\nRow Contents Debug:")
if revenue_idx is not None:
    print(f"Revenue row ({revenue_idx}):")
    print(data_sheet.iloc[revenue_idx, :].tolist())
if expenses_idx is not None:
    print(f"Expenses row ({expenses_idx}):")
    print(data_sheet.iloc[expenses_idx, :].tolist())
if debt_service_idx is not None:
    print(f"Debt Service row ({debt_service_idx}):")
    print(data_sheet.iloc[debt_service_idx, :].tolist())

# For cost tracking
budget_idx = find_row_index(data_sheet, "Total Project Cost")
actual_cost_idx = find_row_index(data_sheet, "Actual GC Costs Disbursement")

# For lease-up tracking
units_leased_idx = find_row_index(data_sheet, "Cumulative Units Lease")
planned_units_idx = find_row_index(data_sheet, "Units Lease in Month")

# Add debug output for both columns
print("\nColumn B (index 1) values:")
print(data_sheet[1].tolist())
print("\nColumn C (index 2) values:")
print(data_sheet[2].tolist())

# Extract data from 'Variance'
revenue = pd.to_numeric(data_sheet.iloc[revenue_idx, 22:], errors='coerce').fillna(0) if revenue_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))
expenses = pd.to_numeric(data_sheet.iloc[expenses_idx, 22:], errors='coerce').fillna(0) if expenses_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))
debt_service = pd.to_numeric(data_sheet.iloc[debt_service_idx, 22:], errors='coerce').fillna(0) if debt_service_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))

# Extract budget and cost data
budget_to_complete = pd.to_numeric(data_sheet.iloc[budget_idx, 22:], errors='coerce').fillna(0) if budget_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))
est_cost_to_complete = pd.to_numeric(data_sheet.iloc[actual_cost_idx, 22:], errors='coerce').fillna(0) if actual_cost_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))

# Extract units leased data
units_leased = pd.to_numeric(data_sheet.iloc[units_leased_idx, 22:], errors='coerce').fillna(0) if units_leased_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))

# Extract planned units leased data
planned_units = pd.to_numeric(data_sheet.iloc[planned_units_idx, 22:], errors='coerce').fillna(0) if planned_units_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))

# Extract S-curve data
s_curve_proj = [0, 5, 15, 30, 45, 60, 70, 78, 84, 89, 93, 96, 98, 99, 99.5, 100, 100, 100]
s_curve_act = [0, 3, 12, 28, 42, 58, 68, 76, 82, 87, 91, 94, 96, 97, 98, 99, 99.5, 100]

# Define the full range of months
months = range(1, 53)  # 52 months total

# After extracting the data, add this debug section:
print("\nValue Debug:")
print(f"Revenue values: {revenue.tolist()}")
print(f"Expenses values: {expenses.tolist()}")
print(f"Debt Service values: {debt_service.tolist()}")
print(f"Budget values: {budget_to_complete.tolist()}")
print(f"Actual Cost values: {est_cost_to_complete.tolist()}")
print(f"Units Leased values: {units_leased.tolist()}")
print(f"Planned Units values: {planned_units.tolist()}")

# After finding revenue_idx, add this debug code
print("\nDetailed Revenue Debug:")
print("Raw revenue row data:")
print(data_sheet.iloc[revenue_idx, :].tolist())
print("\nRevenue values starting at column 22:")
print(data_sheet.iloc[revenue_idx, 22:].tolist())

# Also let's check the actual Excel cell values
print("\nChecking Excel worksheet directly:")
for col in range(22, worksheet.max_column):
    value = worksheet.cell(row=revenue_idx + 1, column=col).value
    if value is not None:
        print(f"Column {col}: {value}")

# After the existing data extraction code, add:
# Extract S-curve cumulative data
dai_scurve_idx = find_row_index(data_sheet, "DAI Estimated S-Curve", column_start=1, column_end=1)
dai_cumulative_idx = find_row_index(data_sheet, "Cumulative", column_start=1, column_end=1)
gc_costs_idx = find_row_index(data_sheet, "Actual GC Costs Disbursement", column_start=1, column_end=1)
gc_cumulative_idx = gc_costs_idx + 1 if gc_costs_idx is not None else None

# Extract cumulative values
dai_cumulative = pd.to_numeric(data_sheet.iloc[dai_cumulative_idx, 22:], errors='coerce').fillna(0) if dai_cumulative_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))
gc_cumulative = pd.to_numeric(data_sheet.iloc[gc_cumulative_idx, 22:], errors='coerce').fillna(0) if gc_cumulative_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))

# Find and extract equity value from Excel
equity_idx = None
print("\nEquity Debug:")
# Find all rows that contain "Equity"
matches = data_sheet[data_sheet[2] == "Equity"].index
if len(matches) > 0:
    # Use the last match (should be the actual Equity row, not CBRE Equity Fee)
    equity_idx = matches[-1]
    print(f"Equity row ({equity_idx}):")
    print(f"Row label: {data_sheet.iloc[equity_idx, 2]}")
    equity_values = pd.to_numeric(data_sheet.iloc[equity_idx, 22:], errors='coerce').fillna(0)
    print(f"Equity values: {equity_values.tolist()}")
    equity = equity_values[equity_values.ne(0)].iloc[-1] if not equity_values.eq(0).all() else 11664124
else:
    print("Equity row not found, using default value")
    equity = 11664124

print(f"Using equity value: ${equity:,.2f}")

# Debug prints for data extraction
print("\nData Extraction Debug:")
print(f"Revenue row index: {revenue_idx}")
print(f"Expenses row index: {expenses_idx}")
print(f"Debt Service row index: {debt_service_idx}")
print(f"Budget row index: {budget_idx}")
print(f"Actual Cost row index: {actual_cost_idx}")
print(f"Units Leased row index: {units_leased_idx}")
print(f"Planned Units row index: {planned_units_idx}")

# After extracting the data, add this debug section:
print("\nValue Debug:")
print(f"Revenue values: {revenue.tolist()}")
print(f"Expenses values: {expenses.tolist()}")
print(f"Debt Service values: {debt_service.tolist()}")
print(f"Budget values: {budget_to_complete.tolist()}")
print(f"Actual Cost values: {est_cost_to_complete.tolist()}")
print(f"Units Leased values: {units_leased.tolist()}")
print(f"Planned Units values: {planned_units.tolist()}")

# After finding revenue_idx, add this debug code
print("\nDetailed Revenue Debug:")
print("Raw revenue row data:")
print(data_sheet.iloc[revenue_idx, :].tolist())
print("\nRevenue values starting at column 22:")
print(data_sheet.iloc[revenue_idx, 22:].tolist())

# Also let's check the actual Excel cell values
print("\nChecking Excel worksheet directly:")
for col in range(22, worksheet.max_column):
    value = worksheet.cell(row=revenue_idx + 1, column=col).value
    if value is not None:
        print(f"Column {col}: {value}")

# After the existing data extraction code, add:
# Constants for schedule tracking, occupancy, and budget
PLANNED_COMPLETION_DAYS = 310  # GC Contract Days
PLANNED_COMPLETION_DATE = pd.to_datetime('2021-03-28')
CURRENT_DATE = pd.to_datetime('2021-06-30')
TOTAL_UNITS = 1098  # Total number of units in the project
INITIAL_BUDGET = 27664124  # Initial project budget

# Extract cash flow data
cash_flow_idx = find_row_index(data_sheet, "Operating Cash Flow After Reserves", column_start=1, column_end=1)
cash_flow = pd.to_numeric(data_sheet.iloc[cash_flow_idx, 22:], errors='coerce').fillna(0) if cash_flow_idx is not None else pd.Series([0] * (len(data_sheet.columns) - 22))

# Create the dashboard layout
app.layout = html.Div([
    html.H1("Coliseum Storage Dashboard", style={'textAlign': 'center'}),
    html.Div([
        html.Label("Select Time Range:"),
        dcc.Dropdown(id='time-filter', options=[
            {'label': 'Construction (1-15)', 'value': 'construction'},
            {'label': 'Lease-Up (16-52)', 'value': 'lease-up'},
            {'label': 'All', 'value': 'all'}
        ], value='all'),
    ], style={'padding': '20px'}),
    html.Div([
        html.Div([
            html.H3("Financial Metrics"), 
            html.Div(id='roe'),
            html.Div(id='dscr'),
            html.Div(id='cash-on-cash'),
            html.Div(id='budget-variance'),
            html.Div(id='cost-variance'),
            html.Div(id='break-even'),
            html.H3("Construction Metrics"),
            html.Div(id='construction-variance'),
            html.Div(id='days-behind'),
            html.H3("Operational Metrics"),
            html.Div(id='occupancy-rate'),
            html.Div(id='lease-velocity')
        ], style={'width': '30%', 'float': 'left'}),
        html.Div([
            html.H3("Progress Metrics"), 
            dcc.Graph(id='construction-progress'),
            dcc.Graph(id='cost-progress'),
            dcc.Graph(id='lease-progress'),
            dcc.Graph(id='cash-flow')
        ], style={'width': '65%', 'float': 'right'}),
    ])
])

@app.callback(
    [Output('roe', 'children'), 
     Output('dscr', 'children'),
     Output('cash-on-cash', 'children'),
     Output('budget-variance', 'children'),
     Output('cost-variance', 'children'),
     Output('construction-variance', 'children'),
     Output('days-behind', 'children'),
     Output('occupancy-rate', 'children'),
     Output('lease-velocity', 'children'),
     Output('break-even', 'children'),
     Output('construction-progress', 'figure'),
     Output('cost-progress', 'figure'),
     Output('lease-progress', 'figure'),
     Output('cash-flow', 'figure')],
    [Input('time-filter', 'value')]
)
def update_dashboard(time_filter):
    month_range = range(1, 16) if time_filter == 'construction' else range(16, 53) if time_filter == 'lease-up' else months
    
    if time_filter == 'construction':
        roi = 0.0
        dscr = 0.0
        cash_on_cash = 0.0
    else:
        # Calculate ROE with detailed debug
        latest_revenue = revenue[revenue.ne(0)].iloc[-1] if not revenue.eq(0).all() else 0  # $225,050
        latest_expenses = expenses[expenses.ne(0)].iloc[-1] if not expenses.eq(0).all() else 0  # $81,526
        latest_debt_service = debt_service[debt_service.ne(0)].iloc[-1] if not debt_service.eq(0).all() else 1  # $87,500
        
        monthly_noi = latest_revenue - latest_expenses  # $225,050 - $81,526 = $143,524
        annual_noi = monthly_noi * 12  # $143,524 * 12 = $1,722,288
        annual_debt_service = latest_debt_service * 12  # $87,500 * 12 = $1,050,000
        
        roi = (annual_noi / equity) * 100  # ($1,722,288 / $11,664,124) * 100 = 14.77%
        
        print("\nROE Calculation Components:")
        print(f"Monthly Revenue: ${latest_revenue:,.2f}")
        print(f"Monthly Expenses: ${latest_expenses:,.2f}")
        print(f"Monthly NOI: ${monthly_noi:,.2f}")
        print(f"Annual NOI: ${annual_noi:,.2f}")
        print(f"Annual Debt Service: ${annual_debt_service:,.2f}")
        print(f"Equity: ${equity:,.2f}")
        print(f"ROE: {roi:.2f}%")
        dscr = monthly_noi / latest_debt_service if latest_debt_service != 0 else 0
        
        print(f"DSCR Calculation: {monthly_noi:,.2f} / {latest_debt_service:,.2f} = {dscr:.2f}x")

        # Calculate Cash-on-Cash Return
        annual_cash_flow = annual_noi - annual_debt_service
        cash_on_cash = (annual_cash_flow / equity) * 100 if equity != 0 else 0
        
        print("\nCash-on-Cash Calculation:")
        print(f"Annual NOI: ${annual_noi:,.2f}")
        print(f"Annual Debt Service: ${annual_debt_service:,.2f}")
        print(f"Annual Cash Flow: ${annual_cash_flow:,.2f}")
        print(f"Equity Invested: ${equity:,.2f}")
        print(f"Cash-on-Cash Return: {cash_on_cash:.2f}%")

    # Calculate cost variance
    latest_budget = budget_to_complete[budget_to_complete.ne(0)].iloc[-1] if not budget_to_complete.eq(0).all() else INITIAL_BUDGET
    budget_variance_pct = ((latest_budget - INITIAL_BUDGET) / INITIAL_BUDGET) * 100

    print("\nBudget Variance Calculation:")
    print(f"Initial Budget: ${INITIAL_BUDGET:,.2f}")
    print(f"Latest Budget: ${latest_budget:,.2f}")
    print(f"Budget Variance: {budget_variance_pct:.2f}%")

    # Calculate cost variance
    latest_est_cost = est_cost_to_complete[est_cost_to_complete.ne(0)].iloc[-1] if not est_cost_to_complete.eq(0).all() else 0
    cost_variance = latest_budget - latest_est_cost

    # Calculate lease-up velocity
    latest_actual = units_leased[units_leased.ne(0)].iloc[-1] if not units_leased.eq(0).all() else 0
    latest_planned = planned_units[planned_units.ne(0)].iloc[-1] if not planned_units.eq(0).all() else 1
    lease_velocity = (latest_actual / latest_planned * 100) if latest_planned != 0 else 0

    # Calculate construction progress variance
    latest_actual = gc_cumulative[gc_cumulative.ne(0)].iloc[-1] if not gc_cumulative.eq(0).all() else 0
    latest_projected = dai_cumulative[dai_cumulative.ne(0)].iloc[-1] if not dai_cumulative.eq(0).all() else 0
    
    construction_variance = latest_actual - latest_projected
    
    print("\nConstruction Progress Debug:")
    print(f"Latest Actual Progress: {latest_actual:.2f}%")
    print(f"Latest Projected Progress: {latest_projected:.2f}%")
    print(f"Progress Variance: {construction_variance:.2f}%")

    # Calculate days behind schedule
    if time_filter == 'construction':
        days_behind = 0
    else:
        # Calculate actual days based on progress
        latest_actual_progress = gc_cumulative[gc_cumulative.ne(0)].iloc[-1] if not gc_cumulative.eq(0).all() else 0
        latest_projected_progress = dai_cumulative[dai_cumulative.ne(0)].iloc[-1] if not dai_cumulative.eq(0).all() else 0
        
        # Calculate implied days behind based on progress difference
        progress_difference = latest_projected_progress - latest_actual_progress
        days_behind = (progress_difference / 100) * PLANNED_COMPLETION_DAYS
        
        print("\nSchedule Delay Calculation:")
        print(f"Planned Completion Days: {PLANNED_COMPLETION_DAYS}")
        print(f"Actual Progress: {latest_actual_progress:.2f}%")
        print(f"Projected Progress: {latest_projected_progress:.2f}%")
        print(f"Progress Difference: {progress_difference:.2f}%")
        print(f"Days Behind Schedule: {days_behind:.1f}")

    # Calculate occupancy rate
    latest_units_leased = units_leased[units_leased.ne(0)].iloc[-1] if not units_leased.eq(0).all() else 0
    occupancy_rate = (latest_units_leased / TOTAL_UNITS) * 100

    print("\nOccupancy Rate Calculation:")
    print(f"Latest Units Leased: {latest_units_leased}")
    print(f"Total Units: {TOTAL_UNITS}")
    print(f"Occupancy Rate: {occupancy_rate:.1f}%")

    # Calculate break-even point
    cumulative_cash_flow = cash_flow.cumsum()
    break_even_month = None
    break_even_cash_flow = None
    
    # Convert Series index to zero-based
    cumulative_cash_flow.index = range(len(cumulative_cash_flow))
    
    for month, value in enumerate(cumulative_cash_flow, start=1):
        if value >= 0:
            break_even_month = month
            break_even_cash_flow = value
            break
    
    print("\nBreak-Even Analysis:")
    print(f"Monthly Cash Flows: {cash_flow.tolist()}")
    print(f"Cumulative Cash Flow: {cumulative_cash_flow.tolist()}")
    if break_even_month:
        print(f"Break-Even Month: {break_even_month}")
        print(f"Break-Even Cash Flow: ${break_even_cash_flow:,.2f}")
    else:
        print("Break-Even not yet reached")

    # Create construction progress figure
    progress_fig = go.Figure()
    progress_fig.add_trace(go.Scatter(
        x=list(range(1, len(dai_cumulative) + 1)),
        y=dai_cumulative,
        name='Projected (S-Curve)',
        mode='lines+markers'
    ))
    progress_fig.add_trace(go.Scatter(
        x=list(range(1, len(gc_cumulative) + 1)),
        y=gc_cumulative,
        name='Actual Progress',
        mode='lines+markers'
    ))
    progress_fig.update_layout(
        title='Construction Progress vs. S-Curve',
        xaxis_title='Month',
        yaxis_title='% Complete'
    )

    # Create cost progress figure
    cost_fig = go.Figure()
    cost_fig.add_trace(go.Scatter(
        x=list(range(1, len(budget_to_complete) + 1)),
        y=budget_to_complete,
        name='Budget to Complete',
        mode='lines+markers'
    ))
    cost_fig.add_trace(go.Scatter(
        x=list(range(1, len(est_cost_to_complete) + 1)),
        y=est_cost_to_complete,
        name='Estimated Cost to Complete',
        mode='lines+markers'
    ))
    cost_fig.update_layout(
        title='Budget vs. Estimated Cost to Complete',
        xaxis_title='Month',
        yaxis_title='Amount ($)'
    )

    # Create lease-up progress figure
    lease_fig = go.Figure()
    lease_fig.add_trace(go.Scatter(
        x=list(range(1, len(planned_units) + 1)),
        y=planned_units,
        name='Planned Units',
        mode='lines+markers'
    ))
    lease_fig.add_trace(go.Scatter(
        x=list(range(1, len(units_leased) + 1)),
        y=units_leased,
        name='Actual Units',
        mode='lines+markers'
    ))
    lease_fig.update_layout(
        title='Lease-up Progress',
        xaxis_title='Month',
        yaxis_title='Units Leased'
    )

    # Create cash flow figure
    cash_flow_fig = go.Figure()
    cash_flow_fig.add_trace(go.Scatter(
        x=list(range(1, len(cumulative_cash_flow) + 1)),
        y=cumulative_cash_flow,
        name='Cumulative Cash Flow',
        mode='lines+markers'
    ))
    if break_even_month and break_even_cash_flow:
        cash_flow_fig.add_vline(
            x=break_even_month,
            line_dash="dash",
            line_color="green",
            annotation_text=f"Break-Even Month {break_even_month}<br>${break_even_cash_flow:,.2f}"
        )
    cash_flow_fig.update_layout(
        title='Cumulative Cash Flow',
        xaxis_title='Month',
        yaxis_title='Cumulative Cash Flow ($)'
    )

    # Debug prints
    print("\nMetrics Debug:")
    print(f"ROE: {roi:.2f}%")
    print(f"DSCR: {dscr:.2f}x")
    print(f"Cash-on-Cash Return: {cash_on_cash:.2f}%")
    print(f"Budget Variance: {budget_variance_pct:+.2f}%")
    print(f"Cost Variance: ${cost_variance:,.2f}")
    print(f"Lease-up Velocity: {lease_velocity:.2f}%")

    return (
        html.Div([
            html.P(f"ROE: {roi:.2f}%", style={'fontSize': '18px'}),
        ]),
        html.Div([
            html.P(f"DSCR: {dscr:.2f}x", style={'fontSize': '18px'}),
        ]),
        html.Div([
            html.P(f"Cash-on-Cash Return: {cash_on_cash:.2f}%", 
                  style={
                      'fontSize': '18px',
                      'color': 'green' if cash_on_cash >= 10 else 'red'
                  }),
        ]),
        html.Div([
            html.P(f"Budget Variance: {budget_variance_pct:+.2f}%", 
                  style={
                      'fontSize': '18px',
                      'color': 'green' if budget_variance_pct <= 0.1 else 'red',
                      'fontWeight': 'bold' if abs(budget_variance_pct) > 1 else 'normal'
                  }),
        ]),
        html.Div([
            html.P(f"Cost Variance: ${cost_variance:,.2f}", 
                  style={
                      'fontSize': '18px',
                      'color': 'green' if cost_variance >= 0 else 'red'
                  }),
        ]),
        html.Div([
            html.P(f"Break-Even Point: {'Month ' + str(break_even_month) if break_even_month else 'Not Yet Reached'}",
                  style={
                      'fontSize': '18px',
                      'color': 'green' if break_even_month and break_even_month <= 40 else 'red',
                      'fontWeight': 'bold'
                  }),
        ]),
        html.Div([
            html.P(f"Construction Progress Variance: {construction_variance:.2f}%",
                  style={
                      'fontSize': '18px',
                      'color': 'green' if construction_variance >= 0 else 'red'
                  }),
        ]),
        html.Div([
            html.P(f"Days Behind Schedule: {days_behind:.1f} days",
                  style={
                      'fontSize': '18px',
                      'color': 'red' if days_behind > 0 else 'green',
                      'fontWeight': 'bold' if days_behind > 90 else 'normal'
                  }),
        ]),
        html.Div([
            html.P(f"Occupancy Rate: {occupancy_rate:.1f}%",
                  style={
                      'fontSize': '18px',
                      'color': 'green' if occupancy_rate >= 46.4 else 'red',
                      'fontWeight': 'bold'
                  }),
        ]),
        html.Div([
            html.P(f"Lease-up Velocity: {lease_velocity:.1f}%",
                  style={
                      'fontSize': '18px',
                      'color': 'green' if lease_velocity >= 100 else 'red'
                  }),
        ]),
        progress_fig,
        cost_fig,
        lease_fig,
        cash_flow_fig
    )

if __name__ == '__main__':
    app.run_server(debug=True)