import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objs as go
from sklearn.linear_model import LinearRegression
import numpy as np
import os

# Define file path
file_path = r"C:\Users\dthom\OneDrive\Personal\Hart Advisors Group\GitHub\Construction_Management_Tool\(2021.08.04) Coliseum Storage Development Tracker June 2021.xlsx"

# Load Excel file
try:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found at: {file_path}")
    df = pd.read_excel(file_path, sheet_name=None)
    print("Excel file loaded successfully.")
    print("Available sheets:", list(df.keys()))
except Exception as e:
    print(f"Error loading file: {e}")
    exit(1)

# Function to find row index by label in a specific column
def find_row_index(df, label, column_idx=None):
    if column_idx is not None:
        for i, val in enumerate(df.iloc[:, column_idx].astype(str)):
            if label.strip().lower() in val.strip().lower():
                return i
        return None
    for col in df.columns:
        for i, val in enumerate(df[col].astype(str)):
            if label.strip().lower() in val.strip().lower():
                return i, df.columns.get_loc(col)
    return None, None

# Select sheets
data_sheet = df['Variance']  # Financial data
s_curve_sheet = df['S Curve']  # S-Curve data

# Extract financial data from 'Variance'
revenue_idx = find_row_index(data_sheet, "Total Operating Revenues", 1)
revenue = data_sheet.iloc[revenue_idx + 1:, 22:].fillna(0).infer_objects(copy=False) if revenue_idx is not None else pd.DataFrame(0, index=[0], columns=range(22, data_sheet.shape[1]))

expenses_idx = find_row_index(data_sheet, "Total Operating Expenses", 1)
expenses = data_sheet.iloc[expenses_idx + 1:, 22:].fillna(0).infer_objects(copy=False) if expenses_idx is not None else pd.DataFrame(0, index=[0], columns=range(22, data_sheet.shape[1]))

cash_flow_idx = find_row_index(data_sheet, "Operating Cash Flow After Reserves", 1)
cash_flow = data_sheet.iloc[cash_flow_idx + 1:, 22:].fillna(0).infer_objects(copy=False) if cash_flow_idx is not None else pd.DataFrame(0, index=[0], columns=range(22, data_sheet.shape[1]))

units_leased_idx = find_row_index(data_sheet, "Cumulative Units Leased", 1)
units_leased = data_sheet.iloc[units_leased_idx + 1:, 22:].fillna(0).infer_objects(copy=False) if units_leased_idx is not None else pd.DataFrame(0, index=[0], columns=range(22, data_sheet.shape[1]))

# Extract S-Curve from 'S Curve' (adjusted logic)
s_curve_indices = []
for i, val in enumerate(s_curve_sheet.iloc[:, 0].astype(str)):
    if "cumulative" in val.strip().lower():
        s_curve_indices.append(i)
if len(s_curve_indices) >= 2:
    s_curve_proj = s_curve_sheet.iloc[s_curve_indices[0] + 1, 1:19].fillna(0).infer_objects(copy=False)
    s_curve_act = s_curve_sheet.iloc[s_curve_indices[1] + 1, 1:19].fillna(0).infer_objects(copy=False)
    print("S-Curve data extracted: Projected:", s_curve_proj.values, "Actual:", s_curve_act.values)
else:
    print("Warning: Insufficient 'Cumulative' rows in S Curve sheet. Checking Variance instead.")
    variance_s_curve_idx = find_row_index(data_sheet, "Cumulative", 1)
    if variance_s_curve_idx is not None and variance_s_curve_idx + 2 < len(data_sheet):
        s_curve_proj = data_sheet.iloc[variance_s_curve_idx + 1, 22:40].fillna(0).infer_objects(copy=False)
        s_curve_act = data_sheet.iloc[variance_s_curve_idx + 2, 22:40].fillna(0).infer_objects(copy=False)
        print("S-Curve data from Variance: Projected:", s_curve_proj.values, "Actual:", s_curve_act.values)
    else:
        print("Warning: No valid S-Curve data found. Using zeros.")
        s_curve_proj = pd.Series([0] * 18)
        s_curve_act = pd.Series([0] * 18)

# Constants
equity = 11664124
total_units = 1098
initial_budget = 27664124
revised_budget = 27688502
budget_to_complete = 8046444
debt_service = 170031
months = list(range(1, 53))

# Predictive Model
X = np.array(range(16, 34)).reshape(-1, 1)
y = units_leased.iloc[0, :18].values if units_leased.shape[1] >= 18 else np.zeros(18)
model = LinearRegression().fit(X, y)
forecast = model.predict(np.array(range(16, 53)).reshape(-1, 1))

# Initialize Dash app
app = dash.Dash(__name__)

# Layout
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
        html.Div([html.H3("Financial Metrics"), html.Div(id='roe')], style={'width': '30%', 'float': 'left'}),
        html.Div([html.H3("Progress Metrics"), dcc.Graph(id='construction-progress')], style={'width': '65%', 'float': 'right'}),
    ])
])

# Callback
@app.callback(
    [Output('roe', 'children'), Output('construction-progress', 'figure')],
    [Input('time-filter', 'value')]
)
def update_dashboard(time_filter):
    month_range = range(1, 16) if time_filter == 'construction' else range(16, 53) if time_filter == 'lease-up' else months
    adjusted_units = [min(30 * (i - 15), total_units) for i in month_range if i >= 16] or [0]
    adjusted_revenue = [u * 214.27 for u in adjusted_units]
    noi = adjusted_revenue[-1] - expenses.iloc[0, min(len(expenses.columns)-1, len(adjusted_revenue)-1)] if adjusted_revenue else 0
    annual_noi = noi * 12
    roe = (annual_noi / equity) * 100

    progress_fig = go.Figure()
    progress_fig.add_trace(go.Scatter(x=list(range(1, 19)), y=s_curve_proj, name='Projected'))
    progress_fig.add_trace(go.Scatter(x=list(range(1, 19)), y=s_curve_act, name='Actual'))
    progress_fig.update_layout(title='Construction Progress vs. S-Curve', xaxis_title='Month', yaxis_title='% Complete')

    return html.Span(f"ROE: {roe:.2f}%"), progress_fig

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)