import dash
from dash import dcc, html, Input, Output, State, dash_table
import pandas as pd
import plotly.express as px
from datetime import date
import uuid
import webbrowser

# Initialize Dash app
app = dash.Dash(__name__)
app.title = "Employee Absence Management"

# Initialize data
initial_employees = pd.DataFrame(columns=["employee_id", "name", "email", "department"])
initial_absences = pd.DataFrame(columns=[
    "absence_id", "employee_id", "absence_start", "absence_end", "absence_duration", "absence_reason", "absence_type"
])

# App layout
app.layout = html.Div([
    html.H1("Employee Absence Management", style={"textAlign": "center"}),

    # Add Employee Section
    html.Div([
        html.H3("Add Employee"),
        html.Div([
            html.Label("Name"), dcc.Input(id="name", type="text", placeholder="Name"),
            html.Label("Email"), dcc.Input(id="email", type="email", placeholder="Email"),
            html.Label("Department"), dcc.Input(id="department", type="text", placeholder="Department")
        ], style={"display": "flex", "gap": "10px"}),
        html.Button("Add Employee", id="add_employee", n_clicks=0),
        html.Div(id="employee_feedback", style={"color": "green", "marginTop": "10px"})
    ], style={"marginBottom": "20px"}),

    # Add Absence Section
    html.Div([
        html.H3("Add Absence"),
        html.Label("Select Employee"),
        dcc.Dropdown(id="select_employee", options=[], placeholder="Select Employee"),
        html.Div([
            html.Label("Absence Start Date"), dcc.DatePickerSingle(id="absence_start", date=date.today()),
            html.Label("Absence End Date"), dcc.DatePickerSingle(id="absence_end", date=date.today()),
            html.Label("Absence Reason"), dcc.Input(id="absence_reason", type="text", placeholder="Reason"),
            html.Label("Absence Type"), dcc.Input(id="absence_type", type="text", placeholder="Type")
        ], style={"display": "flex", "gap": "10px"}),
        html.Button("Add Absence", id="add_absence", n_clicks=0),
        html.Div(id="absence_feedback", style={"color": "green", "marginTop": "10px"})
    ], style={"marginBottom": "20px"}),

    # Save and Load Section
    html.Div([
        html.H3("Save and Load Data"),
        html.Button("Save Employees as CSV", id="save_employees"),
        html.Div(id="employees_feedback_message", style={"marginTop": "10px", "display": "inline-block", "marginLeft": "10px"}),
        dcc.Upload(
            id="upload_employees", children=html.Button("Reimport Employees (CSV)"),
            multiple=False
        ),
        html.Button("Save Absences as CSV", id="save_absences"),
        html.Div(id="absences_feedback_message", style={"marginTop": "10px", "display": "inline-block", "marginLeft": "10px"}),
        dcc.Upload(
            id="upload_absences", children=html.Button("Reimport Absences (CSV)"),
            multiple=False
        )
    ], style={"marginBottom": "20px"}),

    # Tables and Graph
    dcc.Tabs([
        dcc.Tab(label="Employees", children=[
            dash_table.DataTable(id="employee_table", style_table={"overflowX": "auto"})
        ]),
        dcc.Tab(label="Absences", children=[
            dash_table.DataTable(id="absence_table", style_table={"overflowX": "auto"})
        ]),
        dcc.Tab(label="Absence Trends", children=[
            dcc.Graph(id="absence_graph")
        ])
    ])
])

# Global Data
employees = initial_employees.copy()
absences = initial_absences.copy()

# Callbacks
@app.callback(
    Output("employee_feedback", "children"),
    Output("select_employee", "options"),
    Output("employee_table", "data"),
    Input("add_employee", "n_clicks"),
    State("name", "value"), State("email", "value"), State("department", "value"),
    prevent_initial_call="initial_duplicate"
)
def add_employee(n_clicks, name, email, department):
    global employees
    if not name or not email or not department:
        return "Please fill in all fields!", [], employees.to_dict("records")
    
    new_employee = {
        "employee_id": f"EMP-{uuid.uuid4().hex[:8]}",
        "name": name, "email": email, "department": department
    }
    employees = pd.concat([employees, pd.DataFrame([new_employee])], ignore_index=True)
    employee_options = [{"label": row["name"], "value": row["employee_id"]} for _, row in employees.iterrows()]
    return "Employee added successfully!", employee_options, employees.to_dict("records")

@app.callback(
    Output("absence_feedback", "children"),
    Output("absence_table", "data"),
    Input("add_absence", "n_clicks"),
    State("select_employee", "value"), State("absence_start", "date"),
    State("absence_end", "date"), State("absence_reason", "value"), State("absence_type", "value"),
    prevent_initial_call="initial_duplicate"
)
def add_absence(n_clicks, employee_id, absence_start, absence_end, absence_reason, absence_type):
    global absences
    if not employee_id or not absence_start or not absence_end or not absence_reason or not absence_type:
        return "Please fill in all fields!", absences.to_dict("records")
    
    # Calculate duration
    absence_duration = (pd.to_datetime(absence_end) - pd.to_datetime(absence_start)).days

    new_absence = {
        "absence_id": f"ABS-{uuid.uuid4().hex[:8]}",
        "employee_id": employee_id,
        "absence_start": absence_start,
        "absence_end": absence_end,
        "absence_duration": absence_duration,
        "absence_reason": absence_reason,
        "absence_type": absence_type
    }
    absences = pd.concat([absences, pd.DataFrame([new_absence])], ignore_index=True)
    return "Absence added successfully!", absences.to_dict("records")

@app.callback(
    Output("absence_graph", "figure"),
    Input("absence_table", "data"),
    prevent_initial_call=True
)
def update_graph(data):
    df = pd.DataFrame(data)
    if df.empty:
        return px.line(title="No absence data available")
    
    df["absence_start"] = pd.to_datetime(df["absence_start"])
    trends = df.groupby("absence_start").sum().reset_index()
    fig = px.line(trends, x="absence_start", y="absence_duration", title="Absence Trends Over Time",
                  labels={"absence_start": "Date", "absence_duration": "Total Duration (days)"})
    return fig

# Callback to save employees data as CSV
@app.callback(
    Output("employees_feedback_message", "children"),
    Input("save_employees", "n_clicks"),
    prevent_initial_call=True
)
def save_employees_as_csv(n_clicks):
    if n_clicks is None:
        return ""
    
    # Save employees data as CSV
    try:
        employees.to_csv("employees.csv", index=False)
        return "Employee data saved successfully as employees.csv."
    except Exception as e:
        return f"Error saving employees data: {str(e)}"

# Callback to save absences data as CSV
@app.callback(
    Output("absences_feedback_message", "children"),
    Input("save_absences", "n_clicks"),
    prevent_initial_call=True
)
def save_absences_as_csv(n_clicks):
    if n_clicks is None:
        return ""
    
    # Save absences data as CSV
    try:
        absences.to_csv("absences.csv", index=False)
        return "Absences data saved successfully as absences.csv."
    except Exception as e:
        return f"Error saving absences data: {str(e)}"
    
# Run the app
if __name__ == "__main__":
    # Open the app in the default web browser
    webbrowser.open("http://127.0.0.1:8050/")
    app.run_server(debug=True)
