import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import pandas as pd
from datetime import datetime as dt
import dash_bootstrap_components as dbc

dbc_css = "https://cdn.jsdelivr.net/gh/AnnMarieW/dash-bootstrap-templates/dbc.min.css"

app = dash.Dash(
    __name__,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}],
    external_stylesheets=[dbc.themes.LITERA, dbc_css]
)
app.title = "Force Status Summary Report"


df = pd.read_excel("Summary\Checking Summary.xlsx", sheet_name="Sheet1")

# Layout of the app
app.layout = html.Div([
    dcc.Interval(
        id='interval-component',
        interval=2000,  # Update every 2 seconds
        n_intervals=0
    ),
    html.Div(
        id="description-card",
        children=[
            html.Div([html.H5("")], style={"width":"25vw"}),
            html.Div([html.H3("Force Status Summary Report", className="text-light fw-bold text-center")], style={"width":"35vw"},),
            html.Div(
                id="intro",
                children="",
                style={"width":"25vw"}
            ),
        ], className="hstack d-flex align-items-center justify-content-between border rounded-bottom border-dark", 
        style={"background":"#222831", "boxShadow": "0px 4px 8px rgba(0, 0, 0, 0.3)"}),
    html.Br(),
    html.Div([
        html.Div([
            html.Label('Cluster ID', className="text-light fw-bold"),
            dcc.Dropdown(
            options = [{'label': value, 'value': value} for value in df['Cluster ID'].unique()],
            placeholder="Cluster ID",
            clearable=True,
            id = "cluster-id-dropdown",
            style = {"boxShadow": "0px 4px 8px rgba(0, 0, 0, 0.5)"})], style={"width":"25vw", }),
        html.Div([
            html.Label('Checking Date', className="text-light fw-bold"),
            dcc.DatePickerRange(
                id="date-picker-select",
                initial_visible_month=dt.now(),
                clearable=True,
                style = {"boxShadow": "0px 4px 8px rgba(0, 0, 0, 0.5)"}
            )], style={"width":"25vw"}),
        html.Div([
            html.Label('Status', className="text-light fw-bold"),
            dcc.Dropdown(
                options = [{'label': 'REVISE', 'value': 'REVISE'},
                        {'label': 'OK', 'value': 'OK'}],
                placeholder="Status",
                clearable=True,
                id="status-dropdown",
                style = {"boxShadow": "0px 4px 8px rgba(0, 0, 0, 0.5)"})], style={"width":"25vw"})], className="hstack gap-2 d-flex justify-content-around align-items-center"),
    html.Br(),
    html.Div([
        dash_table.DataTable(data = df.to_dict('records'),
                             columns = [{"name": i, "id": i} for i in df.columns],
                             id="data-table",
                             sort_action="native")
    ], style={"width":"90vw","overflowY":"auto", 'overflowX': 'auto', "boxShadow": "0px 4px 8px rgba(0, 0, 0, 0.5)"}, className="dbc align-self-center border rounded")
], className="vstack gap-3", style={"background":"#31363F", "height":"50vw"})

@app.callback(
    [Output("cluster-id-dropdown", "options"),
     Output("status-dropdown", "options")],
    [Input("data-table", "data")]
)
def update_dropdown(dict):
    df = pd.DataFrame(dict)
    if len(df) != 0:
        options_cluster = [{'label': value, 'value': value} for value in df['Cluster ID'].unique()]
        options_status = [{'label': value, 'value': value} for value in df['Status'].unique()]
    else:
        options_cluster = []
        options_status = []
    return options_cluster, options_status

# Callback to update the DataTable based on filter input
@app.callback(
    Output('data-table', 'data'),
    [Input('interval-component', 'n_intervals'),
     Input('cluster-id-dropdown', 'value'),
     Input("date-picker-select", "start_date"),
     Input("date-picker-select", "end_date"),
     Input('status-dropdown', 'value'),
     ]
)
def update_table(n, cluster_id, start_date, end_date, status):
    df = pd.read_excel("Summary\Checking Summary.xlsx", sheet_name="Sheet1")

    if cluster_id:
        df = df[df['Cluster ID'] == cluster_id]
    if start_date and end_date:
        df = df[(df['Checking Date'] >= start_date) & (df['Checking Date'] <= end_date)]
    if status:
        df = df[df['Status'] == status]
    return df.to_dict('records')

if __name__ == '__main__':
    app.run_server(debug=True)