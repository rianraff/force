from dash import Dash, html, dash_table, dcc
import pandas as pd
from dash.dependencies import Input, Output

with pd.ExcelWriter("Summary/Checking Summary.xlsx", 'openpyxl', mode='a',  if_sheet_exists="overlay") as writer:
    # fix line
    reader = pd.read_excel("Summary/Checking Summary.xlsx", sheet_name="Sheet1")

app = Dash(__name__)


# Layout of the app
app.layout = html.Div([
    dcc.Interval(
        id='interval-component',
        interval=2000,  # Update every 2 seconds
        n_intervals=0
    ),
    html.Div(id='table-container', children=[
        dash_table.DataTable(data=reader.to_dict('records'), page_size=10)
    ])
])

# Callback to update the DataTable
@app.callback(
    Output('table-container', 'children'),
    [Input('interval-component', 'n_intervals')]
)
def update_table(n):
    # You can fetch new data here
    # For demonstration, let's just add a new row to the existing DataFrame
    reader = pd.read_excel("Summary/Checking Summary.xlsx", sheet_name="Sheet1")
    # Create the DataTable with updated data
    table = dash_table.DataTable(
        id='datatable',
        data=reader.to_dict('records')
    )

    return table

if __name__ == '__main__':
    app.run_server(debug=True)
