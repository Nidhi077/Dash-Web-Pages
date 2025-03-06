# this version changes the form from being row wise fillable to directly fillable and editable
# -*- coding: utf-8 -*-
import dash
from dash import html, dcc
from dash.dependencies import Input, Output, State
from dash import dash_table
import pandas as pd

# Load the Excel file into a Pandas DataFrame
df = pd.read_excel('updated_Build_order.xlsx')

# Initialize the Dash app
app = dash.Dash(__name__)

# Layout of the Dash app
app.layout = html.Div([
    html.H1('Data Input Dashboard'),
    
    # Display existing data using DataTable
    dash_table.DataTable(
        id ='data-table',
        columns=[{'name': col, 'id': col} for col in df.columns],
        data=df.to_dict('records'),
        editable=True,
    ),
    
    # Submit button
    html.Button('Submit', id='submit-button'),
    
    html.Div(id='output-message')
])

# Callback to update the Excel file with user input
@app.callback(
    Output('output-message', 'children'),
    [Input('submit-button', 'n_clicks')],
    [State('data-table', 'data')]
)
def update_excel(n_clicks, table_data):
    if n_clicks is not None:
        # Convert the table data to a DataFrame
        df_new = pd.DataFrame(table_data)
        
        # Save the updated DataFrame to Excel
        df_new.to_excel('updated_Build_Order.xlsx', index=False)
        
        return 'Data updated in the Excel file.'

    # If no button click, return no message
    return ''

if __name__ == '__main__':
    app.run_server(debug=True)