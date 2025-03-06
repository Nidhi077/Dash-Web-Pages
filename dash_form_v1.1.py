# -*- coding: utf-8 -*-
"""
Created on Thu Dec  7 20:36:31 2023

@author: Nidhi.Dubey
"""
# this version is a directly fillable form with dropdrown list for choosing answers.
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
        id='data-table',
        columns=[{'name': col, 'id': col} for col in df.columns],
        data=df.to_dict('records'),
        editable=True,
    ),
    
    # Dropdowns for user input
    html.Div([
        html.Label('Select Details'),
        dcc.Dropdown(id='input-detail',
                     options=[
                         {'label': 'Option 1', 'value': 'Option 1'},
                         {'label': 'Option 2', 'value': 'Option 2'},
                         # Add more options as needed
                     ],
                     multi=False,
                     value=None
        ),
        
        html.Label('Select Comments'),
        dcc.Dropdown(id='input-comment',
                     options=[
                         {'label': 'Comment 1', 'value': 'Comment 1'},
                         {'label': 'Comment 2', 'value': 'Comment 2'},
                         # Add more options as needed
                     ],
                     multi=False,
                     value=None
        ),
        
        html.Button('Submit', id='submit-button')
    ]),
    
    html.Div(id='output-message')
])

# Callback to update the Excel file with user input
@app.callback(
    Output('output-message', 'children'),
    [Input('submit-button', 'n_clicks')],
    [State('data-table', 'data'),
     State('input-detail', 'value'),
     State('input-comment', 'value')]
)
def update_excel(n_clicks, table_data, detail_value, comment_value):
    if n_clicks is not None:
        # Convert the table data to a DataFrame
        df_new = pd.DataFrame(table_data)
        
        # Update the 'Details' and 'Comments' columns with user input
        df_new['Details'] = detail_value
        df_new['Comments'] = comment_value
        
        # Save the updated DataFrame to Excel
        df_new.to_excel('updated_Build_Order.xlsx', index=False)
        
        return 'Data updated in the Excel file.'

    # If no button click, return no message
    return ''

if __name__ == '__main__':
    app.run_server(debug=True)