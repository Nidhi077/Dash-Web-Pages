# -*- coding: utf-8 -*-
"""
Created on Wed Mar 27 22:22:14 2024

@author: Nidhi.Dubey
"""

import dash
from dash import html, dcc
from dash.dependencies import Input, Output, State
from dash import dash_table
import pandas as pd

# Load the Excel file into a Pandas DataFrame
df = pd.read_excel('updated_Build_order.xlsx')

# Dropdown options for 'details' and 'comments' columns
details_options = ['1030 mm x 12mm dia with 100 mm thread length', 'Option 2', 'Option 3']
comments_options = ['Shank Dia= 12 mm; - Teflon coated, Forged Bolts, and heat shrink and ptfe tube is used', '4.5 mm silicon  O ring gasket(Low compression set)', 'USD 1.5  with flat PPS on back & FF on the back sidewith metal extended part.'
                    ,'Sunrise 2.1 with 140 um flow field tab', 'MAR MEA','S2 Coated']

# Function to create dropdown component
def create_dropdown(id):
    return dcc.Dropdown(
        id={'type': 'dropdown', 'index': id},
        options=[{'label': opt, 'value': opt} for opt in details_options if opt],
        placeholder="Select details",
        value=''
    )

# Function to create dropdown component
def create_comment_dropdown(id):
    return dcc.Dropdown(
        id={'type': 'comment-dropdown', 'index': id},
        options=[{'label': opt, 'value': opt} for opt in comments_options if opt],
        placeholder="Select comment",
        value=''
    )

# Initialize the Dash app
app = dash.Dash(__name__)

# Layout of the Dash app
app.layout = html.Div([
    html.H1('Data Input Dashboard'),
    
    # Display existing data using DataTable
    dash_table.DataTable(
        id='data-table',
        columns=[
            {'name': col, 'id': col, 'editable': True} if col not in ['Details', 'Comments']
            else {'name': col, 'id': col, 'presentation': 'dropdown'} for col in df.columns
        ],
        data=df.to_dict('records'),
        editable=True,
    ),
    
    html.Div(id='dropdown-container', style={'display': 'none'}),
    
    # Submit button
    html.Button('Submit', id='submit-button'),
    
    html.Div(id='output-message')
])

# Callback to update the Excel file with user input
@app.callback(
    Output('dropdown-container', 'children'),
    [Input('data-table', 'data')]
)
def update_dropdowns(data):
    dropdowns = []
    for i, row in enumerate(data):
        if row['Details'] == '':
            dropdowns.append(create_dropdown(i))
        if row['Comments'] == '':
            dropdowns.append(create_comment_dropdown(i))
    return dropdowns

if __name__ == '__main__':
    app.run_server(debug=True)