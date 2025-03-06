#this version is a row wise fillable form
# -*- coding: utf-8 -*-
import dash
from dash.dependencies import Input, Output
from dash import html, Dash
from dash import dcc
import pandas as pd

# Load the Excel file into a Pandas DataFrame
df = pd.read_excel('Build_order.xlsx')

# Find the row index containing additional comments
comments_row_index = df[df['Work Order #'] == 'Additional Comments'].index[0]
                                                                      
# Initialize the Dash app
app = Dash(__name__)

# Layout of the Dash app
app.layout = html.Div([
    html.H1('Data Input Dashboard'),
    
    # Display existing data
    html.Div([
        html.Table([
            html.Thead(html.Tr([html.Th(col) for col in df.columns])),
            html.Tbody([html.Tr([html.Td(df.iloc[i][col]) for col in df.columns]) for i in range(len(df))])
        ])
    ]),
    
    # Input fields for users to fill in data for the remaining columns
    html.Div([
        html.Label('Select Row:'),
        #dcc.Dropdown(id='row-selector', options=[{'label': str(row), 'value': row} for row in df['Work Order #']]),
        #dcc.Dropdown(id='row-selector', options=[{'label': str(row), 'value': row} for row in df['Work Order #']],
                     #value=df['Work Order #'].iloc[0]),
        
        #dcc.Dropdown(id='row-selector', options=[{'label': str(row), 'value': row} for row in df['Work Order #']],
                     #value=df['Work Order #'].iloc[0],  # Set default value to the first row
                     #clearable=False),
        #dcc.Dropdown(id='row-selector', options=[{'label': str(row), 'value': row} for row in df['Work Order #']],
                     #value=df['Work Order #'].iloc[0]),
             
        dcc.Dropdown(id='row-selector', options=[{'label': str(row), 'value': row} for row in df['Work Order #'][:comments_row_index].astype(str)],
                     value=str(df['Work Order #'].iloc[0]),  # Set default value to the first row
                     clearable=False),            
        html.Label('Details'),
        dcc.Input(id='input-detail', type='text'),
        html.Label('Comments'),
        dcc.Input(id='input-comment', type='text'),
        html.Button('Submit', id='submit-button')
    ]),
    
    html.Div(id='output-message')
])

# Callback to update the Excel file with user input
@app.callback(
    dash.dependencies.Output('output-message', 'children'),
    [dash.dependencies.Input('submit-button', 'n_clicks')],
    [dash.dependencies.State('row-selector', 'value')],
    [dash.dependencies.State('input-detail', 'value'),
     dash.dependencies.State('input-comment', 'value')]
)
def update_dataframe(n_clicks, selected_row, detail_value, comment_value):
    global df  # Use the global DataFrame

    if n_clicks is not None and selected_row is not None:
        # Update the specified row with the new values
        selected_row = int(selected_row)
        df.loc[df['Work Order #'] == selected_row, 'Details'] = detail_value
        df.loc[df['Work Order #'] == selected_row, 'Comments'] = comment_value
        
        # Save the updated DataFrame to Excel
        df.to_excel('updated_Build_Order.xlsx', index=False)
        
        return 'Data updated in the Excel file.'

if __name__ == '__main__':
    app.run_server(debug=True)