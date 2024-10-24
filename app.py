import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
import datetime
from openpyxl import load_workbook

# Load the Excel file and specific sheets
excel_file = "./NPI_Tracking.xlsx"
localization_df = pd.read_excel(excel_file, sheet_name="Localization")
others_df = pd.read_excel(excel_file, sheet_name="Others")
energy_df = pd.read_excel(excel_file, sheet_name="Energy")

# Convert date columns to datetime format
date_columns = ["RFQ send date", "DFM close date", "Biz award date", 
                "Line installation date", "Line readiness date", 
                "First off process date", "C exit date", "TQP date"]

for df in [localization_df, others_df, energy_df]:
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')  # Convert to datetime, coerce errors

# Filter data where "Risk Level" is not "Closed"
def filter_risk_level(df):
    return df[df["Risk Level"] != "Closed"]

# Initialize the Dash app
app = dash.Dash(__name__)

# Create a function to generate the plot based on the selected dataframe
def create_timeline_plot(df):
    # Filter out rows where "Risk Level" is "Closed"
    df = filter_risk_level(df)
    
    # May cause error if you don't convert. Get today's date as a string in the YYYY-MM-DD format
    today = datetime.datetime.now().strftime('%Y-%m-%d')

    # Another date, you can add more
    specific_date = "2024-06-29"

    # Melt the dataframe to make it long format for Plotly
    df_melted = df.melt(
        id_vars=["SIE", "Project", "Risk Level"],
        value_vars=date_columns,
        var_name="Milestone",
        value_name="Date"
    )
    
    df_melted['Date'] = pd.to_datetime(df_melted['Date'], errors='coerce')

    # Concatenate the Project, SIE, and Risk Level columns for the y-axis
    df_melted['Project_SIE_Risk'] = df_melted["Project"] + "_" + df_melted["SIE"] + " (" + df_melted["Risk Level"] + ")"

    # Create the scatter plot for the milestones
    fig = px.scatter(df_melted, x="Date", y="Project_SIE_Risk", 
                     color="Milestone", hover_data=["Milestone", "Date"])

    # Manually add vertical lines
    fig.update_layout(
        shapes=[
            # Vertical line for today's date
            dict(
                type="line",
                xref="x",
                yref="paper",
                x0=today, x1=today,  # Using today's date for the vertical line
                y0=0, y1=1,
                line=dict(color="red", width=2, dash="dash"),
            ),
            # Vertical line for June 29, 2024
            dict(
                type="line",
                xref="x",
                yref="paper",
                x0=specific_date, x1=specific_date,  # Using specific date for the vertical line
                y0=0, y1=1,
                line=dict(color="blue", width=2, dash="dash"),
            )
        ],
        annotations=[
            # Annotation for today's date, positioned slightly above the plot
            dict(
                x=today,
                y=1.05,  
                xref="x",
                yref="paper",
                text=f"Today: {today}",
                showarrow=False,
                font=dict(size=10, color="black"),
                bgcolor="white",
                bordercolor="black",
                borderwidth=1
            ),
            # Annotation for the specific date
            dict(
                x=specific_date,
                y=1.05,  
                xref="x",
                yref="paper",
                text=f"June 29, 2024",
                showarrow=False,
                font=dict(size=10, color="black"),
                bgcolor="white",
                bordercolor="black",
                borderwidth=1
            )
        ]
    )

    # Ensure x-axis is formatted correctly for dates
    fig.update_xaxes(type='date', tickformat="%Y-%m-%d")

    fig.update_layout(title="Project Timeline", xaxis_title="Date", yaxis_title="Projects", height=800)

    return fig


# App layout
app.layout = html.Div([
    html.H1("Project Timeline Dashboard"),
    
    # Dropdown for selecting sheets (tabs)
    dcc.Dropdown(
        id='sheet-dropdown',
        options=[
            {'label': 'Localization', 'value': 'Localization'},
            {'label': 'Others', 'value': 'Others'},
            {'label': 'Energy', 'value': 'Energy'}
        ],
        value='Localization',  # Default selection is Localization
        clearable=False
    ),
    
    dcc.Graph(id='timeline-graph'),
    
    html.Div(id='output-text'),
    
    # Input for date modification
    html.Div([
        html.H3("Modify Date for Selected Milestone:"),
        dcc.Input(id='new-date-input', type='text', placeholder="YYYY-MM-DD"),
        html.Button('Submit', id='submit-date-btn', n_clicks=0)
    ], style={'display': 'none'}, id='date-input-section')
])

@app.callback(
    [Output('date-input-section', 'style'),
     Output('new-date-input', 'value'),
     Output('output-text', 'children')],
    [Input('timeline-graph', 'clickData'),
     Input('submit-date-btn', 'n_clicks')],
    [State('new-date-input', 'value'),
     State('timeline-graph', 'clickData'),
     State('sheet-dropdown', 'value')]
)
def handle_date_modification(clickData, n_clicks, new_date, click_data_state, sheet_name):
    ctx = dash.callback_context
    
    if not ctx.triggered:
        return {'display': 'none'}, '', ''
    
    if 'timeline-graph.clickData' in ctx.triggered[0]['prop_id']:
        if clickData:
            clicked_milestone = clickData['points'][0]['customdata'][0]  # Use custom data for Milestone
            current_date = clickData['points'][0]['x']
            selected_project = clickData['points'][0]['y']
            
            # Split the selected project back into Project, SIE, and Risk Level
            try:
                selected_project_split = selected_project.rsplit(" (", 1)
                project_sie = selected_project_split[0]  # This is "Project_SIE"
                risk_level = selected_project_split[1].rstrip(")")  # Extract risk level, remove trailing parenthesis
                
                project_name, sie = project_sie.split("_")  # Split Project_SIE
                
            except (IndexError, ValueError):
                return {'display': 'none'}, '', 'Error: Unable to parse the selected project.'

            # Select the correct DataFrame
            if sheet_name == 'Localization':
                df = localization_df
            elif sheet_name == 'Others':
                df = others_df
            else:
                df = energy_df

            # Filter the DataFrame based on Project, SIE, and Risk Level
            project_info = df[(df["Project"] == project_name) & (df["SIE"] == sie) & (df["Risk Level"] == risk_level)]
            
            if project_info.empty:
                return {'display': 'none'}, '', f'Error: No matching project found for {selected_project}'

            next_step_plan = project_info["Next step plan"].values[0]
            action_items = project_info["Action Items for Cindy"].values[0]
            
            return {'display': 'block'}, current_date, html.Div([
                html.H4(f"Modify the date for {clicked_milestone} of {selected_project}"),
                html.P(f"Next Step Plan: {next_step_plan}"),
                html.P(f"Action Items for Cindy: {action_items}")
            ])

    elif 'submit-date-btn.n_clicks' in ctx.triggered[0]['prop_id']:
        if n_clicks > 0 and click_data_state:
            selected_project = click_data_state['points'][0]['y']
            milestone = click_data_state['points'][0]['customdata'][0]  # Use custom data for Milestone
            
            # Split the selected project back into Project, SIE, and Risk Level
            try:
                selected_project_split = selected_project.rsplit(" (", 1)
                project_sie = selected_project_split[0]
                risk_level = selected_project_split[1].rstrip(")")
                
                project_name, sie = project_sie.split("_")
            except (IndexError, ValueError):
                return {'display': 'none'}, '', 'Error: Unable to parse the selected project.'
            
            # Select the correct DataFrame
            if sheet_name == 'Localization':
                df = localization_df
            elif sheet_name == 'Others':
                df = others_df
            else:
                df = energy_df

            # Update the corresponding date in the DataFrame
            df.loc[(df["Project"] == project_name) & (df["SIE"] == sie) & (df["Risk Level"] == risk_level), milestone] = pd.to_datetime(new_date, format='%Y-%m-%d', errors='coerce')

            # Convert the date columns back to short format when saving
            df[date_columns] = df[date_columns].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.strftime('%Y-%m-%d') if x.dtype == 'datetime64[ns]' else x)

            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

            return {'display': 'none'}, '', f"The date for {milestone} of {selected_project} has been updated to {new_date}."
    
    return {'display': 'none'}, '', ''


# Callback to update the graph based on selected sheet
@app.callback(
    Output('timeline-graph', 'figure'),
    [Input('sheet-dropdown', 'value')]
)
def update_graph(sheet_name):
    if sheet_name == 'Localization':
        df = localization_df
    elif sheet_name == 'Others':
        df = others_df
    else:
        df = energy_df
    return create_timeline_plot(df)

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)

