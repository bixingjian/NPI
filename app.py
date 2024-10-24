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
    
    # Get today's date as a string in the YYYY-MM-DD format
    today = datetime.datetime.now().strftime('%Y-%m-%d')

    # June 29, 2024
    specific_date = "2024-06-29"

    # Melt the dataframe to get all date columns in one column
    df_melted = df.melt(
        id_vars=["SIE", "Project"],
        value_vars=date_columns,
        var_name="Milestone",
        value_name="Date"
    )
    
    # Ensure that the Date column is recognized as datetime by Plotly
    df_melted['Date'] = pd.to_datetime(df_melted['Date'], errors='coerce')

    # Create the scatter plot for the milestones
    fig = px.scatter(df_melted, x="Date", y=df_melted["Project"] + "_" + df_melted["SIE"], 
                     color="Milestone", hover_data=["Milestone", "Date"])

    # Manually add vertical lines for today's date and the specific date using shapes
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
                y=1.05,  # Positioning the annotation slightly higher above the plot
                xref="x",
                yref="paper",
                text=f"Today: {today}",
                showarrow=False,
                font=dict(size=10, color="black"),
                bgcolor="white",
                bordercolor="black",
                borderwidth=1
            ),
            # Annotation for the specific date, positioned slightly above the plot
            dict(
                x=specific_date,
                y=1.05,  # Positioning the annotation slightly higher above the plot
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

    fig.update_layout(title="Project Timeline", xaxis_title="Date", yaxis_title="Projects")

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

# Callback to handle both showing input field and modifying the date
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
    
    # Check which input triggered the callback
    if not ctx.triggered:
        return {'display': 'none'}, '', ''
    
    # If a point was clicked in the graph
    if 'timeline-graph.clickData' in ctx.triggered[0]['prop_id']:
        if clickData:
            # Extract milestone and current date from clickData
            clicked_milestone = clickData['points'][0]['customdata'][0]  # Use custom data for Milestone
            current_date = clickData['points'][0]['x']
            selected_project = clickData['points'][0]['y']
            
            # Retrieve "Next Step Plan" and "Action Items for Cindy" for the selected project
            if sheet_name == 'Localization':
                df = localization_df
            elif sheet_name == 'Others':
                df = others_df
            else:
                df = energy_df

            project_info = df[df["Project"] + "_" + df["SIE"] == selected_project]
            next_step_plan = project_info["Next step plan"].values[0]
            action_items = project_info["Action Items for Cindy"].values[0]
            
            return {'display': 'block'}, current_date, html.Div([
                html.H4(f"Modify the date for {clicked_milestone} of {selected_project}"),
                html.P(f"Next Step Plan: {next_step_plan}"),
                html.P(f"Action Items for Cindy: {action_items}")
            ])

    # If the submit button was clicked
    elif 'submit-date-btn.n_clicks' in ctx.triggered[0]['prop_id']:
        if n_clicks > 0 and click_data_state:
            selected_project = click_data_state['points'][0]['y']
            milestone = click_data_state['points'][0]['customdata'][0]  # Use custom data for Milestone
            
            # Modify the corresponding milestone date in the appropriate DataFrame
            if sheet_name == 'Localization':
                df = localization_df
            elif sheet_name == 'Others':
                df = others_df
            else:
                df = energy_df

            # Update the corresponding date without the time component
            df.loc[df["Project"] + "_" + df["SIE"] == selected_project, milestone] = pd.to_datetime(new_date, errors='coerce').date()

            # Convert the date columns back to short format when saving, ensuring that only datetime types are processed
            df[date_columns] = df[date_columns].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.strftime('%Y-%m-%d') if x.dtype == 'datetime64[ns]' else x)

            # Save the updated DataFrame back to the Excel file
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