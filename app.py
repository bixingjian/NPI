import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import datetime

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

# Initialize the Dash app
app = dash.Dash(__name__)

# Create a function to generate the plot based on the selected dataframe
def create_timeline_plot(df):
    # Ensure today is explicitly a datetime object
    today = pd.Timestamp(datetime.datetime.now().strftime('%Y-%m-%d'))

    # Melt the dataframe to get all date columns in one column
    df_melted = df.melt(
        id_vars=["SIE", "Project"],
        value_vars=date_columns,
        var_name="Milestone",
        value_name="Date"
    )
    
    # Ensure that the Date column is recognized as datetime by Plotly
    df_melted['Date'] = pd.to_datetime(df_melted['Date'], errors='coerce')

    fig = px.scatter(df_melted, x="Date", y=df_melted["Project"] + "_" + df_melted["SIE"], 
                     color="Milestone", hover_data=["Milestone", "Date"])

    # Add today's date as a vertical line using the correctly formatted datetime object
    # fig.add_vline(x=today, line_dash="dash", line_color="red", annotation_text="Today", annotation_position="top left")

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
    
    html.Div(id='output-text')
])

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

# Callback to display the next steps and action items when a project is clicked
@app.callback(
    Output('output-text', 'children'),
    Input('timeline-graph', 'clickData')
)
def display_next_steps(clickData):
    if clickData:
        selected_project = clickData['points'][0]['y']
        df_map = {'Localization': localization_df, 'Others': others_df, 'Energy': energy_df}
        
        # Try to find the project in all dataframes
        for sheet, df in df_map.items():
            project_info = df[df["Project"] + "_" + df["SIE"] == selected_project]
            if not project_info.empty:
                next_step = project_info["Next step plan"].values[0]
                action_items = project_info["Action Items for Cindy"].values[0]
                return html.Div([
                    html.H3(f"Project: {selected_project}"),
                    html.P(f"Next Step Plan: {next_step}"),
                    html.P(f"Action Items for Cindy: {action_items}")
                ])
    return "Click on a project to see details."

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
