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
