import dash
from dash import dcc, html, Input, Output
import pandas as pd
import plotly.express as px
from dash.dash_table import DataTable
import calendar

# Carregar o arquivo XLSX para um DataFrame
df = pd.read_excel('tabelaa.xlsx')

# Extrair o mês da coluna 'Accounting Doc: Document Date'
df['Month'] = pd.to_datetime(df['Accounting Doc: Document Date']).dt.month

# Contagem de accounts sem repetição por tipo e mês
contagem = df.groupby(['Nature_new', 'Month']).agg({'Accounting Doc: Number': 'nunique',
                                                    'Accounting Doc: Header Text': 'first',
                                                    'Division': 'first'}).reset_index()

# Renomear colunas
contagem = contagem.rename(
    columns={'Accounting Doc: Number': 'Quantidade', 'Accounting Doc: Header Text': 'Header Text'})

# Criar o aplicativo Dash
app = dash.Dash(__name__)
server = app.server
# Layout do aplicativo
app.layout = html.Div(style={'backgroundColor': '#e6ffe6'}, children=[
    html.Div([
        html.Img(src="https://www.msd.com.br/wp-content/themes/mhh-mhh2-mcc-theme/images/msd-logo.svg", style={'height': '50px'}),
        html.H1('MJE Threshold Compliance', style={'color': '#008000'})
    ], style={'margin-bottom': '20px'}),

    html.Div([
        dcc.Dropdown(
            id='month-dropdown',
            options=[{'label': f"{calendar.month_abbr[month]}/{str(24)}", 'value': month} for month in contagem['Month'].unique()],
            value=contagem['Month'].min(),
            clearable=False,
            style={'width': '50%', 'color': '#008000'}
        )
    ], style={'margin-bottom': '20px'}),

    html.Div([
        html.Div([
            dcc.Graph(id='combined-chart')
        ], style={'width': '49%', 'display': 'inline-block'}),
        html.Div([
            dcc.Graph(id='pie-chart')
        ], style={'width': '49%', 'display': 'inline-block'})
    ]),

    html.Div([
        DataTable(
            id='table',
            columns=[{"name": i, "id": i} for i in contagem.columns],
            data=contagem.to_dict('records'),
            style_header={'backgroundColor': '#008000', 'fontWeight': 'bold'},
            style_cell={'backgroundColor': '#FFFFFF', 'color': '#000000'}
        )
    ])
])


# Callbacks para atualizar os gráficos e a tabela com base na seleção do mês
@app.callback(
    [Output('combined-chart', 'figure'),
     Output('pie-chart', 'figure'),
     Output('table', 'data')],
    [Input('month-dropdown', 'value')]
)
def update_components(selected_month):
    filtered_df = contagem[contagem['Month'] == selected_month]

    # Gráfico de barras
    fig = px.bar(filtered_df, x='Nature_new', y='Quantidade', title='Quantidade por Tipo',
                 labels={'Nature_new': 'Tipo', 'Quantidade': 'Quantidade'}, color='Nature_new',
                 hover_data={'Division': True, 'Header Text': True})

    # Gráfico de pizza
    pie_fig = px.pie(filtered_df, values='Quantidade', names='Nature_new', title='Distribuição por Tipo')

    # Dados da tabela
    table_data = filtered_df.to_dict('records')

    return fig, pie_fig, table_data


# Executar o aplicativo
if __name__ == '__main__':
    app.run_server(debug=True)
