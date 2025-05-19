import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px
import os

# Caminho relativo para os arquivos na pasta 'data'
data_dir = "data"
df_2020 = pd.read_excel(os.path.join(data_dir, "Base Vendas - 2020.xlsx"))
df_2021 = pd.read_excel(os.path.join(data_dir, "Base Vendas - 2021.xlsx"))
df_2022 = pd.read_excel(os.path.join(data_dir, "Base Vendas - 2022.xlsx"))
df_clientes = pd.read_excel(os.path.join(data_dir, "Cadastro Clientes.xlsx"), header=2)
df_lojas = pd.read_excel(os.path.join(data_dir, "Cadastro Lojas.xlsx"))
df_produtos = pd.read_excel(os.path.join(data_dir, "Cadastro Produtos.xlsx"))

# Unificar dados de vendas
df = pd.concat([df_2020, df_2021, df_2022], ignore_index=True)

# Corrigir espaços nos nomes das colunas
df.columns = df.columns.str.strip()
df_clientes.columns = df_clientes.columns.str.strip()
df_lojas.columns = df_lojas.columns.str.strip()
df_produtos.columns = df_produtos.columns.str.strip()

# Nome completo do cliente
df_clientes['Cliente'] = df_clientes['Primeiro Nome'] + ' ' + df_clientes['Sobrenome']

# Mesclar cadastros
df = df.merge(df_clientes[['ID Cliente', 'Cliente']], on='ID Cliente', how='left')
df = df.merge(df_lojas, on='ID Loja', how='left')
df = df.merge(df_produtos, on='SKU', how='left')

# Corrigir tipos
df['Data da Venda'] = pd.to_datetime(df['Data da Venda'])
df['Ano'] = df['Data da Venda'].dt.year
df['Valor_Venda'] = df['Qtd Vendida'] * df['Preço Unitario']

# Inicializar o app
app = dash.Dash(__name__)
server = app.server  # Necessário para a Render

# Layout
app.layout = html.Div([
    html.H1("Dashboard de Vendas", style={'textAlign': 'center'}),

    html.Div([
        html.Label('Tipo de Produto'),
        dcc.Dropdown(
            id='tipo-produto-filter',
            options=[{'label': i, 'value': i} for i in sorted(df['Tipo do Produto'].dropna().unique())],
            placeholder='Selecione um tipo de produto'
        ),

        html.Label('Marca'),
        dcc.Dropdown(id='marca-filter', placeholder='Selecione uma ou mais marcas', multi=True),

        html.Label('Produto'),
        dcc.Dropdown(id='produto-filter', placeholder='Selecione um produto'),

        html.Label('Loja'),
        dcc.Dropdown(id='loja-filter', placeholder='Selecione uma ou mais lojas', multi=True),

        html.Label('Cliente'),
        dcc.Dropdown(id='cliente-filter', placeholder='Selecione um cliente'),
    ], style={'columnCount': 2, 'padding': '20px'}),

    html.Div([
        dcc.Graph(id='grafico-ano'),
        dcc.Graph(id='grafico-cliente'),
        dcc.Graph(id='grafico-produto'),
        dcc.Graph(id='grafico-loja'),
        dcc.Graph(id='grafico-area'),
        dcc.Graph(id='grafico-pizza'),
    ])
])

# Callbacks
@app.callback(
    Output('marca-filter', 'options'),
    Input('tipo-produto-filter', 'value')
)
def update_marcas(tipo):
    dff = df[df['Tipo do Produto'] == tipo] if tipo else df
    return [{'label': m, 'value': m} for m in sorted(dff['Marca'].dropna().unique())]

@app.callback(
    Output('produto-filter', 'options'),
    [Input('tipo-produto-filter', 'value'),
     Input('marca-filter', 'value')]
)
def update_produtos(tipo, marcas):
    dff = df.copy()
    if tipo:
        dff = dff[dff['Tipo do Produto'] == tipo]
    if marcas:
        dff = dff[dff['Marca'].isin(marcas)]
    return [{'label': p, 'value': p} for p in sorted(dff['Produto'].dropna().unique())]

@app.callback(
    Output('loja-filter', 'options'),
    [Input('tipo-produto-filter', 'value'),
     Input('marca-filter', 'value'),
     Input('produto-filter', 'value')]
)
def update_lojas(tipo, marcas, produto):
    dff = df.copy()
    if tipo:
        dff = dff[dff['Tipo do Produto'] == tipo]
    if marcas:
        dff = dff[dff['Marca'].isin(marcas)]
    if produto:
        dff = dff[dff['Produto'] == produto]
    return [{'label': l, 'value': l} for l in sorted(dff['Nome da Loja'].dropna().unique())]

@app.callback(
    Output('cliente-filter', 'options'),
    [Input('tipo-produto-filter', 'value'),
     Input('marca-filter', 'value'),
     Input('produto-filter', 'value'),
     Input('loja-filter', 'value')]
)
def update_clientes(tipo, marcas, produto, lojas):
    dff = df.copy()
    if tipo:
        dff = dff[dff['Tipo do Produto'] == tipo]
    if marcas:
        dff = dff[dff['Marca'].isin(marcas)]
    if produto:
        dff = dff[dff['Produto'] == produto]
    if lojas:
        dff = dff[dff['Nome da Loja'].isin(lojas)]
    return [{'label': c, 'value': c} for c in sorted(dff['Cliente'].dropna().unique())]

@app.callback(
    [Output('grafico-ano', 'figure'),
     Output('grafico-cliente', 'figure'),
     Output('grafico-produto', 'figure'),
     Output('grafico-loja', 'figure'),
     Output('grafico-area', 'figure'),
     Output('grafico-pizza', 'figure')],
    [Input('tipo-produto-filter', 'value'),
     Input('marca-filter', 'value'),
     Input('produto-filter', 'value'),
     Input('loja-filter', 'value'),
     Input('cliente-filter', 'value')]
)
def update_graficos(tipo, marcas, produto, lojas, cliente):
    dff = df.copy()
    if tipo:
        dff = dff[dff['Tipo do Produto'] == tipo]
    if marcas:
        dff = dff[dff['Marca'].isin(marcas)]
    if produto:
        dff = dff[dff['Produto'] == produto]
    if lojas:
        dff = dff[dff['Nome da Loja'].isin(lojas)]
    if cliente:
        dff = dff[dff['Cliente'] == cliente]

    if dff.empty:
        fig = px.bar(title="Nenhum dado encontrado")
        return [fig] * 6

    fig_ano = px.bar(dff, x='Ano', y='Valor_Venda', title='Vendas por Ano')
    fig_cliente = px.bar(dff, x='Cliente', y='Valor_Venda', title='Vendas por Cliente')
    fig_produto = px.bar(dff, x='Valor_Venda', y='Produto', title='Vendas por Produto', orientation='h')
    fig_loja = px.bar(
        dff.groupby('Nome da Loja', as_index=False)['Valor_Venda'].sum().sort_values('Valor_Venda'),
        x='Valor_Venda', y='Nome da Loja',
        title='Vendas por Loja', orientation='h'
    )
    fig_area = px.area(dff, x='Data da Venda', y='Valor_Venda', title='Vendas ao Longo do Tempo')
    fig_pizza = px.pie(dff, names='Marca', values='Valor_Venda', title='Vendas por Marca')

    return fig_ano, fig_cliente, fig_produto, fig_loja, fig_area, fig_pizza

# Rodar app
if __name__ == '__main__':
    app.run_server(debug=False, host='0.0.0.0', port=int(os.environ.get("PORT", 8050)))
