from dash import Dash, html, dcc, Input, Output, dash_table
import pandas as pd
import plotly.express as px
import os
import unicodedata
from datetime import datetime

print("🚀 Iniciando Dashboard de Auditoria...")

# ---------- Helpers ----------
def normalize_colname(name: str) -> str:
    if not isinstance(name, str):
        return name
    nfkd = unicodedata.normalize('NFKD', name)
    only_ascii = ''.join([c for c in nfkd if not unicodedata.combining(c)])
    clean = only_ascii.replace(' ', '').replace('-', '').replace('.', '').strip()
    return clean

def normalize_df_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    mapping = {col: normalize_colname(col) for col in df.columns}
    df.rename(columns=mapping, inplace=True)
    return df

def canonical_status(s):
    if pd.isna(s):
        return "Não Iniciado"
    
    s = str(s).strip()
    s_lower = s.lower()
    
    # Mapeamento direto para os status do checklist
    if s_lower in ['conforme', 'c']:
        return "Conforme"
    elif s_lower in ['conforme parcialmente', 'parcialmente', 'parcial', 'conforme parcial']:
        return "Conforme Parcialmente"
    elif s_lower in ['não conforme', 'nao conforme', 'não', 'nao', 'nc', 'não conf', 'nao conf']:
        return "Não Conforme"
    elif s_lower in ['finalizado', 'finalizada', 'fim']:
        return "Finalizado"
    elif s_lower in ['pendente', 'pendencia', 'pend']:
        return "Pendente"
    elif s_lower in ['não iniciado', 'nao iniciado', '']:
        return "Não Iniciado"
    else:
        # Para qualquer outra coisa, capitaliza a primeira letra
        return s.title()

def get_status_color(status):
    """Retorna cor baseada no status"""
    status_str = str(status).strip().lower()
    
    if 'não iniciado' in status_str or 'nao iniciado' in status_str:
        return {
            'bg_color': '#fdecea',
            'text_color': '#c0392b',
            'border_color': '#c0392b'
        }
    elif 'pendente' in status_str:
        return {
            'bg_color': '#fff8e1',
            'text_color': '#f39c12',
            'border_color': '#f39c12'
        }
    elif 'finalizado' in status_str:
        return {
            'bg_color': '#eafaf1',
            'text_color': '#27ae60',
            'border_color': '#27ae60'
        }
    else:
        return {
            'bg_color': '#f8f9fa',
            'text_color': '#2c3e50',
            'border_color': '#95a5a6'
        }

def calcular_status_prazo(prazo_str, finalizacao_str):
    """Calcula o status baseado no prazo e data de finalização"""
    try:
        # Se não tem data de finalização
        if pd.isna(finalizacao_str) or str(finalizacao_str).strip() in ['', 'NaT', 'None']:
            return "Não Concluído"
        
        # Converter strings para datetime
        prazo = pd.to_datetime(prazo_str, errors='coerce')
        finalizacao = pd.to_datetime(finalizacao_str, errors='coerce')
        
        if pd.isna(prazo) or pd.isna(finalizacao):
            return "Não Concluído"
        
        # Comparar datas
        if finalizacao <= prazo:
            return "Concluído no Prazo"
        else:
            return "Concluído Fora do Prazo"
            
    except:
        return "Não Concluído"

def formatar_data(data_str):
    """Formata a data para DD/MM/YYYY"""
    try:
        if pd.isna(data_str) or str(data_str).strip() in ['', 'NaT', 'None']:
            return ""
        data = pd.to_datetime(data_str, errors='coerce')
        if pd.isna(data):
            return ""
        return data.strftime('%d/%m/%Y')
    except:
        return str(data_str)

def carregar_dados_da_planilha():
planilha_path = 'base_auditoria.xlsx'
    if not os.path.exists(planilha_path):
        print("❌ Planilha não encontrada: data/base_auditoria.xlsx")
        return None, None, None, None

    try:
        print(f"📁 Carregando dados da planilha: {planilha_path}")

        df_checklist = pd.read_excel(planilha_path, sheet_name='Checklist_Unidades')
        df_politicas = pd.read_excel(planilha_path, sheet_name='Politicas')
        df_risco = pd.read_excel(planilha_path, sheet_name='Auditoria_Risco')
        df_melhorias = pd.read_excel(planilha_path, sheet_name='Melhorias_Logistica')

        for i, df in enumerate([df_checklist, df_politicas, df_risco, df_melhorias]):
            if df is not None:
                df = normalize_df_columns(df)
                if 'Status' in df.columns:
                    print(f"DEBUG: Valores únicos de Status antes da normalização: {df['Status'].unique()}")
                    df['Status'] = df['Status'].apply(canonical_status)
                    print(f"DEBUG: Valores únicos de Status após normalização: {df['Status'].unique()}")
                # Normalizar datas para checklist e melhorias, se houver
                if 'Data' in df.columns:
                    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
                    df['Ano'] = df['Data'].dt.year
                    df['Mes'] = df['Data'].dt.month
                    df['Mes_Ano'] = df['Data'].dt.strftime('%m/%Y')
                    try:
                        df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')
                    except:
                        pass
                if i == 0: df_checklist = df
                elif i == 1: df_politicas = df
                elif i == 2: df_risco = df
                elif i == 3: df_melhorias = df

        print("✅ Dados carregados da planilha com sucesso!")
        return df_checklist, df_politicas, df_risco, df_melhorias

    except Exception as e:
        print(f"❌ Erro ao carregar planilha: {e}")
        return None, None, None, None

# ---------- Funções utilitárias ----------
def obter_anos_disponiveis(df_checklist):
    if df_checklist is None or 'Ano' not in df_checklist.columns:
        return []
    return sorted(df_checklist['Ano'].dropna().unique(), reverse=True)

def obter_meses_disponiveis(df_checklist, ano_selecionado):
    if df_checklist is None or 'Mes' not in df_checklist.columns:
        return []
    if ano_selecionado == 'todos':
        meses = sorted(df_checklist['Mes'].dropna().unique())
    else:
        ano_selecionado = int(ano_selecionado)
        df_filtrado = df_checklist[df_checklist['Ano'] == ano_selecionado]
        meses = sorted(df_filtrado['Mes'].dropna().unique())
    nomes_meses = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
                   7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}
    return [{'label': f'{nomes_meses.get(int(m), str(m))}', 'value': int(m)} for m in meses]

# ---------- Carregar dados ----------
df_checklist, df_politicas, df_risco, df_melhorias = carregar_dados_da_planilha()
if df_checklist is None:
    app = Dash(__name__)
    server = app.server
    app.layout = html.Div([html.H1("❌ Planilha não encontrada")])
    if __name__ == '__main__':
        app.run(debug=True, port=8050)
    exit()

anos_disponiveis = obter_anos_disponiveis(df_checklist)

# ---------- App ----------
app = Dash(__name__)

app.layout = html.Div([
    html.Div([html.H1("📊 DASHBOARD DE AUDITORIA", style={'textAlign':'center', 'marginBottom':'20px'})]),
    html.Div([
        html.Div([
            html.Label("Ano:"),
            dcc.Dropdown(
                id='filtro-ano',
                options=[{'label':'Todos','value':'todos'}]+[{'label':str(a),'value':a} for a in anos_disponiveis],
                value='todos'
            )
        ], style={'marginRight':'20px','width':'200px'}),
        html.Div([
            html.Label("Mês:"),
            dcc.Dropdown(
                id='filtro-mes',
                options=[{'label':'Todos','value':'todos'}],
                value='todos'
            )
        ], style={'marginRight':'20px','width':'200px'}),
        html.Div([
            html.Label("Unidade:"),
            dcc.Dropdown(
                id='filtro-unidade',
                options=[{'label':'Todas','value':'todas'}]+
                        [{'label':u,'value':u} for u in sorted(df_checklist['Unidade'].unique())],
                value='todas'
            )
        ], style={'width':'250px'})
    ], style={'display':'flex','justifyContent':'center','marginBottom':'30px','flexWrap':'wrap'}),
    html.Div(id='conteudo-principal', style={'padding':'20px'})
])

@app.callback(
    Output('filtro-mes','options'),
    Input('filtro-ano','value')
)
def atualizar_opcoes_mes(ano_selecionado):
    return [{'label':'Todos','value':'todos'}]+obter_meses_disponiveis(df_checklist, ano_selecionado)

@app.callback(
    Output('conteudo-principal','children'),
    [Input('filtro-ano','value'),
     Input('filtro-mes','value'),
     Input('filtro-unidade','value')]
)
def atualizar_conteudo_principal(ano, mes, unidade):
    # ---------- FILTRAR CHECKLIST ----------
    df = df_checklist.copy()
    if ano != 'todos':
        df = df[df['Ano']==int(ano)]
    if mes != 'todos':
        df = df[df['Mes']==int(mes)]
    if unidade != 'todas':
        df = df[df['Unidade']==unidade]

    total = len(df)

    # ---------- Contagem correta dos status ----------
    status_series = df['Status']
    
    conforme = status_series.value_counts().get('Conforme', 0)
    parcial = status_series.value_counts().get('Conforme Parcialmente', 0)
    nao = status_series.value_counts().get('Não Conforme', 0)
    
    # Verificação de debug
    if conforme + parcial + nao != total:
        print(f"⚠️ Aviso: Contagem de status não bate com total!")
        conforme = len(df[df['Status'] == 'Conforme'])
        parcial = len(df[df['Status'] == 'Conforme Parcialmente'])
        nao = len(df[df['Status'] == 'Não Conforme'])

    # ---------- KPIs ----------
    kpis = html.Div([
        html.Div([
            html.H4("Conforme", style={'color':'#27ae60','margin':'0'}),
            html.H2(f"{conforme}", style={'color':'#27ae60','margin':'0'}),
            html.P(f"{(conforme/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#27ae60'})
        ], style={'borderLeft':'5px solid #27ae60','borderRadius':'5px','padding':'20px','margin':'10px','flex':'1',
                  'backgroundColor':'#eafaf1','textAlign':'center','boxShadow':'2px 2px 5px rgba(0,0,0,0.1)'}),

        html.Div([
            html.H4("Conforme Parcialmente", style={'color':'#f39c12','margin':'0'}),
            html.H2(f"{parcial}", style={'color':'#f39c12','margin':'0'}),
            html.P(f"{(parcial/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#f39c12'})
        ], style={'borderLeft':'5px solid #f39c12','borderRadius':'5px','padding':'20px','margin':'10px','flex':'1',
                  'backgroundColor':'#fff8e1','textAlign':'center','boxShadow':'2px 2px 5px rgba(0,0,0,0.1)'}),

        html.Div([
            html.H4("Não Conforme", style={'color':'#e74c3c','margin':'0'}),
            html.H2(f"{nao}", style={'color':'#e74c3c','margin':'0'}),
            html.P(f"{(nao/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#e74c3c'})
        ], style={'borderLeft':'5px solid #e74c3c','borderRadius':'5px','padding':'20px','margin':'10px','flex':'1',
                  'backgroundColor':'#fdecea','textAlign':'center','boxShadow':'2px 2px 5px rgba(0,0,0,0.1)'})
    ], style={'display':'flex','justifyContent':'center','flexWrap':'wrap','marginBottom':'30px'})

    # ---------- Gráfico ----------
    fig = px.pie(
        names=['Conforme','Conforme Parcialmente','Não Conforme'],
        values=[conforme,parcial,nao],
        color=['Conforme','Conforme Parcialmente','Não Conforme'],
        color_discrete_map={'Conforme':'#27ae60','Conforme Parcialmente':'#f39c12','Não Conforme':'#e74c3c'},
        title="Distribuição de Status - Checklist"
    )

    # ---------- Tabela de NÃO CONFORMES ----------
    df_nao_conforme = df[df['Status']=='Não Conforme']
    
    if len(df_nao_conforme) > 0:
        # Verificar se as colunas de prazo e data de finalização existem
        colunas_disponiveis = df_nao_conforme.columns.tolist()
        
        # Normalizar nomes de colunas para busca case-insensitive
        colunas_lower = [str(col).lower() for col in colunas_disponiveis]
        
        # Procurar colunas de prazo e data de finalização
        coluna_prazo = None
        coluna_finalizacao = None
        
        # Possíveis nomes para coluna de prazo
        possiveis_prazos = ['prazo', 'data_limite', 'data_prazo', 'prazo_final', 'limite']
        for prazo_nome in possiveis_prazos:
            for idx, col_lower in enumerate(colunas_lower):
                if prazo_nome in col_lower:
                    coluna_prazo = colunas_disponiveis[idx]
                    break
            if coluna_prazo:
                break
        
        # Possíveis nomes para coluna de data de finalização
        possiveis_finalizacoes = ['data_finalizacao', 'data_conclusao', 'finalizacao', 'conclusao', 
                                  'data_encerramento', 'data_termino', 'finalizado_em']
        for final_nome in possiveis_finalizacoes:
            for idx, col_lower in enumerate(colunas_lower):
                if final_nome in col_lower:
                    coluna_finalizacao = colunas_disponiveis[idx]
                    break
            if coluna_finalizacao:
                break
        
        # Criar cópia do DataFrame para modificar
        df_nao_conforme_display = df_nao_conforme.copy()
        
        # Remover colunas que não queremos mostrar
        colunas_para_remover = ['Ano', 'Mes', 'Mes_Ano']
        for col in colunas_para_remover:
            if col in df_nao_conforme_display.columns:
                df_nao_conforme_display = df_nao_conforme_display.drop(columns=[col])
        
        # Verificar se temos colunas de prazo e finalização
        tem_prazo = coluna_prazo is not None and coluna_prazo in df_nao_conforme_display.columns
        tem_finalizacao = coluna_finalizacao is not None and coluna_finalizacao in df_nao_conforme_display.columns
        
        # Se temos ambas as colunas, calcular status do prazo
        if tem_prazo and tem_finalizacao:
            # Garantir que as colunas estejam formatadas
            df_nao_conforme_display['Prazo_Formatado'] = df_nao_conforme_display[coluna_prazo].apply(formatar_data)
            df_nao_conforme_display['Finalizacao_Formatada'] = df_nao_conforme_display[coluna_finalizacao].apply(formatar_data)
            
            # Calcular status do prazo
            df_nao_conforme_display['Status_Prazo'] = df_nao_conforme_display.apply(
                lambda row: calcular_status_prazo(row[coluna_prazo], row[coluna_finalizacao]), axis=1
            )
            
            # Reordenar colunas para mostrar as importantes primeiro
            colunas_ordenadas = []
            colunas_restantes = []
            
            # Priorizar certas colunas
            colunas_prioridade = ['Unidade', 'Status', 'Status_Prazo', 'Prazo_Formatado', 'Finalizacao_Formatada']
            
            for col in colunas_prioridade:
                if col in df_nao_conforme_display.columns:
                    colunas_ordenadas.append(col)
            
            # Adicionar as outras colunas (exceto as que não queremos)
            for col in df_nao_conforme_display.columns:
                if col not in colunas_ordenadas and col not in [coluna_prazo, coluna_finalizacao, 'Prazo_Formatado', 'Finalizacao_Formatada']:
                    colunas_restantes.append(col)
            
            colunas_finais = colunas_ordenadas + colunas_restantes
            df_nao_conforme_display = df_nao_conforme_display[colunas_finais]
            
            # Renomear colunas para exibição
            rename_dict = {
                'Prazo_Formatado': 'Prazo',
                'Finalizacao_Formatada': 'Data Finalização',
                'Status_Prazo': 'Status Prazo'
            }
            df_nao_conforme_display = df_nao_conforme_display.rename(columns=rename_dict)
            
            # Condições de estilo baseadas no status do prazo
            style_data_conditional = [
                # Status do prazo - Cores da linha inteira
                {
                    'if': {
                        'filter_query': '{Status Prazo} = "Concluído no Prazo"'
                    },
                    'backgroundColor': '#eafaf1'
                },
                {
                    'if': {
                        'filter_query': '{Status Prazo} = "Concluído Fora do Prazo"'
                    },
                    'backgroundColor': '#fff8e1'
                },
                {
                    'if': {
                        'filter_query': '{Status Prazo} = "Não Concluído"'
                    },
                    'backgroundColor': '#fdecea'
                },
                
                # Status do prazo - Destaque na coluna
                {
                    'if': {
                        'filter_query': '{Status Prazo} = "Concluído no Prazo"',
                        'column_id': 'Status Prazo'
                    },
                    'backgroundColor': '#27ae60',
                    'color': 'white',
                    'fontWeight': 'bold'
                },
                {
                    'if': {
                        'filter_query': '{Status Prazo} = "Concluído Fora do Prazo"',
                        'column_id': 'Status Prazo'
                    },
                    'backgroundColor': '#f39c12',
                    'color': 'white',
                    'fontWeight': 'bold'
                },
                {
                    'if': {
                        'filter_query': '{Status Prazo} = "Não Concluído"',
                        'column_id': 'Status Prazo'
                    },
                    'backgroundColor': '#e74c3c',
                    'color': 'white',
                    'fontWeight': 'bold'
                },
                
                # Status original (Não Conforme)
                {
                    'if': {
                        'filter_query': '{Status} = "Não Conforme"',
                        'column_id': 'Status'
                    },
                    'backgroundColor': '#c0392b',
                    'color': 'white',
                    'fontWeight': 'bold'
                }
            ]
        else:
            # Se não tem colunas de prazo, usar estilo básico
            style_data_conditional = [
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9e6e6'},
                {'if': {'row_index': 'even'}, 'backgroundColor': '#fdecea'},
                {
                    'if': {
                        'filter_query': '{Status} = "Não Conforme"',
                        'column_id': 'Status'
                    },
                    'backgroundColor': '#c0392b',
                    'color': 'white',
                    'fontWeight': 'bold'
                }
            ]
        
        tabela_nao_conforme = dash_table.DataTable(
            df_nao_conforme_display.to_dict('records'),
            columns=[{"name": col, "id": col} for col in df_nao_conforme_display.columns],
            page_size=10,
            style_table={'overflowX':'auto'},
            style_header={'backgroundColor': '#c0392b','color': 'white','fontWeight': 'bold','textAlign':'center'},
            style_cell={'textAlign': 'center','padding': '5px','whiteSpace':'normal','height':'auto'},
            style_data_conditional=style_data_conditional
        )
        tabela_titulo = html.H3(f"❌ Itens Não Conformes ({len(df_nao_conforme)} itens)")
    else:
        tabela_nao_conforme = html.Div([
            html.P("✅ Nenhum item não conforme encontrado com os filtros atuais.", 
                   style={'textAlign': 'center', 'padding': '20px', 'color': '#27ae60'})
        ])
        tabela_titulo = html.H3("❌ Itens Não Conformes (0 itens)")

    # ---------- Matriz de Risco (COM FILTROS) ----------
    abas_extra = []

    if df_risco is not None and len(df_risco) > 0:
        df_risco_filtrado = df_risco.copy()
        if ano != 'todos' and 'Ano' in df_risco_filtrado.columns:
            df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Ano'] == int(ano)]
        if mes != 'todos' and 'Mes' in df_risco_filtrado.columns:
            df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Mes'] == int(mes)]
        if unidade != 'todas' and 'Unidade' in df_risco_filtrado.columns:
            df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade]

        # Criar Mes_Ano com base em Mes e Ano
        if 'Mes' in df_risco_filtrado.columns and 'Ano' in df_risco_filtrado.columns:
            df_risco_filtrado['Mes_Ano'] = df_risco_filtrado['Mes'].astype(str).str.zfill(2) + '/' + df_risco_filtrado['Ano'].astype(str)
        else:
            df_risco_filtrado['Mes_Ano'] = ""

        if len(df_risco_filtrado) > 0:
            # Agrupar dados por Unidade e Mes_Ano
            unidades = sorted(df_risco_filtrado['Unidade'].unique())
            meses_anos = sorted(df_risco_filtrado['Mes_Ano'].unique())
            
            # Criar estrutura de dados para a matriz
            matriz_data = []
            
            for unidade_nome in unidades:
                linha = {'Unidade': unidade_nome}
                df_unidade = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade_nome]
                
                for mes_ano in meses_anos:
                    if pd.isna(mes_ano):
                        continue
                    
                    df_mes = df_unidade[df_unidade['Mes_Ano'] == mes_ano]
                    
                    if len(df_mes) > 0:
                        # Criar lista de relatórios com cores
                        relatorios_html = []
                        for _, row in df_mes.iterrows():
                            relatorio = row['Relatorio']
                            status = row['Status']
                            cores = get_status_color(status)
                            
                            relatorio_item = html.Span(
                                str(relatorio),
                                style={
                                    'display': 'inline-block',
                                    'backgroundColor': cores['bg_color'],
                                    'color': cores['text_color'],
                                    'padding': '4px 8px',
                                    'margin': '2px',
                                    'borderRadius': '3px',
                                    'fontSize': '12px',
                                    'fontWeight': 'bold',
                                    'borderLeft': f'3px solid {cores["border_color"]}'
                                }
                            )
                            relatorios_html.append(relatorio_item)
                        
                        # Se houver múltiplos relatórios, colocar em container
                        if len(relatorios_html) > 0:
                            linha[mes_ano] = html.Div(
                                relatorios_html,
                                style={
                                    'display': 'flex',
                                    'flexWrap': 'wrap',
                                    'gap': '3px',
                                    'justifyContent': 'center',
                                    'alignItems': 'center',
                                    'minHeight': '40px'
                                }
                            )
                        else:
                            linha[mes_ano] = ""
                    else:
                        linha[mes_ano] = ""
                
                matriz_data.append(linha)
            
            # Criar tabela HTML manualmente
            tabela_cabecalho = [html.Th("Unidade", style={
                'backgroundColor': '#34495e',
                'color': 'white',
                'padding': '12px',
                'textAlign': 'center',
                'fontWeight': 'bold',
                'border': '1px solid #2c3e50',
                'minWidth': '150px',
                'position': 'sticky',
                'left': '0',
                'zIndex': '1'
            })]
            
            for mes_ano in meses_anos:
                if pd.isna(mes_ano):
                    continue
                tabela_cabecalho.append(html.Th(str(mes_ano), style={
                    'backgroundColor': '#34495e',
                    'color': 'white',
                    'padding': '12px',
                    'textAlign': 'center',
                    'fontWeight': 'bold',
                    'border': '1px solid #2c3e50',
                    'minWidth': '200px'
                }))
            
            tabela_linhas = []
            
            for i, linha in enumerate(matriz_data):
                bg_color = '#f8f9fa' if i % 2 == 0 else 'white'
                
                celulas = [html.Td(linha['Unidade'], style={
                    'backgroundColor': bg_color,
                    'padding': '10px',
                    'textAlign': 'center',
                    'border': '1px solid #dee2e6',
                    'fontWeight': 'bold',
                    'position': 'sticky',
                    'left': '0',
                    'zIndex': '1'
                })]
                
                for mes_ano in meses_anos:
                    if pd.isna(mes_ano):
                        continue
                    
                    conteudo = linha.get(mes_ano, "")
                    celulas.append(html.Td(
                        conteudo,
                        style={
                            'backgroundColor': bg_color,
                            'padding': '8px',
                            'textAlign': 'center',
                            'border': '1px solid #dee2e6',
                            'verticalAlign': 'middle',
                            'minHeight': '50px'
                        }
                    ))
                
                tabela_linhas.append(html.Tr(celulas, style={'borderBottom': '1px solid #dee2e6'}))
            
            # Criar tabela HTML
            tabela_html = html.Table([
                html.Thead(html.Tr(tabela_cabecalho)),
                html.Tbody(tabela_linhas)
            ], style={
                'width': '100%',
                'borderCollapse': 'collapse',
                'marginTop': '10px',
                'fontFamily': 'Arial, sans-serif',
                'fontSize': '14px'
            })
            
            # Container com scroll horizontal
            tabela_container = html.Div(
                tabela_html,
                style={
                    'overflowX': 'auto',
                    'maxWidth': '100%',
                    'marginTop': '15px',
                    'border': '1px solid #dee2e6',
                    'borderRadius': '5px',
                    'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'
                }
            )
            
            # Legenda
            legenda = html.Div([
                html.H4("Legenda de Status:", style={'marginBottom': '10px', 'color': '#2c3e50'}),
                html.Div([
                    html.Div([
                        html.Span("🔴 ", style={'fontSize': '16px', 'marginRight': '5px'}),
                        html.Span("Não Iniciado", style={'color': '#2c3e50'})
                    ], style={'display': 'inline-block', 'marginRight': '20px'}),
                    
                    html.Div([
                        html.Span("🟡 ", style={'fontSize': '16px', 'marginRight': '5px'}),
                        html.Span("Pendente", style={'color': '#2c3e50'})
                    ], style={'display': 'inline-block', 'marginRight': '20px'}),
                    
                    html.Div([
                        html.Span("🟢 ", style={'fontSize': '16px', 'marginRight': '5px'}),
                        html.Span("Finalizado", style={'color': '#2c3e50'})
                    ], style={'display': 'inline-block'})
                ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '15px'})
            ], style={
                'backgroundColor': '#f8f9fa',
                'padding': '15px',
                'borderRadius': '5px',
                'marginBottom': '20px',
                'border': '1px solid #dee2e6'
            })
            
            abas_extra.append(html.Div([
                html.H3("📋 Matriz Auditoria Risco", style={'marginBottom': '20px'}),
                html.P(f"Total de registros: {len(df_risco_filtrado)}", 
                       style={'color': '#7f8c8d', 'marginBottom': '10px'}),
                legenda,
                tabela_container
            ], style={
                'marginTop': '30px',
                'padding': '25px',
                'backgroundColor': 'white',
                'borderRadius': '8px',
                'boxShadow': '0 2px 8px rgba(0,0,0,0.1)'
            }))
        else:
            abas_extra.append(html.Div([
                html.H3("📋 Matriz Auditoria Risco"),
                html.P("Nenhum dado encontrado com os filtros atuais.", 
                       style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '40px'})
            ], style={'marginTop':'30px'}))
    else:
        abas_extra.append(html.Div([
            html.H3("📋 Matriz Auditoria Risco"),
            html.P("Não há dados de risco disponíveis.", 
                   style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '40px'})
        ], style={'marginTop':'30px'}))

    # ---------- Melhorias e Políticas (SEM FILTROS) ----------
    # Mostrar sempre todos os dados sem aplicar filtros
    if df_melhorias is not None and len(df_melhorias) > 0:
        tabela_melhorias = dash_table.DataTable(
            df_melhorias.to_dict('records'),
            page_size=10,
            style_table={'overflowX':'auto','marginTop':'10px'},
            style_header={'backgroundColor': '#34495e','color': 'white','fontWeight': 'bold','textAlign':'center'},
            style_cell={'textAlign': 'center','padding': '5px','whiteSpace':'normal','height':'auto'},
            style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': '#ecf0f1'},
                                    {'if': {'row_index': 'even'}, 'backgroundColor': 'white'}]
        )
        abas_extra.append(html.Div([
            html.H3(f"📈 Melhorias ({len(df_melhorias)} registros)"),
            html.P("Todos os registros de melhorias (filtros não aplicados)", 
                   style={'color': '#7f8c8d', 'marginBottom': '10px'}),
            tabela_melhorias
        ], style={'marginTop':'30px'}))
    else:
        abas_extra.append(html.Div([
            html.H3("📈 Melhorias"),
            html.P("Não há dados de melhorias disponíveis.", 
                   style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '20px'})
        ], style={'marginTop':'30px'}))

    if df_politicas is not None and len(df_politicas) > 0:
        tabela_politicas = dash_table.DataTable(
            df_politicas.to_dict('records'),
            page_size=10,
            style_table={'overflowX':'auto','marginTop':'10px'},
            style_header={'backgroundColor': '#34495e','color': 'white','fontWeight': 'bold','textAlign':'center'},
            style_cell={'textAlign': 'center','padding': '5px','whiteSpace':'normal','height':'auto'},
            style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': '#ecf0f1'},
                                    {'if': {'row_index': 'even'}, 'backgroundColor': 'white'}]
        )
        abas_extra.append(html.Div([
            html.H3(f"📑 Políticas ({len(df_politicas)} registros)"),
            html.P("Todos os registros de políticas (filtros não aplicados)", 
                   style={'color': '#7f8c8d', 'marginBottom': '10px'}),
            tabela_politicas
        ], style={'marginTop':'30px'}))
    else:
        abas_extra.append(html.Div([
            html.H3("📑 Políticas"),
            html.P("Não há dados de políticas disponíveis.", 
                   style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '20px'})
        ], style={'marginTop':'30px'}))

    return html.Div([
        html.Div([
            html.H4(f"📊 Resumo - {len(df)} itens auditados", 
                    style={'textAlign':'center', 'color':'#2c3e50', 'marginBottom':'20px'})
        ]),
        kpis,
        dcc.Graph(figure=fig),
        tabela_titulo,
        tabela_nao_conforme,
        *abas_extra
    ])

if __name__ == '__main__':
    print("🌐 DASHBOARD RODANDO: http://localhost:8050")
    app.run(debug=True, host='0.0.0.0', port=8050)

