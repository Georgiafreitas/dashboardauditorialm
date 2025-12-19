def criar_matriz_risco_anual(df_risco_filtrado, ano_filtro):
    """Cria matriz de risco com todas as siglas visíveis e lista completa embaixo"""
    
    # Verificar se temos dados
    if df_risco_filtrado is None or len(df_risco_filtrado) == 0:
        return html.Div([
            html.H3("📋 Matriz Auditoria Risco"),
            html.P("Nenhum dado encontrado para o ano selecionado.", 
                   style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '20px'})
        ], style={'marginTop':'30px'})
    
    # Obter unidades únicas
    unidades = sorted(df_risco_filtrado['Unidade'].dropna().unique())
    
    # Lista de meses do ano (1 a 12)
    meses_ano = list(range(1, 13))
    nomes_meses = {
        1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN',
        7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
    }
    nomes_completos = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    
    print(f"\n📊 CRIANDO MATRIZ DE RISCO PARA O ANO {ano_filtro}")
    print(f"  Unidades: {unidades}")
    print(f"  Total de registros: {len(df_risco_filtrado)}")
    
    # Criar dicionário para mapear siglas para relatórios completos
    mapeamento_siglas = {}
    siglas_por_unidade_mes = {}  # Para armazenar todas as siglas por unidade/mês
    
    # Criar estrutura de dados para a matriz
    matriz_data = []
    
    for unidade_nome in unidades:
        linha = {'Unidade': unidade_nome}
        df_unidade = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade_nome]
        
        for mes in meses_ano:
            # Filtrar por mês (converter para o tipo correto)
            df_unidade['Mes'] = pd.to_numeric(df_unidade['Mes'], errors='coerce')
            df_mes = df_unidade[df_unidade['Mes'] == mes]
            
            if len(df_mes) > 0:
                # Para cada relatório no mês
                relatorios_info = []
                siglas_no_mes = []
                
                for _, row in df_mes.iterrows():
                    relatorio = str(row.get('Relatorio', ''))
                    sigla = str(row.get('Sigla', 'REL'))
                    status = str(row.get('Status', 'Sem Status'))
                    cores = get_status_color(status)
                    
                    # Armazenar mapeamento sigla -> relatório completo
                    mapeamento_siglas[sigla] = {
                        'relatorio': relatorio,
                        'status': status,
                        'unidade': unidade_nome,
                        'mes': mes
                    }
                    
                    # Armazenar sigla para este mês
                    siglas_no_mes.append(sigla)
                    
                    # Criar elemento com a SIGLA visível e status
                    relatorio_item = html.Div([
                        html.Div(
                            sigla,
                            style={
                                'fontSize': '9px',
                                'fontWeight': 'bold',
                                'color': cores['text_color'],
                                'textAlign': 'center',
                                'backgroundColor': cores['bg_color'],
                                'padding': '2px 3px',
                                'margin': '1px',
                                'borderRadius': '2px',
                                'border': f'1px solid {cores["border_color"]}',
                                'minWidth': '40px',
                                'maxWidth': '45px',
                                'height': '22px',
                                'display': 'flex',
                                'alignItems': 'center',
                                'justifyContent': 'center',
                                'cursor': 'default'
                            }
                        )
                    ], style={'display': 'inline-block'})
                    relatorios_info.append(relatorio_item)
                
                # Armazenar siglas para esta unidade/mês
                siglas_por_unidade_mes[f"{unidade_nome}_{mes}"] = siglas_no_mes
                
                if len(relatorios_info) > 0:
                    # Criar container para todas as siglas do mês (verticalmente)
                    if len(relatorios_info) > 4:
                        # Se tiver muitas siglas, mostrar as primeiras 4 e indicador
                        linha[mes] = html.Div([
                            html.Div(relatorios_info[:4], style={'display': 'grid', 'gridTemplateColumns': 'repeat(2, 1fr)', 'gap': '2px'}),
                            html.Div(f"+{len(relatorios_info)-4}", style={
                                'fontSize': '8px',
                                'backgroundColor': '#3498db',
                                'color': 'white',
                                'padding': '1px 3px',
                                'borderRadius': '2px',
                                'marginTop': '2px',
                                'fontWeight': 'bold',
                                'textAlign': 'center'
                            })
                        ], style={
                            'display': 'flex',
                            'flexDirection': 'column',
                            'alignItems': 'center',
                            'justifyContent': 'center',
                            'padding': '3px'
                        })
                    else:
                        # Mostrar todas as siglas
                        if len(relatorios_info) <= 2:
                            linha[mes] = html.Div(relatorios_info, style={
                                'display': 'flex',
                                'flexDirection': 'column',
                                'alignItems': 'center',
                                'justifyContent': 'center',
                                'gap': '2px',
                                'padding': '3px'
                            })
                        else:
                            linha[mes] = html.Div(relatorios_info, style={
                                'display': 'grid',
                                'gridTemplateColumns': 'repeat(2, 1fr)',
                                'gap': '2px',
                                'padding': '3px'
                            })
                else:
                    linha[mes] = html.Div("-", style={'color': '#ccc', 'fontSize': '10px', 'padding': '10px 0'})
            else:
                linha[mes] = html.Div("-", style={'color': '#ccc', 'fontSize': '10px', 'padding': '10px 0'})
        
        matriz_data.append(linha)
    
    # Criar lista de siglas únicas
    siglas_unicas = sorted(set(mapeamento_siglas.keys()))
    
    # Criar tabela HTML com siglas visíveis
    tabela_cabecalho = [html.Th("UNIDADE", style={
        'backgroundColor': '#2c3e50',
        'color': 'white',
        'padding': '8px 10px',
        'textAlign': 'center',
        'fontWeight': 'bold',
        'border': '1px solid #2c3e50',
        'minWidth': '120px',
        'maxWidth': '150px',
        'fontSize': '12px',
        'position': 'sticky',
        'left': '0',
        'zIndex': '2'
    })]
    
    for mes in meses_ano:
        tabela_cabecalho.append(html.Th(
            html.Div([
                html.Div(nomes_meses[mes], style={'fontSize': '11px', 'fontWeight': 'bold'}),
                html.Div(str(mes), style={'fontSize': '9px', 'opacity': '0.8'})
            ]),
            style={
                'backgroundColor': '#2c3e50',
                'color': 'white',
                'padding': '6px 3px',
                'textAlign': 'center',
                'fontWeight': 'bold',
                'border': '1px solid #2c3e50',
                'minWidth': '55px',
                'maxWidth': '65px',
                'fontSize': '11px'
            }
        ))
    
    tabela_linhas = []
    
    for i, linha in enumerate(matriz_data):
        bg_color = '#f8f9fa' if i % 2 == 0 else 'white'
        
        celulas = [html.Td(
            html.Div(linha['Unidade'], style={
                'fontSize': '11px',
                'fontWeight': 'bold',
                'overflow': 'hidden',
                'textOverflow': 'ellipsis',
                'whiteSpace': 'nowrap',
                'padding': '2px 0'
            }),
            style={
                'backgroundColor': bg_color,
                'padding': '8px 10px',
                'textAlign': 'left',
                'border': '1px solid #dee2e6',
                'fontSize': '11px',
                'minWidth': '120px',
                'maxWidth': '150px',
                'position': 'sticky',
                'left': '0',
                'zIndex': '1'
            }
        )]
        
        for mes in meses_ano:
            conteudo = linha.get(mes, html.Div("-", style={'color': '#ccc', 'fontSize': '10px', 'padding': '10px 0'}))
            celulas.append(html.Td(
                conteudo,
                style={
                    'backgroundColor': bg_color,
                    'padding': '4px 2px',
                    'textAlign': 'center',
                    'border': '1px solid #dee2e6',
                    'verticalAlign': 'middle',
                    'minHeight': '60px',
                    'minWidth': '55px',
                    'maxWidth': '65px'
                }
            ))
        
        tabela_linhas.append(html.Tr(celulas, style={'borderBottom': '1px solid #dee2e6'}))
    
    tabela_html = html.Table([
        html.Thead(html.Tr(tabela_cabecalho)),
        html.Tbody(tabela_linhas)
    ], style={
        'width': '100%',
        'borderCollapse': 'collapse',
        'marginTop': '5px',
        'fontFamily': 'Arial, sans-serif',
        'fontSize': '11px',
        'tableLayout': 'fixed'
    })
    
    # Container compacto
    tabela_container = html.Div(
        tabela_html,
        style={
            'overflowX': 'auto',
            'maxWidth': '100%',
            'marginTop': '10px',
            'border': '1px solid #dee2e6',
            'borderRadius': '3px',
            'boxShadow': '0 1px 3px rgba(0,0,0,0.1)',
            'maxHeight': '550px',
            'overflowY': 'auto'
        }
    )
    
    # Criar lista completa de siglas e relatórios (MUITO IMPORTANTE)
    lista_siglas_completa = []
    
    # Agrupar por status primeiro para melhor organização
    siglas_por_status = {
        'Não Iniciado': [],
        'Pendente': [],
        'Finalizado': [],
        'Conforme': [],
        'Conforme Parcialmente': [],
        'Não Conforme': [],
        'Outros': []
    }
    
    for sigla, info in mapeamento_siglas.items():
        status = info['status']
        relatorio = info['relatorio']
        unidade = info['unidade']
        mes = info['mes']
        mes_nome = nomes_completos.get(mes, f"Mês {mes}")
        
        item = {
            'sigla': sigla,
            'relatorio': relatorio,
            'unidade': unidade,
            'mes': mes_nome,
            'status': status
        }
        
        # Classificar por status
        status_lower = status.lower()
        if 'não iniciado' in status_lower or 'nao iniciado' in status_lower:
            siglas_por_status['Não Iniciado'].append(item)
        elif 'pendente' in status_lower:
            siglas_por_status['Pendente'].append(item)
        elif 'finalizado' in status_lower:
            siglas_por_status['Finalizado'].append(item)
        elif 'conforme' in status_lower:
            if 'parcial' in status_lower:
                siglas_por_status['Conforme Parcialmente'].append(item)
            else:
                siglas_por_status['Conforme'].append(item)
        elif 'não conforme' in status_lower or 'nao conforme' in status_lower:
            siglas_por_status['Não Conforme'].append(item)
        else:
            siglas_por_status['Outros'].append(item)
    
    # Criar containers por status
    containers_status = []
    
    for status_nome, itens in siglas_por_status.items():
        if itens:
            itens_ordenados = sorted(itens, key=lambda x: x['sigla'])
            
            # Dividir em colunas para melhor visualização
            colunas = []
            itens_por_coluna = 10
            for i in range(0, len(itens_ordenados), itens_por_coluna):
                coluna_itens = itens_ordenados[i:i + itens_por_coluna]
                
                lista_coluna = []
                for item in coluna_itens:
                    cores = get_status_color(item['status'])
                    
                    lista_coluna.append(html.Div([
                        html.Span(f"{item['sigla']}: ", style={
                            'fontWeight': 'bold',
                            'color': cores['text_color'],
                            'fontSize': '10px',
                            'minWidth': '50px',
                            'display': 'inline-block',
                            'backgroundColor': cores['bg_color'],
                            'padding': '2px 4px',
                            'borderRadius': '2px',
                            'border': f'1px solid {cores["border_color"]}',
                            'marginRight': '5px'
                        }),
                        html.Span(f"{item['relatorio'][:35]}", style={
                            'color': '#2c3e50',
                            'fontSize': '10px',
                            'overflow': 'hidden',
                            'textOverflow': 'ellipsis',
                            'whiteSpace': 'nowrap',
                            'maxWidth': '200px',
                            'display': 'inline-block'
                        }),
                        html.Span(f" ({item['unidade']}, {item['mes']})", style={
                            'color': '#7f8c8d',
                            'fontSize': '9px',
                            'fontStyle': 'italic'
                        })
                    ], style={
                        'marginBottom': '3px',
                        'padding': '3px 5px',
                        'borderBottom': '1px solid #f0f0f0',
                        'display': 'flex',
                        'alignItems': 'center'
                    }))
                
                colunas.append(html.Div(lista_coluna, style={
                    'flex': '1',
                    'minWidth': '250px',
                    'marginRight': '15px'
                }))
            
            # Container para este status
            status_color_map = {
                'Não Iniciado': '#c0392b',
                'Pendente': '#f39c12',
                'Finalizado': '#27ae60',
                'Conforme': '#27ae60',
                'Conforme Parcialmente': '#f39c12',
                'Não Conforme': '#c0392b',
                'Outros': '#7f8c8d'
            }
            
            cor_status = status_color_map.get(status_nome, '#7f8c8d')
            
            containers_status.append(html.Div([
                html.H5(f"📋 {status_nome.upper()} ({len(itens)} relatórios)", style={
                    'color': cor_status,
                    'marginBottom': '8px',
                    'fontSize': '12px',
                    'borderBottom': f'2px solid {cor_status}',
                    'paddingBottom': '3px',
                    'fontWeight': 'bold'
                }),
                html.Div(colunas, style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '10px'})
            ], style={
                'backgroundColor': '#f8f9fa',
                'padding': '12px 15px',
                'borderRadius': '4px',
                'marginBottom': '15px',
                'border': f'1px solid {cor_status}20'  # Cor com transparência
            }))
    
    # Container da lista completa
    lista_completa_container = html.Div([
        html.H4("📋 LISTA COMPLETA DE RELATÓRIOS POR SIGLA", style={
            'marginBottom': '15px',
            'color': '#2c3e50',
            'fontSize': '14px',
            'textAlign': 'center',
            'backgroundColor': '#ecf0f1',
            'padding': '10px',
            'borderRadius': '4px',
            'border': '1px solid #bdc3c7'
        }),
        html.Div(containers_status, style={'maxHeight': '300px', 'overflowY': 'auto', 'padding': '5px'})
    ], style={
        'marginTop': '20px',
        'padding': '15px',
        'backgroundColor': 'white',
        'borderRadius': '5px',
        'border': '1px solid #dee2e6',
        'boxShadow': '0 1px 3px rgba(0,0,0,0.1)'
    })
    
    # Legenda de cores
    legenda_cores = html.Div([
        html.P("🎨 LEGENDA DE CORES:", style={
            'marginBottom': '8px',
            'color': '#2c3e50',
            'fontSize': '11px',
            'fontWeight': 'bold'
        }),
        html.Div([
            html.Div([
                html.Div(style={
                    'width': '12px',
                    'height': '12px',
                    'backgroundColor': '#c0392b',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #c0392b'
                }),
                html.Span("Não Iniciado", style={'color': '#2c3e50', 'fontSize': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '15px'}),
            
            html.Div([
                html.Div(style={
                    'width': '12px',
                    'height': '12px',
                    'backgroundColor': '#f39c12',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #f39c12'
                }),
                html.Span("Pendente", style={'color': '#2c3e50', 'fontSize': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '15px'}),
            
            html.Div([
                html.Div(style={
                    'width': '12px',
                    'height': '12px',
                    'backgroundColor': '#27ae60',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #27ae60'
                }),
                html.Span("Finalizado/Conforme", style={'color': '#2c3e50', 'fontSize': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '15px'}),
            
            html.Div([
                html.Div(style={
                    'width': '12px',
                    'height': '12px',
                    'backgroundColor': '#fff8e1',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #f39c12'
                }),
                html.Span("Conforme Parcialmente", style={'color': '#2c3e50', 'fontSize': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center'})
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '10px', 'alignItems': 'center'})
    ], style={
        'backgroundColor': '#f8f9fa',
        'padding': '10px 12px',
        'borderRadius': '4px',
        'marginBottom': '15px',
        'border': '1px solid #dee2e6',
        'fontSize': '11px'
    })
    
    titulo_matriz = f"📋 MATRIZ DE RISCO - {ano_filtro} ({len(df_risco_filtrado)} registros, {len(siglas_unicas)} siglas)"
    
    return html.Div([
        html.H4(titulo_matriz, style={
            'marginBottom': '15px',
            'color': '#2c3e50',
            'fontSize': '16px',
            'display': 'flex',
            'alignItems': 'center',
            'gap': '8px',
            'backgroundColor': '#ecf0f1',
            'padding': '10px',
            'borderRadius': '4px'
        }),
        legenda_cores,
        html.P("📊 Cada célula mostra TODAS as siglas dos relatórios daquela unidade/mês", 
               style={'color': '#7f8c8d', 'marginBottom': '10px', 'fontSize': '11px'}),
        tabela_container,
        lista_completa_container
    ], style={
        'marginTop': '20px',
        'padding': '15px',
        'backgroundColor': 'white',
        'borderRadius': '5px',
        'boxShadow': '0 1px 3px rgba(0,0,0,0.1)',
        'overflow': 'hidden'
    })
