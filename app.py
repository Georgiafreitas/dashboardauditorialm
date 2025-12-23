from dash import Dash, html, dcc, Input, Output, dash_table
import pandas as pd
import plotly.express as px
import os
import unicodedata
from datetime import datetime
import dash_auth  # Importação para autenticação
import re

print("🚀 Iniciando Dashboard de Auditoria...")

# ========== CONFIGURAÇÃO DE AUTENTICAÇÃO ==========
USUARIOS_VALIDOS = {
    'admin': 'wne@2026',
    'diretoria': 'lagoa@2026'
}

# ========== DICIONÁRIO DE SIGLAS FORNECIDAS PELO USUÁRIO ==========
DICIONARIO_SIGLAS = {
    'BM': 'Baixas Manuais',
    'BO': 'Bonificações',
    'FF': 'Faturamento sem Financeiro',
    'PR': 'Prorrogação',
    'PM': 'Pagamento Manual',
    'TM': 'Títulos Pagos a Menor',
    'RJ': 'Recebimento sem Juros',
    'VD': 'Vendas',
    'DC': 'Descontos Concedidos',
    
    # SIGLAS ADICIONAIS PARA COMPLETAR
    'CL': 'Checklist',
    'AU': 'Auditoria',
    'NC': 'Não Conformidade',
    'MT': 'Monitoramento',
    'AD': 'Análise Documental',
    'RE': 'Relatório',
    'CH': 'Check',
    'IN': 'Inspeção',
    'VI': 'Visita',
    'RA': 'Risco Alto',
    'RM': 'Risco Médio',
    'RB': 'Risco Baixo',
    'RP': 'Risco Potencial',
    'RC': 'Risco Crítico',
    'RV': 'Risco Vulnerabilidade'
}

def obter_significado_sigla(sigla):
    """Retorna o significado de uma sigla, ou a própria sigla se não encontrada"""
    sigla_upper = str(sigla).strip().upper()
    
    # Verifica se a sigla completa está no dicionário
    if sigla_upper in DICIONARIO_SIGLAS:
        return DICIONARIO_SIGLAS[sigla_upper]
    
    # Se não encontrou, verifica as 2 primeiras letras
    if len(sigla_upper) >= 2:
        sigla_2 = sigla_upper[:2]
        if sigla_2 in DICIONARIO_SIGLAS:
            return DICIONARIO_SIGLAS[sigla_2]
    
    # Se ainda não encontrou, retorna a sigla original
    return sigla_upper

# ========== HELPERS ==========
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
    elif 'conforme' in status_str:
        return {
            'bg_color': '#eafaf1',
            'text_color': '#27ae60',
            'border_color': '#27ae60'
        }
    elif 'conforme parcialmente' in status_str or 'parcial' in status_str:
        return {
            'bg_color': '#fff8e1',
            'text_color': '#f39c12',
            'border_color': '#f39c12'
        }
    elif 'não conforme' in status_str or 'nao conforme' in status_str:
        return {
            'bg_color': '#fdecea',
            'text_color': '#c0392b',
            'border_color': '#c0392b'
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
        if pd.isna(finalizacao_str) or str(finalizacao_str).strip() in ['', 'NaT', 'None']:
            return "Não Concluído"
        
        prazo = pd.to_datetime(prazo_str, errors='coerce', dayfirst=True)
        finalizacao = pd.to_datetime(finalizacao_str, errors='coerce', dayfirst=True)
        
        if pd.isna(prazo) or pd.isna(finalizacao):
            return "Não Concluído"
        
        if finalizacao <= prazo:
            return "Concluído no Prazo"
        else:
            return "Concluído Fora do Prazo"
            
    except:
        return "Não Concluído"

def formatar_data(data_str):
    """Formata a data para DD/MM/YYYY (formato brasileiro)"""
    try:
        if pd.isna(data_str) or str(data_str).strip() in ['', 'NaT', 'None']:
            return ""
        
        if isinstance(data_str, (pd.Timestamp, datetime)):
            return data_str.strftime('%d/%m/%Y')
        
        data = pd.to_datetime(data_str, errors='coerce', dayfirst=True)
        if pd.isna(data):
            return str(data_str)
        
        return data.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"Erro ao formatar data '{data_str}': {e}")
        return str(data_str)

def criar_sigla_relatorio(relatorio, index):
    """Cria uma sigla para o relatório - versão simplificada para usar siglas do dicionário"""
    if pd.isna(relatorio) or str(relatorio).strip() == '':
        return f"R{index:03d}"
    
    relatorio_str = str(relatorio).strip().upper()
    
    # Procura por palavras-chave que correspondam às siglas conhecidas
    palavras_chave = {
        'BAIXA': 'BM',
        'MANUAL': 'BM',
        'BONIF': 'BO',
        'BONIFICA': 'BO',
        'FATURAMENTO': 'FF',
        'FINANCEIRO': 'FF',
        'PRORROGA': 'PR',
        'PAGAMENTO': 'PM',
        'MANUAL': 'PM',
        'TITULO': 'TM',
        'MENOR': 'TM',
        'RECEBIMENTO': 'RJ',
        'JUROS': 'RJ',
        'VENDA': 'VD',
        'DESCONTO': 'DC',
        'CONCEDIDO': 'DC',
        'CHECKLIST': 'CL',
        'AUDITORIA': 'AU',
        'NÃO CONFORME': 'NC',
        'NAO CONFORME': 'NC',
        'MONITORAMENTO': 'MT',
        'ANALISE': 'AD',
        'ANÁLISE': 'AD',
        'RELATORIO': 'RE',
        'RELATÓRIO': 'RE',
        'CHECK': 'CH',
        'INSPECAO': 'IN',
        'INSPEÇÃO': 'IN',
        'VISITA': 'VI'
    }
    
    # Verifica se alguma palavra-chave está no nome do relatório
    for palavra, sigla in palavras_chave.items():
        if palavra in relatorio_str:
            return sigla
    
    # Se não encontrou palavra-chave, tenta extrair sigla do início do nome
    palavras = relatorio_str.split()
    if palavras:
        # Tenta combinar primeiras letras com siglas conhecidas
        primeira_palavra = palavras[0]
        if len(primeira_palavra) >= 2:
            sigla_tentativa = primeira_palavra[:2].upper()
            if sigla_tentativa in DICIONARIO_SIGLAS:
                return sigla_tentativa
        
        # Tenta criar sigla com iniciais
        if len(palavras) >= 2:
            sigla_tentativa = ''.join([p[0] for p in palavras[:2]])
            if sigla_tentativa in DICIONARIO_SIGLAS:
                return sigla_tentativa
    
    # Se nada funcionou, usa as 2 primeiras letras da primeira palavra
    if palavras:
        return palavras[0][:2].upper()
    
    # Último recurso
    return f"R{index:03d}"

def criar_matriz_risco_anual(df_risco_filtrado, ano_filtro):
    """Cria matriz de risco COMPACTA com TODAS as siglas dos relatórios"""
    
    if df_risco_filtrado is None or len(df_risco_filtrado) == 0:
        return html.Div([
            html.H3("📋 Matriz Auditoria Risco", style={'fontSize': '16px'}),
            html.P("Nenhum dado encontrado para o ano selecionado.", 
                   style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '20px', 'fontSize': '12px'})
        ], style={'marginTop':'20px'})
    
    # Obter unidades únicas
    unidades = sorted(df_risco_filtrado['Unidade'].dropna().unique())
    
    # Lista de meses do ano (1 a 12)
    meses_ano = list(range(1, 13))
    nomes_meses = {
        1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN',
        7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
    }
    
    print(f"\n📊 CRIANDO MATRIZ DE RISCO COMPACTA PARA O ANO {ano_filtro}")
    print(f"  Unidades: {len(unidades)}")
    print(f"  Total de registros: {len(df_risco_filtrado)}")
    
    # Conjunto para armazenar TODAS as siglas únicas encontradas
    siglas_encontradas = set()
    
    # Criar estrutura de dados para a matriz - GUARDA TODOS OS REGISTROS
    matriz_data = []
    
    for unidade_nome in unidades:
        linha = {'Unidade': unidade_nome}
        df_unidade = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade_nome]
        
        for mes in meses_ano:
            # Filtrar por mês
            df_unidade['Mes'] = pd.to_numeric(df_unidade['Mes'], errors='coerce')
            df_mes = df_unidade[df_unidade['Mes'] == mes]
            
            if len(df_mes) > 0:
                # Para cada relatório no mês - AGORA PEGA TODOS
                siglas_no_mes = []
                
                for _, row in df_mes.iterrows():
                    sigla = str(row.get('Sigla', '')).strip().upper()
                    status = str(row.get('Status', 'Sem Status'))
                    
                    if sigla:  # Só adiciona se tiver sigla
                        # Adicionar ao conjunto de siglas encontradas
                        siglas_encontradas.add(sigla)
                        
                        # Obter cor baseada no status
                        cores = get_status_color(status)
                        
                        # Armazenar sigla para este mês
                        siglas_no_mes.append({
                            'sigla': sigla,
                            'status': status,
                            'cor': cores,
                            'relatorio': row.get('Relatorio', '')
                        })
                
                if siglas_no_mes:
                    # Ordenar siglas alfabeticamente
                    siglas_no_mes.sort(key=lambda x: x['sigla'])
                    
                    # DEBUG: Verificar quantas siglas temos
                    print(f"  Unidade: {unidade_nome}, Mês: {mes}, Siglas: {len(siglas_no_mes)}")
                    
                    # Determinar tamanho da fonte baseado na quantidade de siglas
                    num_siglas = len(siglas_no_mes)
                    if num_siglas <= 4:
                        font_size = '8px'
                        padding = '2px 3px'
                        min_width = '28px'
                        height = '20px'
                        gap = '1px'
                        grid_cols = min(num_siglas, 4)
                    elif num_siglas <= 8:
                        font_size = '7px'
                        padding = '1px 2px'
                        min_width = '26px'
                        height = '18px'
                        gap = '1px'
                        grid_cols = 4
                    else:
                        font_size = '6px'
                        padding = '1px 2px'
                        min_width = '24px'
                        height = '16px'
                        gap = '0.5px'
                        grid_cols = 5  # Mais colunas para muitas siglas
                    
                    # Criar container ultra compacto para TODAS as siglas
                    siglas_html = []
                    for item in siglas_no_mes:
                        title_text = f"{item['sigla']}: {item['status']}"
                        if item['relatorio']:
                            title_text += f"\n{item['relatorio'][:50]}..."
                        
                        siglas_html.append(html.Div(
                            item['sigla'],
                            style={
                                'fontSize': font_size,
                                'fontWeight': '600',
                                'color': item['cor']['text_color'],
                                'textAlign': 'center',
                                'backgroundColor': item['cor']['bg_color'],
                                'padding': padding,
                                'margin': '0',
                                'borderRadius': '2px',
                                'border': f'1px solid {item["cor"]["border_color"]}',
                                'minWidth': min_width,
                                'width': min_width,
                                'height': height,
                                'display': 'flex',
                                'alignItems': 'center',
                                'justifyContent': 'center',
                                'cursor': 'default',
                                'boxShadow': '0 0.5px 1px rgba(0,0,0,0.05)',
                                'overflow': 'hidden',
                                'flexShrink': '0',
                                'flexGrow': '0'
                            },
                            title=title_text
                        ))
                    
                    # Calcular altura da célula baseado no número de linhas necessárias
                    rows_needed = (num_siglas + grid_cols - 1) // grid_cols
                    cell_height = max(35, rows_needed * (int(height.replace('px', '')) + 3))
                    
                    # Criar container com grid compacto
                    linha[mes] = html.Div(
                        siglas_html,
                        style={
                            'display': 'grid',
                            'gridTemplateColumns': f'repeat({grid_cols}, 1fr)',
                            'gap': gap,
                            'padding': '2px',
                            'justifyContent': 'center',
                            'alignItems': 'center',
                            'minHeight': f'{cell_height}px',
                            'height': f'{cell_height}px',
                            'borderRadius': '2px',
                            'backgroundColor': '#f8fafc',
                            'width': '100%',
                            'boxSizing': 'border-box',
                            'overflow': 'hidden'
                        }
                    )
                else:
                    linha[mes] = html.Div("-", 
                        style={
                            'color': '#bdc3c7', 
                            'fontSize': '9px', 
                            'padding': '5px 0',
                            'textAlign': 'center',
                            'fontStyle': 'italic',
                            'height': '35px',
                            'display': 'flex',
                            'alignItems': 'center',
                            'justifyContent': 'center'
                        })
            else:
                linha[mes] = html.Div("-", 
                    style={
                        'color': '#bdc3c7', 
                        'fontSize': '9px', 
                        'padding': '5px 0',
                        'textAlign': 'center',
                        'fontStyle': 'italic',
                        'height': '35px',
                        'display': 'flex',
                        'alignItems': 'center',
                        'justifyContent': 'center'
                    })
        
        matriz_data.append(linha)
    
    # DEBUG: Mostrar todas as siglas encontradas
    print(f"\n📋 SIGLAS ENCONTRADAS ({len(siglas_encontradas)} total):")
    for i, sigla in enumerate(sorted(siglas_encontradas)):
        print(f"  {i+1}. {sigla}")
    
    # Ordenar siglas encontradas alfabeticamente
    siglas_ordenadas = sorted(siglas_encontradas)
    
    # Criar tabela HTML com design ULTRA COMPACTO
    tabela_cabecalho = [html.Th("UNIDADE", style={
        'backgroundColor': '#2c3e50',
        'color': 'white',
        'padding': '6px 8px',
        'textAlign': 'center',
        'fontWeight': '600',
        'border': '1px solid #1a252f',
        'minWidth': '100px',
        'width': '100px',
        'fontSize': '10px',
        'position': 'sticky',
        'left': '0',
        'zIndex': '2',
        'boxShadow': '1px 0 1px rgba(0,0,0,0.1)',
        'height': '30px'
    })]
    
    for mes in meses_ano:
        tabela_cabecalho.append(html.Th(
            html.Div([
                html.Div(nomes_meses[mes], style={'fontSize': '9px', 'fontWeight': '600', 'marginBottom': '1px'}),
                html.Div(str(mes), style={'fontSize': '7px', 'opacity': '0.8', 'fontWeight': '400'})
            ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center'}),
            style={
                'backgroundColor': '#2c3e50',
                'color': 'white',
                'padding': '4px 2px',
                'textAlign': 'center',
                'fontWeight': '600',
                'border': '1px solid #1a252f',
                'minWidth': '45px',
                'width': '45px',
                'fontSize': '9px',
                'height': '30px',
                'verticalAlign': 'middle'
            }
        ))
    
    tabela_linhas = []
    
    for i, linha in enumerate(matriz_data):
        bg_color = '#ffffff' if i % 2 == 0 else '#f8f9fa'
        
        celulas = [html.Td(
            html.Div([
                html.Div(linha['Unidade'], style={
                    'fontSize': '9px',
                    'fontWeight': '600',
                    'overflow': 'hidden',
                    'textOverflow': 'ellipsis',
                    'whiteSpace': 'nowrap',
                    'color': '#2c3e50',
                    'textAlign': 'left',
                    'padding': '0 3px'
                })
            ]),
            style={
                'backgroundColor': bg_color,
                'padding': '6px 3px',
                'textAlign': 'left',
                'border': '1px solid #dde1e6',
                'fontSize': '9px',
                'minWidth': '100px',
                'width': '100px',
                'height': '35px',
                'position': 'sticky',
                'left': '0',
                'zIndex': '1',
                'boxShadow': '1px 0 1px rgba(0,0,0,0.05)',
                'verticalAlign': 'top'
            }
        )]
        
        for mes in meses_ano:
            conteudo = linha.get(mes, html.Div("-", style={'color': '#bdc3c7', 'fontSize': '9px', 'padding': '5px 0', 'textAlign': 'center'}))
            celulas.append(html.Td(
                conteudo,
                style={
                    'backgroundColor': bg_color,
                    'padding': '0',
                    'textAlign': 'center',
                    'border': '1px solid #dde1e6',
                    'verticalAlign': 'top',
                    'minHeight': '35px',
                    'minWidth': '45px',
                    'width': '45px',
                    'height': 'auto',
                    'maxHeight': '80px',
                    'overflow': 'hidden'
                }
            ))
        
        tabela_linhas.append(html.Tr(celulas))
    
    tabela_html = html.Table([
        html.Thead(html.Tr(tabela_cabecalho)),
        html.Tbody(tabela_linhas)
    ], style={
        'width': '100%',
        'borderCollapse': 'separate',
        'borderSpacing': '0',
        'marginTop': '3px',
        'fontFamily': "'Segoe UI', 'Helvetica Neue', Arial, sans-serif",
        'fontSize': '9px',
        'tableLayout': 'fixed',
        'borderRadius': '3px',
        'overflow': 'hidden',
        'boxShadow': '0 1px 2px rgba(0,0,0,0.1)'
    })
    
    # Container da tabela - COMPACTO
    tabela_container = html.Div(
        tabela_html,
        style={
            'overflowX': 'auto',
            'maxWidth': '100%',
            'marginTop': '8px',
            'borderRadius': '3px',
            'maxHeight': '400px',  # ALTURA FIXA COMPACTA
            'overflowY': 'hidden',  # SEM rolagem vertical
            'border': '1px solid #dde1e6',
            'backgroundColor': 'white'
        }
    )
    
    # Criar lista de siglas e seus significados com design COMPACTO
    lista_siglas = []
    
    for sigla in siglas_ordenadas:
        significado = obter_significado_sigla(sigla)
        lista_siglas.append(html.Div([
            html.Span(f"{sigla}", style={
                'fontWeight': '600',
                'color': '#2c3e50',
                'fontSize': '9px',
                'minWidth': '30px',
                'display': 'inline-block',
                'backgroundColor': '#ecf0f1',
                'padding': '2px 4px',
                'borderRadius': '2px',
                'marginRight': '6px',
                'textAlign': 'center',
                'border': '1px solid #bdc3c7',
                'flexShrink': '0'
            }),
            html.Span(significado, style={
                'color': '#2c3e50',
                'fontSize': '9px',
                'lineHeight': '1.2',
                'flexGrow': '1'
            })
        ], style={
            'marginBottom': '4px',
            'padding': '3px 5px',
            'borderBottom': '1px solid #ecf0f1',
            'backgroundColor': '#ffffff',
            'borderRadius': '2px',
            'display': 'flex',
            'alignItems': 'center',
            'minHeight': '24px'
        }))
    
    # Se não encontrou siglas conhecidas, mostrar mensagem
    if not siglas_encontradas:
        lista_siglas = [html.P("Nenhuma sigla conhecida encontrada nos dados.", 
                               style={'color': '#7f8c8d', 'fontSize': '10px', 'textAlign': 'center', 'padding': '10px'})]
    
    # Container da lista de siglas com design COMPACTO
    lista_siglas_container = html.Div([
        html.Div([
            html.H4("📋 LEGENDA DE SIGLAS", style={
                'margin': '0 0 8px 0',
                'color': '#2c3e50',
                'fontSize': '11px',
                'fontWeight': '600'
            }),
            html.P(f"Total de {len(siglas_encontradas)} siglas distintas encontradas", 
                   style={'color': '#7f8c8d', 'marginBottom': '8px', 'fontSize': '9px'})
        ], style={
            'backgroundColor': '#ecf0f1',
            'padding': '8px',
            'borderRadius': '2px 2px 0 0',
            'borderBottom': '1px solid #bdc3c7'
        }),
        html.Div(lista_siglas, style={
            'maxHeight': '120px',  # ALTURA COMPACTA
            'overflowY': 'auto', 
            'padding': '8px',
            'backgroundColor': '#ffffff'
        })
    ], style={
        'marginTop': '10px',
        'borderRadius': '3px',
        'border': '1px solid #dde1e6',
        'boxShadow': '0 1px 2px rgba(0,0,0,0.1)',
        'overflow': 'hidden'
    })
    
    # Legenda de cores com design COMPACTO
    legenda_cores = html.Div([
        html.Div([
            html.H5("🎨 LEGENDA DE STATUS", style={
                'margin': '0 0 6px 0',
                'color': '#2c3e50',
                'fontSize': '10px',
                'fontWeight': '600'
            })
        ], style={'marginBottom': '6px'}),
        html.Div([
            html.Div([
                html.Div(style={
                    'width': '10px',
                    'height': '10px',
                    'backgroundColor': '#c0392b',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #c0392b',
                    'flexShrink': '0'
                }),
                html.Span("Não Iniciado", style={'color': '#2c3e50', 'fontSize': '9px', 'fontWeight': '500'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '12px', 'marginBottom': '4px'}),
            
            html.Div([
                html.Div(style={
                    'width': '10px',
                    'height': '10px',
                    'backgroundColor': '#f39c12',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #f39c12',
                    'flexShrink': '0'
                }),
                html.Span("Pendente", style={'color': '#2c3e50', 'fontSize': '9px', 'fontWeight': '500'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '12px', 'marginBottom': '4px'}),
            
            html.Div([
                html.Div(style={
                    'width': '10px',
                    'height': '10px',
                    'backgroundColor': '#27ae60',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #27ae60',
                    'flexShrink': '0'
                }),
                html.Span("Finalizado", style={'color': '#2c3e50', 'fontSize': '9px', 'fontWeight': '500'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '12px', 'marginBottom': '4px'}),
            
            html.Div([
                html.Div(style={
                    'width': '10px',
                    'height': '10px',
                    'backgroundColor': '#fff8e1',
                    'borderRadius': '2px',
                    'marginRight': '5px',
                    'border': '1px solid #f39c12',
                    'flexShrink': '0'
                }),
                html.Span("Conforme Parcial", style={'color': '#2c3e50', 'fontSize': '9px', 'fontWeight': '500'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '4px'})
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '8px', 'alignItems': 'center'})
    ], style={
        'backgroundColor': '#f8fafc',
        'padding': '8px',
        'borderRadius': '3px',
        'marginBottom': '10px',
        'border': '1px solid #dde1e6',
        'fontSize': '9px'
    })
    
    titulo_matriz = html.Div([
        html.H4(f"📋 MATRIZ DE RISCO - {ano_filtro}", style={
            'margin': '0 0 2px 0',
            'color': '#2c3e50',
            'fontSize': '13px',
            'fontWeight': '600'
        }),
        html.P(f"{len(df_risco_filtrado)} registros • {len(unidades)} unidades • {len(siglas_encontradas)} siglas", 
               style={'color': '#7f8c8d', 'margin': '0', 'fontSize': '9px'})
    ], style={
        'marginBottom': '10px',
        'padding': '8px',
        'backgroundColor': '#ecf0f1',
        'borderRadius': '3px',
        'border': '1px solid #bdc3c7',
        'fontSize': '10px'
    })
    
    # DEBUG: Mostrar quantos registros estão sendo processados
    print(f"\n📊 RESUMO DA MATRIZ:")
    print(f"  Unidades processadas: {len(unidades)}")
    print(f"  Registros totais: {len(df_risco_filtrado)}")
    print(f"  Siglas únicas: {len(siglas_ordenadas)}")
    print(f"  Tamanho da matriz: {len(unidades)} linhas x 13 colunas")
    
    return html.Div([
        titulo_matriz,
        legenda_cores,
        html.P("Cada célula mostra TODAS as siglas dos relatórios daquela unidade/mês", 
               style={'color': '#7f8c8d', 'marginBottom': '8px', 'fontSize': '9px', 'textAlign': 'center'}),
        html.P(f"Dica: Passe o mouse sobre uma sigla para ver detalhes", 
               style={'color': '#3498db', 'marginBottom': '5px', 'fontSize': '8px', 'textAlign': 'center', 'fontStyle': 'italic'}),
        tabela_container,
        lista_siglas_container
    ], style={
        'marginTop': '10px',
        'padding': '10px',
        'backgroundColor': 'white',
        'borderRadius': '3px',
        'boxShadow': '0 1px 3px rgba(0,0,0,0.05)'
    })

def carregar_dados_da_planilha():
    planilha_path = 'base_auditoria.xlsx'
    if not os.path.exists(planilha_path):
        print("❌ Planilha não encontrada:", planilha_path)
        return None, None, None, None

    try:
        print(f"📁 Carregando dados da planilha: {planilha_path}")

        # Leitura das planilhas
        print("  Lendo aba Checklist_Unidades...")
        df_checklist = pd.read_excel(planilha_path, sheet_name='Checklist_Unidades', engine='openpyxl')
        
        print("  Lendo aba Politicas...")
        df_politicas = pd.read_excel(planilha_path, sheet_name='Politicas', engine='openpyxl')
        
        print("  Lendo aba Auditoria_Risco...")
        df_risco = pd.read_excel(planilha_path, sheet_name='Auditoria_Risco', engine='openpyxl')
        
        print("  Lendo aba Melhorias_Logistica...")
        df_melhorias = pd.read_excel(planilha_path, sheet_name='Melhorias_Logistica', engine='openpyxl')

        print("✅ Leitura inicial da planilha concluída. Processando dados...")

        for i, df in enumerate([df_checklist, df_politicas, df_risco, df_melhorias]):
            if df is not None:
                print(f"\n{'='*50}")
                print(f"Processando aba {i}:")
                print(f"Colunas originais: {df.columns.tolist()}")
                print(f"Total de registros: {len(df)}")
                
                df = normalize_df_columns(df)
                print(f"Colunas após normalização: {df.columns.tolist()}")
                
                # CORREÇÃO ESPECÍFICA PARA CADA ABA
                if i == 0:  # df_checklist
                    print("📋 Processando CHECKLIST...")
                    
                    # Normalizar Status
                    if 'Status' in df.columns:
                        df['Status'] = df['Status'].astype(str).str.strip()
                        print(f"  Status únicos antes: {df['Status'].unique()[:10]}")
                        df['Status'] = df['Status'].apply(canonical_status)
                        print(f"  Status únicos depois: {df['Status'].unique()[:10]}")
                    
                    # Processar datas
                    if 'Data' in df.columns:
                        print(f"  Processando coluna Data...")
                        print(f"  Tipo da coluna Data: {df['Data'].dtype}")
                        print(f"  Amostra de datas: {df['Data'].head(5).tolist()}")
                        
                        # Converter a coluna Data para datetime
                        df['Data_DT'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
                        
                        falhas = df['Data_DT'].isna().sum()
                        if falhas > 0:
                            print(f"  ⚠️ {falhas} datas não puderam ser convertidas")
                        
                        # Extrair Ano e Mes
                        df['Ano'] = df['Data_DT'].dt.year
                        df['Mes'] = df['Data_DT'].dt.month
                        
                        df['Ano'] = df['Ano'].fillna(0).astype(int)
                        df['Mes'] = df['Mes'].fillna(0).astype(int)
                        df['Ano'] = df['Ano'].replace(0, pd.NA)
                        df['Mes'] = df['Mes'].replace(0, pd.NA)
                        
                        print(f"  Ano únicos: {df['Ano'].dropna().unique()}")
                        print(f"  Mês únicos: {df['Mes'].dropna().unique()}")
                        
                        df['Data'] = df['Data_DT'].apply(
                            lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else ''
                        )
                        
                        df = df.drop(columns=['Data_DT'])
                
                elif i == 1:  # df_politicas
                    print("📑 Processando POLÍTICAS...")
                    if 'Status' in df.columns:
                        df['Status'] = df['Status'].apply(canonical_status)
                
                elif i == 2:  # df_risco
                    print("🔄 Processando dados de RISCO...")
                    print(f"  🔍 Colunas disponíveis: {df.columns.tolist()}")
                    
                    # 1. Encontrar e processar coluna de Status
                    coluna_status = None
                    for col in df.columns:
                        if 'status' in col.lower():
                            coluna_status = col
                            break
                    
                    if coluna_status:
                        print(f"  ✅ Coluna de Status encontrada: '{coluna_status}'")
                        df[coluna_status] = df[coluna_status].astype(str).str.strip()
                        df['Status'] = df[coluna_status].apply(canonical_status)
                        print(f"  Status únicos: {df['Status'].unique()[:10]}")
                    else:
                        print(f"  ⚠️ Coluna de Status não encontrada")
                        df['Status'] = "Não Iniciado"
                    
                    # 2. Encontrar e processar coluna de Data
                    coluna_data = None
                    for col in df.columns:
                        if col.lower() == 'data':
                            coluna_data = col
                            break
                    
                    if coluna_data:
                        print(f"\n  ✅ Coluna de Data encontrada: '{coluna_data}'")
                        print(f"  Tipo da coluna Data: {df[coluna_data].dtype}")
                        
                        # Converter datas para datetime
                        print(f"\n  🔍 Convertendo datas para datetime...")
                        
                        # Primeiro, converter tudo para string para análise
                        df['Data_Str'] = df[coluna_data].astype(str)
                        
                        # Tentar extrair datas de diferentes formatos
                        def converter_data_agressiva(data_str):
                            if pd.isna(data_str) or data_str in ['nan', 'NaT', 'None', '']:
                                return pd.NaT
                            
                            data_str = str(data_str).strip()
                            
                            # Padrões comuns
                            padroes = [
                                r'(\d{1,2})/(\d{1,2})/(\d{4})',
                                r'(\d{1,2})-(\d{1,2})-(\d{4})',
                                r'(\d{1,2})\.(\d{1,2})\.(\d{4})',
                                r'(\d{4})-(\d{1,2})-(\d{1,2})',
                            ]
                            
                            for padrao in padroes:
                                match = re.search(padrao, data_str)
                                if match:
                                    grupos = match.groups()
                                    if len(grupos) == 3:
                                        try:
                                            if '/' in data_str or '-' in data_str:
                                                if int(grupos[0]) <= 31:
                                                    dia, mes, ano = grupos
                                                    dia = int(dia)
                                                    mes = int(mes)
                                                    ano = int(ano)
                                                    return pd.Timestamp(year=ano, month=mes, day=dia)
                                                else:
                                                    ano, mes, dia = grupos
                                                    dia = int(dia)
                                                    mes = int(mes)
                                                    ano = int(ano)
                                                    return pd.Timestamp(year=ano, month=mes, day=dia)
                                        except:
                                            continue
                            
                            # Se não encontrou padrão, tentar pandas diretamente
                            try:
                                data_dt = pd.to_datetime(data_str, dayfirst=True, errors='coerce')
                                if pd.notna(data_dt):
                                    return data_dt
                                
                                data_dt = pd.to_datetime(data_str, dayfirst=False, errors='coerce')
                                if pd.notna(data_dt):
                                    return data_dt
                                
                                data_dt = pd.to_datetime(data_str, errors='coerce')
                                return data_dt
                            except:
                                return pd.NaT
                        
                        # Aplicar conversão agressiva
                        df['Data_DT'] = df['Data_Str'].apply(converter_data_agressiva)
                        
                        # Verificar resultados
                        total = len(df)
                        sucesso = df['Data_DT'].notna().sum()
                        falhas = total - sucesso
                        
                        print(f"\n  ✅ Resultado da conversão:")
                        print(f"     Total de registros: {total}")
                        print(f"     Conversões bem-sucedidas: {sucesso} ({sucesso/total*100:.1f}%)")
                        print(f"     Falhas: {falhas}")
                        
                        if falhas > 0:
                            print(f"  ⚠️ Exemplos de datas que falharam:")
                            falhas_df = df[df['Data_DT'].isna()]
                            for j, data in enumerate(falhas_df['Data_Str'].head(5).tolist()):
                                print(f"      {j+1}. '{data}'")
                        
                        # Extrair mês e ano
                        df['Mes'] = df['Data_DT'].dt.month
                        df['Ano'] = df['Data_DT'].dt.year
                        
                        # Converter para inteiros
                        df['Mes'] = df['Mes'].fillna(0).astype(int)
                        df['Ano'] = df['Ano'].fillna(0).astype(int)
                        df['Mes'] = df['Mes'].replace(0, pd.NA)
                        df['Ano'] = df['Ano'].replace(0, pd.NA)
                        
                        # Criar Mes_Ano para exibição
                        df['Mes_Ano'] = df.apply(
                            lambda row: f"{int(row['Mes']):02d}/{int(row['Ano'])}" 
                            if pd.notna(row['Mes']) and pd.notna(row['Ano']) 
                            else "Sem Data", 
                            axis=1
                        )
                        
                        # Formatar data para exibição
                        df['Data_Formatada'] = df['Data_DT'].apply(
                            lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else ''
                        )
                        
                        # Remover colunas temporárias
                        df = df.drop(columns=['Data_DT', 'Data_Str'])
                    else:
                        print(f"  ❌ Coluna de Data não encontrada!")
                        df['Mes'] = pd.NA
                        df['Ano'] = pd.NA
                        df['Mes_Ano'] = "Sem Data"
                        df['Data_Formatada'] = ""
                    
                    # 3. Encontrar coluna de Relatório
                    coluna_relatorio = None
                    for col in df.columns:
                        if 'relatorio' in col.lower():
                            coluna_relatorio = col
                            break
                    
                    if coluna_relatorio:
                        print(f"  ✅ Coluna de Relatório encontrada: '{coluna_relatorio}'")
                        df['Relatorio'] = df[coluna_relatorio].astype(str)
                        
                        # DEBUG: Mostrar alguns exemplos de relatórios e siglas
                        print(f"\n  🔤 Exemplos de relatórios e siglas:")
                        siglas = []
                        for idx, (relatorio, row) in enumerate(zip(df['Relatorio'].head(10), df.head(10).iterrows())):
                            sigla = criar_sigla_relatorio(relatorio, idx)
                            siglas.append(sigla)
                            print(f"     {idx+1}. '{relatorio[:50]}...' -> {sigla}")
                        
                        # Criar siglas para os relatórios usando o dicionário
                        print(f"\n  🔤 Criando siglas para TODOS os relatórios...")
                        siglas = []
                        for idx, relatorio in enumerate(df['Relatorio']):
                            sigla = criar_sigla_relatorio(relatorio, idx)
                            siglas.append(sigla)
                        df['Sigla'] = siglas
                        
                        # Contar siglas únicas
                        siglas_unicas = df['Sigla'].nunique()
                        print(f"  📊 Total de siglas únicas criadas: {siglas_unicas}")
                        print(f"  ✅ Exemplos de siglas: {df['Sigla'].unique()[:20]}")
                        
                    else:
                        print(f"  ⚠️ Coluna de Relatório não encontrada")
                        df['Relatorio'] = df.get('ID', 'Sem Relatório').astype(str)
                        # Criar siglas padrão
                        siglas = []
                        for idx in range(len(df)):
                            siglas.append(f"R{idx:03d}")
                        df['Sigla'] = siglas
                    
                    # 4. Garantir coluna Unidade
                    if 'Unidade' not in df.columns:
                        for col in df.columns:
                            if 'unidade' in col.lower():
                                df['Unidade'] = df[col].astype(str)
                                print(f"  ✅ Coluna Unidade mapeada de: '{col}'")
                                break
                        else:
                            print(f"  ⚠️ Coluna Unidade não encontrada, criando padrão")
                            df['Unidade'] = "Sem Unidade"
                    
                    # DEBUG: Mostrar estrutura final
                    print(f"\n  📊 ESTRUTURA FINAL DO DATAFRAME DE RISCO:")
                    print(f"     Colunas: {df.columns.tolist()}")
                    print(f"     Total de registros: {len(df)}")
                    print(f"     Unidades únicas: {df['Unidade'].nunique()}")
                    print(f"     Meses com dados: {df['Mes'].dropna().nunique()}")
                    print(f"     Anos com dados: {df['Ano'].dropna().nunique()}")
                    print(f"     Siglas únicas: {df['Sigla'].nunique()}")
                
                elif i == 3:  # df_melhorias
                    print("📈 Processando MELHORIAS...")
                    if 'Status' in df.columns:
                        df['Status'] = df['Status'].apply(canonical_status)
                
                print(f"Colunas finais: {df.columns.tolist()}")
                
                if i == 0: df_checklist = df
                elif i == 1: df_politicas = df
                elif i == 2: df_risco = df
                elif i == 3: df_melhorias = df

        print("\n" + "="*50)
        print("✅ Dados carregados da planilha com sucesso!")
        print("="*50)
        
        print(f"\n📊 RESUMO DOS DADOS CARREGADOS:")
        print(f"  Checklist: {len(df_checklist)} registros")
        print(f"  Políticas: {len(df_politicas)} registros")
        print(f"  Risco: {len(df_risco)} registros")
        print(f"  Melhorias: {len(df_melhorias)} registros")
        
        if df_risco is not None and len(df_risco) > 0:
            print(f"\n📋 DETALHES DA MATRIZ DE RISCO:")
            print(f"  Colunas: {df_risco.columns.tolist()}")
            if 'Ano' in df_risco.columns:
                anos_unicos = df_risco['Ano'].dropna().unique()
                print(f"  Anos únicos encontrados ({len(anos_unicos)}):")
                for ano in sorted(anos_unicos):
                    contagem = len(df_risco[df_risco['Ano'] == ano])
                    print(f"    {int(ano)}: {contagem} registros")
        
        return df_checklist, df_politicas, df_risco, df_melhorias

    except Exception as e:
        print(f"❌ Erro ao carregar planilha: {e}")
        import traceback
        traceback.print_exc()
        return None, None, None, None

# ========== FUNÇÕES UTILITÁRIAS ==========
def obter_anos_disponiveis(df_checklist):
    if df_checklist is None or 'Ano' not in df_checklist.columns:
        return []
    anos = sorted(df_checklist['Ano'].dropna().unique(), reverse=True)
    anos_int = []
    for ano in anos:
        try:
            ano_int = int(float(ano)) if pd.notna(ano) else None
            if ano_int and ano_int not in anos_int:
                anos_int.append(ano_int)
        except:
            continue
    return sorted(anos_int, reverse=True)

def obter_meses_disponiveis(df_checklist, ano_selecionado):
    """Retorna todos os meses (1-12) independentemente dos dados existentes"""
    nomes_meses = {
        1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 
        5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago', 
        9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
    }
    
    return [{'label': f'{nomes_meses[m]}', 'value': m} for m in range(1, 13)]

# ========== CARREGAR DADOS ==========
df_checklist, df_politicas, df_risco, df_melhorias = carregar_dados_da_planilha()
if df_checklist is None:
    app = Dash(__name__)
    server = app.server
    app.layout = html.Div([html.H1("❌ Planilha não encontrada")])
    if __name__ == '__main__':
        app.run(debug=True, port=8050)
    exit()

anos_disponiveis = obter_anos_disponiveis(df_checklist)
print(f"\nDEBUG: Anos disponíveis no filtro: {anos_disponiveis}")

# ========== APP DASH ==========
app = Dash(__name__)

# ========== APLICAR AUTENTICAÇÃO ==========
auth = dash_auth.BasicAuth(app, USUARIOS_VALIDOS)

# ========== LAYOUT DO DASHBOARD ==========
app.layout = html.Div([
    html.Div([
        html.H1("📊 DASHBOARD DE AUDITORIA", 
                style={
                    'textAlign':'center', 
                    'marginBottom':'10px',
                    'fontSize': '18px',
                    'color': '#2c3e50'
                })
    ]),
    html.Div([
        html.Div([
            html.Label("Ano:", style={'fontSize': '11px', 'fontWeight': 'bold', 'marginBottom': '2px'}),
            dcc.Dropdown(
                id='filtro-ano',
                options=[{'label':'Todos','value':'todos'}]+
                       [{'label':str(a),'value':a} for a in anos_disponiveis],
                value='todos',
                style={'fontSize': '11px', 'minHeight': '30px'}
            )
        ], style={'marginRight':'10px','width':'150px'}),
        html.Div([
            html.Label("Mês:", style={'fontSize': '11px', 'fontWeight': 'bold', 'marginBottom': '2px'}),
            dcc.Dropdown(
                id='filtro-mes',
                options=[{'label':'Todos','value':'todos'}]+
                       [{'label':f'{m}','value':m} for m in range(1, 13)],
                value='todos',
                style={'fontSize': '11px', 'minHeight': '30px'}
            )
        ], style={'marginRight':'10px','width':'120px'}),
        html.Div([
            html.Label("Unidade:", style={'fontSize': '11px', 'fontWeight': 'bold', 'marginBottom': '2px'}),
            dcc.Dropdown(
                id='filtro-unidade',
                options=[{'label':'Todas','value':'todas'}]+
                        [{'label':str(u),'value':str(u)} for u in sorted(df_checklist['Unidade'].dropna().unique())],
                value='todas',
                style={'fontSize': '11px', 'minHeight': '30px'}
            )
        ], style={'width':'160px'})
    ], style={'display':'flex','justifyContent':'center','marginBottom':'15px','flexWrap':'wrap', 'padding': '5px'}),
    html.Div(id='conteudo-principal', style={'padding':'10px', 'maxWidth': '1200px', 'margin': '0 auto'})
])

# ========== CALLBACKS ==========
@app.callback(
    Output('conteudo-principal','children'),
    [Input('filtro-ano','value'),
     Input('filtro-mes','value'),
     Input('filtro-unidade','value')]
)
def atualizar_conteudo_principal(ano, mes, unidade):
    # ---------- FILTRAR CHECKLIST ----------
    df = df_checklist.copy()
    
    print(f"\n🔍 DEBUG FILTROS: Ano='{ano}', Mês='{mes}', Unidade='{unidade}'")
    
    if 'Ano' in df.columns:
        df['Ano'] = pd.to_numeric(df['Ano'], errors='coerce')
    
    if 'Mes' in df.columns:
        df['Mes'] = pd.to_numeric(df['Mes'], errors='coerce')
    
    total_antes = len(df)
    
    if ano != 'todos':
        try:
            ano_filtro = int(ano)
            df = df[df['Ano'] == ano_filtro]
            print(f"  ✅ Filtro ANO aplicado: {ano_filtro} | Registros: {len(df)}/{total_antes}")
        except Exception as e:
            print(f"  ❌ Erro ao filtrar por ano '{ano}': {e}")
    
    if mes != 'todos':
        try:
            mes_filtro = int(mes)
            df = df[df['Mes'] == mes_filtro]
            print(f"  ✅ Filtro MÊS aplicado: {mes_filtro} | Registros: {len(df)}")
        except Exception as e:
            print(f"  ❌ Erro ao filtrar por mês '{mes}': {e}")
    
    if unidade != 'todas':
        try:
            df['Unidade'] = df['Unidade'].astype(str).str.strip()
            df = df[df['Unidade'] == unidade.strip()]
            print(f"  ✅ Filtro UNIDADE aplicado: '{unidade}' | Registros: {len(df)}")
        except Exception as e:
            print(f"  ❌ Erro ao filtrar por unidade '{unidade}': {e}")
    
    total = len(df)
    print(f"📊 TOTAL APÓS FILTROS: {total} registros")
    
    # ---------- Contagem correta dos status ----------
    if total > 0:
        df['Status'] = df['Status'].astype(str).str.strip()
        conforme = len(df[df['Status'].str.lower() == 'conforme'])
        parcial = len(df[df['Status'].str.lower().str.contains('parcial')])
        nao = len(df[df['Status'].str.lower().str.contains('não|nao')])
    else:
        conforme = 0
        parcial = 0
        nao = 0

    # ---------- KPIs GERAIS COMPACTOS ----------
    kpis = html.Div([
        html.Div([
            html.H4("Conforme", style={'color':'#27ae60','margin':'0', 'fontSize': '12px'}),
            html.H2(f"{conforme}", style={'color':'#27ae60','margin':'0', 'fontSize': '22px'}),
            html.P(f"{(conforme/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#27ae60', 'fontSize': '10px'})
        ], style={'borderLeft':'4px solid #27ae60','borderRadius':'3px','padding':'10px','margin':'5px','flex':'1',
                  'backgroundColor':'#eafaf1','textAlign':'center','boxShadow':'0 1px 2px rgba(0,0,0,0.05)',
                  'minWidth': '120px', 'maxWidth': '140px'}),

        html.Div([
            html.H4("Conforme Parcial", style={'color':'#f39c12','margin':'0', 'fontSize': '12px'}),
            html.H2(f"{parcial}", style={'color':'#f39c12','margin':'0', 'fontSize': '22px'}),
            html.P(f"{(parcial/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#f39c12', 'fontSize': '10px'})
        ], style={'borderLeft':'4px solid #f39c12','borderRadius':'3px','padding':'10px','margin':'5px','flex':'1',
                  'backgroundColor':'#fff8e1','textAlign':'center','boxShadow':'0 1px 2px rgba(0,0,0,0.05)',
                  'minWidth': '120px', 'maxWidth': '140px'}),

        html.Div([
            html.H4("Não Conforme", style={'color':'#e74c3c','margin':'0', 'fontSize': '12px'}),
            html.H2(f"{nao}", style={'color':'#e74c3c','margin':'0', 'fontSize': '22px'}),
            html.P(f"{(nao/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#e74c3c', 'fontSize': '10px'})
        ], style={'borderLeft':'4px solid #e74c3c','borderRadius':'3px','padding':'10px','margin':'5px','flex':'1',
                  'backgroundColor':'#fdecea','textAlign':'center','boxShadow':'0 1px 2px rgba(0,0,0,0.05)',
                  'minWidth': '120px', 'maxWidth': '140px'})
    ], style={'display':'flex','justifyContent':'center','flexWrap':'wrap','marginBottom':'15px', 'gap': '3px'})

    # ---------- Tabela de NÃO CONFORMES COM COMPARAÇÃO DE PRAZOS ----------
    df_nao_conforme = df[df['Status']=='Não Conforme']
    
    # Inicializar contadores de prazo
    dentro_prazo = 0
    fora_prazo = 0
    nao_concluido = 0
    
    if len(df_nao_conforme) > 0:
        # Fazer uma cópia para não modificar o original
        df_nao_conforme_display = df_nao_conforme.copy()
        
        # Procurar colunas de prazo e data de finalização
        colunas_disponiveis = df_nao_conforme_display.columns.tolist()
        colunas_lower = [str(col).lower() for col in colunas_disponiveis]
        
        # Procurar coluna de prazo
        coluna_prazo = None
        prazo_keywords = ['prazo', 'prazo_final', 'data_prazo', 'data_limite', 'limite']
        for keyword in prazo_keywords:
            for idx, col_lower in enumerate(colunas_lower):
                if keyword in col_lower:
                    coluna_prazo = colunas_disponiveis[idx]
                    break
            if coluna_prazo:
                break
        
        # Procurar coluna de data de finalização
        coluna_finalizacao = None
        finalizacao_keywords = ['data_finalizacao', 'data_conclusao', 'finalizacao', 'conclusao', 
                               'data_encerramento', 'data_termino']
        for keyword in finalizacao_keywords:
            for idx, col_lower in enumerate(colunas_lower):
                if keyword in col_lower:
                    coluna_finalizacao = colunas_disponiveis[idx]
                    break
            if coluna_finalizacao:
                break
        
        # Se encontrou ambas as colunas, calcular status do prazo
        if coluna_prazo and coluna_finalizacao:
            print(f"✅ Encontradas colunas de prazo: '{coluna_prazo}' e finalização: '{coluna_finalizacao}'")
            
            # Formatar datas
            df_nao_conforme_display['Prazo_Formatado'] = df_nao_conforme_display[coluna_prazo].apply(formatar_data)
            df_nao_conforme_display['Finalizacao_Formatada'] = df_nao_conforme_display[coluna_finalizacao].apply(formatar_data)
            
            # Calcular status do prazo
            df_nao_conforme_display['Status_Prazo'] = df_nao_conforme_display.apply(
                lambda row: calcular_status_prazo(row[coluna_prazo], row[coluna_finalizacao]), 
                axis=1
            )
            
            # Contar status dos prazos
            status_prazos = df_nao_conforme_display['Status_Prazo'].value_counts()
            dentro_prazo = status_prazos.get('Concluído no Prazo', 0)
            fora_prazo = status_prazos.get('Concluído Fora do Prazo', 0)
            nao_concluido = status_prazos.get('Não Concluído', 0)
            
            print(f"📊 STATUS DOS PRAZOS:")
            print(f"  Dentro do prazo: {dentro_prazo}")
            print(f"  Fora do prazo: {fora_prazo}")
            print(f"  Não concluído: {nao_concluido}")
            
            # Reordenar colunas para melhor visualização
            colunas_ordenadas = ['Unidade', 'Status', 'Status_Prazo', 'Prazo_Formatado', 'Finalizacao_Formatada']
            colunas_restantes = [col for col in df_nao_conforme_display.columns 
                                if col not in colunas_ordenadas + [coluna_prazo, coluna_finalizacao]]
            
            colunas_finais = colunas_ordenadas + colunas_restantes
            df_nao_conforme_display = df_nao_conforme_display[colunas_finais]
            
            # Renomear colunas para exibição
            df_nao_conforme_display = df_nao_conforme_display.rename(columns={
                'Prazo_Formatado': 'Prazo',
                'Finalizacao_Formatada': 'Data Finalização',
                'Status_Prazo': 'Status do Prazo'
            })
            
            # Criar tabela COMPACTA com cores condicionais
            tabela_nao_conforme = dash_table.DataTable(
                columns=[{"name": col, "id": col} for col in df_nao_conforme_display.columns],
                data=df_nao_conforme_display.to_dict('records'),
                page_size=5,  # Menos linhas por página
                style_table={'overflowX':'auto', 'fontSize': '10px', 'marginTop': '5px'},
                style_header={
                    'backgroundColor': '#c0392b',
                    'color': 'white',
                    'fontWeight': 'bold',
                    'textAlign':'center',
                    'fontSize': '10px',
                    'padding': '4px 5px',
                    'minHeight': '30px',
                    'height': '30px'
                },
                style_cell={
                    'textAlign': 'center',
                    'padding': '3px 4px',
                    'whiteSpace':'normal',
                    'height':'auto',
                    'fontSize': '9px',
                    'minWidth': '40px',
                    'maxWidth': '120px',
                    'overflow': 'hidden',
                    'textOverflow': 'ellipsis'
                },
                style_data_conditional=[
                    {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9e6e6'},
                    {'if': {'row_index': 'even'}, 'backgroundColor': '#fdecea'},
                    {
                        'if': {
                            'filter_query': '{Status do Prazo} = "Concluído no Prazo"',
                        },
                        'backgroundColor': '#d4edda',
                        'color': '#155724',
                        'fontWeight': 'bold'
                    },
                    {
                        'if': {
                            'filter_query': '{Status do Prazo} = "Concluído Fora do Prazo"',
                        },
                        'backgroundColor': '#f8d7da',
                        'color': '#721c24',
                        'fontWeight': 'bold'
                    },
                    {
                        'if': {
                            'filter_query': '{Status do Prazo} = "Não Concluído"',
                        },
                        'backgroundColor': '#fff3cd',
                        'color': '#856404',
                        'fontWeight': 'bold'
                    }
                ]
            )
            
        else:
            # Se não encontrou as colunas, mostrar tabela normal COMPACTA
            print(f"⚠️ Não encontrou colunas de prazo/finalização. Colunas disponíveis: {colunas_disponiveis}")
            
            # Remover colunas desnecessárias
            colunas_para_remover = ['Ano', 'Mes', 'Mes_Ano']
            for col in colunas_para_remover:
                if col in df_nao_conforme_display.columns:
                    df_nao_conforme_display = df_nao_conforme_display.drop(columns=[col])
            
            # Formatar datas se houver
            colunas_data = [col for col in df_nao_conforme_display.columns 
                           if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
            
            for coluna_data in colunas_data:
                if coluna_data in df_nao_conforme_display.columns:
                    df_nao_conforme_display[coluna_data] = df_nao_conforme_display[coluna_data].apply(formatar_data)
            
            # Limitar número de colunas para visualização
            if len(df_nao_conforme_display.columns) > 8:
                # Manter apenas as colunas mais importantes
                colunas_importantes = ['Unidade', 'Status', 'Data', 'Descricao']
                colunas_selecionadas = [col for col in colunas_importantes if col in df_nao_conforme_display.columns]
                colunas_adicionais = [col for col in df_nao_conforme_display.columns if col not in colunas_importantes][:4]
                df_nao_conforme_display = df_nao_conforme_display[colunas_selecionadas + colunas_adicionais]
            
            tabela_nao_conforme = dash_table.DataTable(
                columns=[{"name": col, "id": col} for col in df_nao_conforme_display.columns],
                data=df_nao_conforme_display.to_dict('records'),
                page_size=5,
                style_table={'overflowX':'auto', 'fontSize': '10px', 'marginTop': '5px'},
                style_header={
                    'backgroundColor': '#c0392b',
                    'color': 'white',
                    'fontWeight': 'bold',
                    'textAlign':'center',
                    'fontSize': '10px',
                    'padding': '4px 5px',
                    'minHeight': '30px',
                    'height': '30px'
                },
                style_cell={
                    'textAlign': 'center',
                    'padding': '3px 4px',
                    'whiteSpace':'normal',
                    'height':'auto',
                    'fontSize': '9px',
                    'minWidth': '40px',
                    'maxWidth': '120px',
                    'overflow': 'hidden',
                    'textOverflow': 'ellipsis'
                },
                style_data_conditional=[
                    {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9e6e6'},
                    {'if': {'row_index': 'even'}, 'backgroundColor': '#fdecea'}
                ]
            )
    else:
        tabela_nao_conforme = html.Div([
            html.P("✅ Nenhum item não conforme encontrado com os filtros atuais.", 
                   style={'textAlign': 'center', 'padding': '10px', 'color': '#27ae60', 'fontSize': '11px'})
        ])
    
    tabela_titulo = html.H3(f"❌ Itens Não Conformes ({len(df_nao_conforme)} itens)", 
                           style={'marginTop': '15px', 'marginBottom': '8px', 'color': '#c0392b', 'fontSize': '14px'})
    
    # ---------- KPIs de PRAZOS dos Itens Não Conformes COMPACTOS ----------
    if len(df_nao_conforme) > 0 and (coluna_prazo and coluna_finalizacao):
        kpis_prazos = html.Div([
            html.Div([
                html.H4("Dentro Prazo", style={'color':'#27ae60','margin':'0', 'fontSize': '10px'}),
                html.H2(f"{dentro_prazo}", style={'color':'#27ae60','margin':'0', 'fontSize': '18px'}),
                html.P(f"{(dentro_prazo/len(df_nao_conforme)*100 if len(df_nao_conforme)>0 else 0):.1f}%", 
                       style={'margin':'0','color':'#27ae60', 'fontSize': '9px'})
            ], style={'borderLeft':'3px solid #27ae60','borderRadius':'2px','padding':'8px','margin':'3px','flex':'1',
                      'backgroundColor':'#d4edda','textAlign':'center','boxShadow':'0 1px 2px rgba(0,0,0,0.05)',
                      'minWidth': '90px', 'maxWidth': '100px'}),

            html.Div([
                html.H4("Fora Prazo", style={'color':'#e74c3c','margin':'0', 'fontSize': '10px'}),
                html.H2(f"{fora_prazo}", style={'color':'#e74c3c','margin':'0', 'fontSize': '18px'}),
                html.P(f"{(fora_prazo/len(df_nao_conforme)*100 if len(df_nao_conforme)>0 else 0):.1f}%", 
                       style={'margin':'0','color':'#e74c3c', 'fontSize': '9px'})
            ], style={'borderLeft':'3px solid #e74c3c','borderRadius':'2px','padding':'8px','margin':'3px','flex':'1',
                      'backgroundColor':'#f8d7da','textAlign':'center','boxShadow':'0 1px 2px rgba(0,0,0,0.05)',
                      'minWidth': '90px', 'maxWidth': '100px'}),

            html.Div([
                html.H4("Não Concluído", style={'color':'#f39c12','margin':'0', 'fontSize': '10px'}),
                html.H2(f"{nao_concluido}", style={'color':'#f39c12','margin':'0', 'fontSize': '18px'}),
                html.P(f"{(nao_concluido/len(df_nao_conforme)*100 if len(df_nao_conforme)>0 else 0):.1f}%", 
                       style={'margin':'0','color':'#f39c12', 'fontSize': '9px'})
            ], style={'borderLeft':'3px solid #f39c12','borderRadius':'2px','padding':'8px','margin':'3px','flex':'1',
                      'backgroundColor':'#fff3cd','textAlign':'center','boxShadow':'0 1px 2px rgba(0,0,0,0.05)',
                      'minWidth': '90px', 'maxWidth': '100px'})
        ], style={'display':'flex','justifyContent':'center','flexWrap':'wrap','marginBottom':'10px', 'gap': '3px'})
        
        # Adicionar legenda para KPIs de prazo COMPACTA
        legenda_prazo = html.Div([
            html.P("📊 Status dos Prazos:", 
                   style={'fontWeight':'bold','marginBottom':'2px', 'fontSize': '10px', 'textAlign': 'center'}),
            html.Div([
                html.Span("🟢 ", style={'color':'#27ae60','marginRight':'2px', 'fontSize': '9px'}),
                html.Span("Dentro Prazo", style={'marginRight':'8px','color':'#27ae60', 'fontSize': '9px'}),
                html.Span("🔴 ", style={'color':'#e74c3c','marginRight':'2px', 'fontSize': '9px'}),
                html.Span("Fora Prazo", style={'marginRight':'8px','color':'#e74c3c', 'fontSize': '9px'}),
                html.Span("🟡 ", style={'color':'#f39c12','marginRight':'2px', 'fontSize': '9px'}),
                html.Span("Não Concluído", style={'color':'#f39c12', 'fontSize': '9px'})
            ], style={'backgroundColor':'#f8f9fa','padding':'4px 6px','borderRadius':'2px','marginBottom':'6px', 
                      'fontSize': '9px', 'textAlign': 'center'})
        ], style={'marginBottom':'10px', 'textAlign': 'center'})
    else:
        kpis_prazos = html.Div()
        legenda_prazo = html.Div()

    # ---------- Matriz de Risco (APENAS ANO) ----------
    abas_extra = []

    if df_risco is not None and len(df_risco) > 0:
        print(f"\n📋 PROCESSANDO MATRIZ DE RISCO:")
        print(f"  Total de registros: {len(df_risco)}")
        
        df_risco_filtrado = df_risco.copy()
        
        # Aplicar filtros para matriz de risco
        if 'Ano' in df_risco_filtrado.columns:
            df_risco_filtrado['Ano'] = pd.to_numeric(df_risco_filtrado['Ano'], errors='coerce')
        
        # Aplicar filtro de ano se não for 'todos'
        if ano != 'todos' and 'Ano' in df_risco_filtrado.columns:
            try:
                ano_int = int(ano)
                df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Ano'] == ano_int]
                print(f"  ✅ Filtro ANO aplicado para matriz de risco: {ano_int}")
            except:
                pass
        
        # Aplicar filtro de unidade se não for 'todas'
        if unidade != 'todas' and 'Unidade' in df_risco_filtrado.columns:
            df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade]
            print(f"  ✅ Filtro UNIDADE aplicado para matriz de risco: '{unidade}'")
        
        # NÃO aplicar filtro de mês para a matriz de risco (mostrar ano completo)

        print(f"\n📋 Matriz após filtros (ano completo): {len(df_risco_filtrado)} registros")
        
        if len(df_risco_filtrado) > 0:
            # Determinar qual ano usar para a matriz
            if ano != 'todos':
                ano_matriz = int(ano)
            else:
                # Se 'todos', usar o primeiro ano disponível
                anos_disponiveis = sorted(df_risco_filtrado['Ano'].dropna().unique())
                if len(anos_disponiveis) > 0:
                    ano_matriz = int(anos_disponiveis[0])
                else:
                    ano_matriz = datetime.now().year
            
            # Criar matriz de risco anual
            matriz_risco = criar_matriz_risco_anual(df_risco_filtrado, ano_matriz)
            abas_extra.append(matriz_risco)
        else:
            abas_extra.append(html.Div([
                html.H3("📋 Matriz Auditoria Risco", style={'fontSize': '14px'}),
                html.P("Nenhum dado encontrado para o ano selecionado.", 
                       style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '15px', 'fontSize': '11px'})
            ], style={'marginTop':'15px'}))
    else:
        abas_extra.append(html.Div([
            html.H3("📋 Matriz Auditoria Risco", style={'fontSize': '14px'}),
            html.P("Não há dados de risco disponíveis.", 
                   style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '15px', 'fontSize': '11px'})
        ], style={'marginTop':'15px'}))

    # ---------- Melhorias e Políticas COMPACTAS ----------
    if df_melhorias is not None and len(df_melhorias) > 0:
        colunas_data_melhorias = [col for col in df_melhorias.columns 
                                 if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
        
        df_melhorias_display = df_melhorias.copy()
        for coluna_data in colunas_data_melhorias:
            if coluna_data in df_melhorias_display.columns:
                df_melhorias_display[coluna_data] = df_melhorias_display[coluna_data].apply(formatar_data)
        
        # Limitar número de colunas para visualização
        if len(df_melhorias_display.columns) > 6:
            # Manter apenas as colunas mais importantes
            colunas_importantes = ['Unidade', 'Status', 'Data', 'Descricao']
            colunas_selecionadas = [col for col in colunas_importantes if col in df_melhorias_display.columns]
            colunas_adicionais = [col for col in df_melhorias_display.columns if col not in colunas_importantes][:2]
            df_melhorias_display = df_melhorias_display[colunas_selecionadas + colunas_adicionais]
        
        tabela_melhorias = dash_table.DataTable(
            columns=[{"name": col, "id": col} for col in df_melhorias_display.columns],
            data=df_melhorias_display.to_dict('records'),
            page_size=5,
            style_table={'overflowX':'auto','marginTop':'5px', 'fontSize': '10px'},
            style_header={
                'backgroundColor': '#34495e',
                'color': 'white',
                'fontWeight': 'bold',
                'textAlign':'center',
                'fontSize': '10px',
                'padding': '4px 5px',
                'minHeight': '30px',
                'height': '30px'
            },
            style_cell={
                'textAlign': 'center',
                'padding': '3px 4px',
                'whiteSpace':'normal',
                'height':'auto',
                'fontSize': '9px',
                'minWidth': '40px',
                'maxWidth': '120px'
            },
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#ecf0f1'},
                {'if': {'row_index': 'even'}, 'backgroundColor': 'white'}
            ]
        )
        abas_extra.append(html.Div([
            html.H3(f"📈 Melhorias ({len(df_melhorias)} registros)", style={'fontSize': '14px', 'marginBottom': '5px'}),
            html.P("Todos os registros de melhorias (filtros não aplicados)", 
                   style={'color': '#7f8c8d', 'marginBottom': '5px', 'fontSize': '10px'}),
            tabela_melhorias
        ], style={'marginTop':'15px', 'padding': '5px'}))

    if df_politicas is not None and len(df_politicas) > 0:
        colunas_data_politicas = [col for col in df_politicas.columns 
                                 if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
        
        df_politicas_display = df_politicas.copy()
        for coluna_data in colunas_data_politicas:
            if coluna_data in df_politicas_display.columns:
                df_politicas_display[coluna_data] = df_politicas_display[coluna_data].apply(formatar_data)
        
        # Limitar número de colunas para visualização
        if len(df_politicas_display.columns) > 6:
            # Manter apenas as colunas mais importantes
            colunas_importantes = ['Unidade', 'Status', 'Data', 'Descricao']
            colunas_selecionadas = [col for col in colunas_importantes if col in df_politicas_display.columns]
            colunas_adicionais = [col for col in df_politicas_display.columns if col not in colunas_importantes][:2]
            df_politicas_display = df_politicas_display[colunas_selecionadas + colunas_adicionais]
        
        tabela_politicas = dash_table.DataTable(
            columns=[{"name": col, "id": col} for col in df_politicas_display.columns],
            data=df_politicas_display.to_dict('records'),
            page_size=5,
            style_table={'overflowX':'auto','marginTop':'5px', 'fontSize': '10px'},
            style_header={
                'backgroundColor': '#34495e',
                'color': 'white',
                'fontWeight': 'bold',
                'textAlign':'center',
                'fontSize': '10px',
                'padding': '4px 5px',
                'minHeight': '30px',
                'height': '30px'
            },
            style_cell={
                'textAlign': 'center',
                'padding': '3px 4px',
                'whiteSpace':'normal',
                'height':'auto',
                'fontSize': '9px',
                'minWidth': '40px',
                'maxWidth': '120px'
            },
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#ecf0f1'},
                {'if': {'row_index': 'even'}, 'backgroundColor': 'white'}
            ]
        )
        abas_extra.append(html.Div([
            html.H3(f"📑 Políticas ({len(df_politicas)} registros)", style={'fontSize': '14px', 'marginBottom': '5px'}),
            html.P("Todos os registros de políticas (filtros não aplicados)", 
                   style={'color': '#7f8c8d', 'marginBottom': '5px', 'fontSize': '10px'}),
            tabela_politicas
        ], style={'marginTop':'15px', 'padding': '5px'}))

    # ---------- Layout Final COMPACTO ----------
    return html.Div([
        html.Div([
            html.H4(f"📊 Resumo - {len(df)} itens auditados", 
                    style={'textAlign': 'center', 'color': '#2c3e50', 'marginBottom': '10px', 'fontSize': '16px'})
        ]),
        kpis,
        tabela_titulo,
        kpis_prazos,
        legenda_prazo,
        tabela_nao_conforme,
        *abas_extra
    ], style={'fontSize': '11px'})

# ========== EXECUÇÃO DO APP ==========
if __name__ == '__main__':
    print("\n" + "="*50)
    print("🌐 DASHBOARD RODANDO: http://localhost:8050")
    print("📊 DASHBOARD COMPACTO OTIMIZADO:")
    print("  - ✅ Matriz de risco com TODAS as siglas (altura 400px)")
    print("  - ✅ Fontes reduzidas (6px-9px) para mais conteúdo")
    print("  - ✅ Tabelas compactas (5 linhas por página)")
    print("  - ✅ KPIs menores e mais compactos")
    print("  - ✅ Tooltips com detalhes dos relatórios")
    print("  - ✅ SEM rolagem vertical na matriz")
    print("="*50)
    app.run(debug=True, host='0.0.0.0', port=8050)

# ========== SERVER PARA O RENDER ==========
server = app.server

