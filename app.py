from dash import Dash, html, dcc, Input, Output, dash_table
import pandas as pd
import plotly.express as px
import os
import unicodedata
from datetime import datetime
import dash_auth  # Importação para autenticação
import re  # Adicionado para processamento de datas

print("🚀 Iniciando Dashboard de Auditoria...")

# ========== CONFIGURAÇÃO DE AUTENTICAÇÃO ==========
# Defina os usuários e senhas aqui
# Formato: 'usuário': 'senha'
USUARIOS_VALIDOS = {
    'admin': 'wne@2026',    # ALTERE PARA SUA SENHA SEGURA
    'diretoria': 'lagoa@2026'     # ALTERE PARA SUA SENHA SEGURA
}

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
        prazo = pd.to_datetime(prazo_str, errors='coerce', dayfirst=True)
        finalizacao = pd.to_datetime(finalizacao_str, errors='coerce', dayfirst=True)
        
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
    """Formata a data para DD/MM/YYYY (formato brasileiro)"""
    try:
        if pd.isna(data_str) or str(data_str).strip() in ['', 'NaT', 'None']:
            return ""
        
        # Se já for um objeto datetime, formata diretamente
        if isinstance(data_str, (pd.Timestamp, datetime)):
            return data_str.strftime('%d/%m/%Y')
        
        # Se for string, tenta converter para datetime primeiro
        # dayfirst=True ajuda no formato brasileiro (dd/mm/aaaa)
        data = pd.to_datetime(data_str, errors='coerce', dayfirst=True)
        if pd.isna(data):
            return str(data_str)  # Retorna original se não puder converter
        
        return data.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"Erro ao formatar data '{data_str}': {e}")
        return str(data_str)

def carregar_dados_da_planilha():
    planilha_path = 'base_auditoria.xlsx'
    if not os.path.exists(planilha_path):
        print("❌ Planilha não encontrada:", planilha_path)
        return None, None, None, None

    try:
        print(f"📁 Carregando dados da planilha: {planilha_path}")

        # Leitura das planilhas
        df_checklist = pd.read_excel(planilha_path, sheet_name='Checklist_Unidades', 
                                     engine='openpyxl', dtype=str)
        df_politicas = pd.read_excel(planilha_path, sheet_name='Politicas', 
                                     engine='openpyxl', dtype=str)
        df_risco = pd.read_excel(planilha_path, sheet_name='Auditoria_Risco', 
                                 engine='openpyxl', dtype=str)
        df_melhorias = pd.read_excel(planilha_path, sheet_name='Melhorias_Logistica', 
                                     engine='openpyxl', dtype=str)

        print("✅ Leitura inicial da planilha concluída. Processando dados...")

        for i, df in enumerate([df_checklist, df_politicas, df_risco, df_melhorias]):
            if df is not None:
                df = normalize_df_columns(df)
                
                # CORREÇÃO ESPECÍFICA PARA CADA ABA
                if i == 0:  # df_checklist - CORREÇÃO CRÍTICA AQUI
                    print("📋 Processando CHECKLIST...")
                    
                    # Normalizar Status
                    if 'Status' in df.columns:
                        df['Status'] = df['Status'].astype(str).str.strip()
                        print(f"  Status únicos antes: {df['Status'].unique()[:10]}")
                        df['Status'] = df['Status'].apply(canonical_status)
                        print(f"  Status únicos depois: {df['Status'].unique()[:10]}")
                    
                    # CORREÇÃO CRÍTICA: Processar datas corretamente
                    if 'Data' in df.columns:
                        print(f"  Processando coluna Data...")
                        
                        # Converter a coluna Data para datetime CORRETAMENTE
                        # Primeiro, tentar converter com dayfirst=True (formato brasileiro)
                        df['Data_DT'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
                        
                        # Verificar se alguma conversão falhou
                        falhas = df['Data_DT'].isna().sum()
                        if falhas > 0:
                            print(f"  ⚠️ {falhas} datas não puderam ser convertidas com dayfirst=True")
                            # Tentar sem dayfirst
                            df.loc[df['Data_DT'].isna(), 'Data_DT'] = pd.to_datetime(
                                df.loc[df['Data_DT'].isna(), 'Data'], errors='coerce'
                            )
                        
                        # Extrair Ano e Mes como INTEIROS
                        df['Ano'] = df['Data_DT'].dt.year
                        df['Mes'] = df['Data_DT'].dt.month
                        
                        # Converter para inteiros explicitamente
                        df['Ano'] = df['Ano'].fillna(0).astype(int)
                        df['Mes'] = df['Mes'].fillna(0).astype(int)
                        
                        # Substituir 0 por NaN
                        df['Ano'] = df['Ano'].replace(0, pd.NA)
                        df['Mes'] = df['Mes'].replace(0, pd.NA)
                        
                        # DEBUG: Mostrar distribuição de anos
                        anos_distribuicao = df['Ano'].value_counts().sort_index()
                        print(f"  📅 DISTRIBUIÇÃO DE ANOS:")
                        for ano, contagem in anos_distribuicao.items():
                            print(f"     {ano}: {contagem} registros")
                        
                        print(f"  Total de registros: {len(df)}")
                        print(f"  Ano (int) únicos: {df['Ano'].dropna().unique()}")
                        print(f"  Mês (int) únicos: {df['Mes'].dropna().unique()}")
                        
                        # Formatar data para exibição
                        df['Data'] = df['Data_DT'].apply(
                            lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else ''
                        )
                        
                        # Remover coluna temporária
                        df = df.drop(columns=['Data_DT'])
                    
                    # DEBUG: Contar Conformes por ano
                    if 'Status' in df.columns and 'Ano' in df.columns:
                        print(f"  📊 CONFORMES POR ANO:")
                        for ano in sorted(df['Ano'].dropna().unique()):
                            conformes_ano = len(df[(df['Ano'] == ano) & (df['Status'] == 'Conforme')])
                            total_ano = len(df[df['Ano'] == ano])
                            print(f"     Ano {ano}: {conformes_ano} conformes de {total_ano} total")
                
                elif i == 1:  # df_politicas
                    print("📑 Processando POLÍTICAS...")
                    if 'Status' in df.columns:
                        df['Status'] = df['Status'].apply(canonical_status)
                
                elif i == 2:  # df_risco
                    print("🔄 Processando dados de RISCO...")
                    
                    # Normalizar Status
                    if 'Status' in df.columns:
                        df['Status'] = df['Status'].apply(canonical_status)
                    
                    # Processar datas para a matriz - VERSÃO CORRIGIDA
                    if 'Data' in df.columns:
                        print(f"  🔍 Processando datas de RISCO...")
                        print(f"  Amostra de datas brutas: {df['Data'].head(10).tolist()}")
                        
                        # Converter para datetime com tratamento robusto
                        df['Data_DT'] = pd.to_datetime(
                            df['Data'], 
                            errors='coerce', 
                            dayfirst=True,  # Formato brasileiro dd/mm/aaaa
                            exact=False     # Mais flexível
                        )
                        
                        # DEBUG: Verificar resultados da conversão
                        total_dates = len(df['Data'])
                        success_dates = df['Data_DT'].notna().sum()
                        failed_dates = total_dates - success_dates
                        
                        print(f"  Total de datas: {total_dates}")
                        print(f"  Conversões bem-sucedidas: {success_dates}")
                        print(f"  Conversões falhadas: {failed_dates}")
                        
                        if failed_dates > 0:
                            # Mostrar exemplos de datas que falharam
                            failed_samples = df[df['Data_DT'].isna()]['Data'].head(5).tolist()
                            print(f"  ⚠️ Exemplos de datas que falharam: {failed_samples}")
                        
                        # Criar colunas de mês e ano a partir da data convertida
                        df['Mes'] = df['Data_DT'].dt.month
                        df['Ano'] = df['Data_DT'].dt.year
                        
                        # Para datas que não puderam ser convertidas, tentar extrair de outras formas
                        if failed_dates > 0:
                            # Tentar extrair mês e ano do texto da data
                            for idx in df[df['Data_DT'].isna()].index:
                                data_str = str(df.at[idx, 'Data'])
                                if data_str and data_str != 'nan' and data_str != 'NaT':
                                    # Tentar encontrar padrões de data no texto
                                    # Padrão dd/mm/aaaa ou dd-mm-aaaa
                                    padrao_data = r'(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})'
                                    match = re.search(padrao_data, data_str)
                                    if match:
                                        dia, mes, ano = match.groups()
                                        try:
                                            mes_int = int(mes)
                                            ano_int = int(ano) if len(ano) == 4 else 2000 + int(ano)
                                            df.at[idx, 'Mes'] = mes_int
                                            df.at[idx, 'Ano'] = ano_int
                                            print(f"    ✓ Extraído: {data_str} -> Mês: {mes_int}, Ano: {ano_int}")
                                        except:
                                            pass
                        
                        # Converter para inteiros
                        df['Mes'] = df['Mes'].fillna(0).astype(int)
                        df['Ano'] = df['Ano'].fillna(0).astype(int)
                        df['Mes'] = df['Mes'].replace(0, pd.NA)
                        df['Ano'] = df['Ano'].replace(0, pd.NA)
                        
                        # DEBUG: Mostrar distribuição de anos e meses
                        print(f"  📅 DISTRIBUIÇÃO ANOS (RISCO):")
                        if 'Ano' in df.columns:
                            anos_dist = df['Ano'].value_counts().sort_index()
                            for ano, contagem in anos_dist.items():
                                print(f"     Ano {ano}: {contagem} registros")
                        
                        print(f"  📅 DISTRIBUIÇÃO MESES (RISCO):")
                        if 'Mes' in df.columns:
                            meses_dist = df['Mes'].value_counts().sort_index()
                            for mes_val, contagem in meses_dist.items():
                                print(f"     Mês {mes_val}: {contagem} registros")
                        
                        # Formatar data para exibição
                        df['Data'] = df['Data_DT'].apply(
                            lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else ''
                        )
                        
                        # Remover coluna temporária
                        df = df.drop(columns=['Data_DT'])
                
                elif i == 3:  # df_melhorias
                    print("📈 Processando MELHORIAS...")
                    if 'Status' in df.columns:
                        df['Status'] = df['Status'].apply(canonical_status)
                
                if i == 0: df_checklist = df
                elif i == 1: df_politicas = df
                elif i == 2: df_risco = df
                elif i == 3: df_melhorias = df

        print("✅ Dados carregados da planilha com sucesso!")
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
    # CORREÇÃO: Remover .0 e converter para inteiro
    anos = sorted(df_checklist['Ano'].dropna().unique(), reverse=True)
    # Converter para inteiro e remover duplicados
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
    if df_checklist is None or 'Mes' not in df_checklist.columns:
        return []
    if ano_selecionado == 'todos':
        meses = sorted(df_checklist['Mes'].dropna().unique())
    else:
        try:
            ano_selecionado = int(ano_selecionado)
            df_filtrado = df_checklist[df_checklist['Ano'] == ano_selecionado]
            meses = sorted(df_filtrado['Mes'].dropna().unique())
        except:
            meses = []
    
    nomes_meses = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
                   7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}
    
    return [{'label': f'{nomes_meses.get(int(m), str(m))}', 'value': int(m)} for m in meses]

# ========== CARREGAR DADOS ==========
df_checklist, df_politicas, df_risco, df_melhorias = carregar_dados_da_planilha()
if df_checklist is None:
    app = Dash(__name__)
    server = app.server
    app.layout = html.Div([html.H1("❌ Planilha não encontrada")])
    if __name__ == '__main__':
        app.run(debug=True, port=8050)
    exit()

# Verificação adicional da matriz de risco
print("="*50)
print("VERIFICAÇÃO MATRIZ DE RISCO:")
print("="*50)

if df_risco is not None:
    print(f"Total de registros: {len(df_risco)}")
    print(f"Colunas disponíveis: {df_risco.columns.tolist()}")
    
    if 'Data' in df_risco.columns:
        print(f"\n📅 10 primeiras datas na matriz de risco:")
        print(df_risco['Data'].head(10).tolist())
    
    if 'Mes' in df_risco.columns and 'Ano' in df_risco.columns:
        print(f"\n📊 Combinações únicas de Mês/Ano:")
        combinacoes = df_risco[['Mes', 'Ano']].drop_duplicates()
        print(combinacoes.head(20))
        
        # Criar Mes_Ano para verificação
        if 'Mes_Ano' not in df_risco.columns:
            df_risco['Mes_Ano'] = df_risco.apply(
                lambda row: f"{int(row['Mes']):02d}/{int(row['Ano'])}" 
                if pd.notna(row['Mes']) and pd.notna(row['Ano']) 
                else "Sem Data", 
                axis=1
            )
        
        print(f"\n📈 Valores únicos de Mes_Ano:")
        print(df_risco['Mes_Ano'].value_counts().sort_index())

anos_disponiveis = obter_anos_disponiveis(df_checklist)
print(f"DEBUG: Anos disponíveis no filtro: {anos_disponiveis}")

# ========== APP DASH ==========
app = Dash(__name__)

# ========== APLICAR AUTENTICAÇÃO ==========
# Esta linha DEVE vir DEPOIS de criar a app e ANTES do layout
auth = dash_auth.BasicAuth(app, USUARIOS_VALIDOS)

# ========== LAYOUT DO DASHBOARD ==========
app.layout = html.Div([
    html.Div([html.H1("📊 DASHBOARD DE AUDITORIA", 
                      style={'textAlign':'center', 'marginBottom':'20px'})]),
    html.Div([
        html.Div([
            html.Label("Ano:"),
            dcc.Dropdown(
                id='filtro-ano',
                options=[{'label':'Todos','value':'todos'}]+
                       [{'label':str(a),'value':a} for a in anos_disponiveis],
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
                        [{'label':str(u),'value':str(u)} for u in sorted(df_checklist['Unidade'].dropna().unique())],
                value='todas'
            )
        ], style={'width':'250px'})
    ], style={'display':'flex','justifyContent':'center','marginBottom':'30px','flexWrap':'wrap'}),
    html.Div(id='conteudo-principal', style={'padding':'20px'})
])

# ========== CALLBACKS ==========
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
    
    # CORREÇÃO CRÍTICA: Converter colunas para tipos consistentes antes de filtrar
    print(f"🔍 DEBUG FILTROS: Ano='{ano}', Mês='{mes}', Unidade='{unidade}'")
    
    # Converter colunas numéricas para o tipo correto
    if 'Ano' in df.columns:
        df['Ano'] = pd.to_numeric(df['Ano'], errors='coerce')
        print(f"  Ano - Valores únicos: {df['Ano'].dropna().unique()}")
    
    if 'Mes' in df.columns:
        df['Mes'] = pd.to_numeric(df['Mes'], errors='coerce')
        print(f"  Mês - Valores únicos: {df['Mes'].dropna().unique()}")
    
    # Aplicar filtros com CONVERSÃO CORRETA
    total_antes = len(df)
    
    if ano != 'todos':
        try:
            ano_filtro = int(ano)
            df = df[df['Ano'] == ano_filtro]
            print(f"  ✅ Filtro ANO aplicado: {ano_filtro} | Registros: {len(df)}/{total_antes}")
        except Exception as e:
            print(f"  ❌ Erro ao filtrar por ano '{ano}': {e}")
            # Mostrar exemplos para debug
            if 'Ano' in df.columns:
                print(f"    Exemplos de valores na coluna Ano: {df['Ano'].head(10).tolist()}")
    
    if mes != 'todos':
        try:
            mes_filtro = int(mes)
            df = df[df['Mes'] == mes_filtro]
            print(f"  ✅ Filtro MÊS aplicado: {mes_filtro} | Registros: {len(df)}")
        except Exception as e:
            print(f"  ❌ Erro ao filtrar por mês '{mes}': {e}")
            if 'Mes' in df.columns:
                print(f"    Exemplos de valores na coluna Mes: {df['Mes'].head(10).tolist()}")
    
    if unidade != 'todas':
        try:
            # Converter unidade para string para comparação
            df['Unidade'] = df['Unidade'].astype(str).str.strip()
            df = df[df['Unidade'] == unidade.strip()]
            print(f"  ✅ Filtro UNIDADE aplicado: '{unidade}' | Registros: {len(df)}")
        except Exception as e:
            print(f"  ❌ Erro ao filtrar por unidade '{unidade}': {e}")
            if 'Unidade' in df.columns:
                print(f"    Exemplos de valores na coluna Unidade: {df['Unidade'].head(10).tolist()}")
    
    total = len(df)
    print(f"📊 TOTAL APÓS FILTROS: {total} registros")
    
    # ---------- Contagem correta dos status ----------
    if total > 0:
        # CORREÇÃO: Garantir que Status seja string e remover espaços
        df['Status'] = df['Status'].astype(str).str.strip()
        
        # Contagem DIRETA e precisa
        conforme = len(df[df['Status'].str.lower() == 'conforme'])
        parcial = len(df[df['Status'].str.lower().str.contains('parcial')])
        nao = len(df[df['Status'].str.lower().str.contains('não|nao')])
        
        # Debug detalhado
        print(f"🔢 CONTAGEM STATUS:")
        print(f"   Total registros: {total}")
        print(f"   Conforme: {conforme} (query: Status == 'Conforme')")
        print(f"   Parcial: {parcial} (query: Status contém 'parcial')")
        print(f"   Não Conforme: {nao} (query: Status contém 'não|nao')")
        
        # Verificar se há outros status
        outros_status = df[~df['Status'].str.lower().str.contains('conforme|parcial|não|nao')]['Status'].unique()
        if len(outros_status) > 0:
            print(f"   ⚠️ Outros status encontrados: {outros_status}")
            
        # Soma para verificação
        soma = conforme + parcial + nao
        if soma != total:
            print(f"   ⚠️ ATENÇÃO: Soma ({soma}) ≠ Total ({total})")
            print(f"   Diferença: {total - soma} registros")
            print(f"   Valores únicos de Status: {df['Status'].unique()}")
    else:
        conforme = 0
        parcial = 0
        nao = 0

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
            # Garantir que as colunas estejam formatadas (usando a função formatar_data)
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
        
        # Aplicar formatação de data em todas as colunas que parecem ser datas
        colunas_data = [col for col in df_nao_conforme_display.columns 
                       if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
        
        for coluna_data in colunas_data:
            if coluna_data in df_nao_conforme_display.columns:
                df_nao_conforme_display[coluna_data] = df_nao_conforme_display[coluna_data].apply(formatar_data)
        
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
        
        # CORREÇÃO: Converter colunas Ano e Mes para numérico
        if 'Ano' in df_risco_filtrado.columns:
            df_risco_filtrado['Ano'] = pd.to_numeric(df_risco_filtrado['Ano'], errors='coerce')
        if 'Mes' in df_risco_filtrado.columns:
            df_risco_filtrado['Mes'] = pd.to_numeric(df_risco_filtrado['Mes'], errors='coerce')
        
        # Aplicar filtros
        if ano != 'todos' and 'Ano' in df_risco_filtrado.columns:
            try:
                ano_int = int(ano)
                df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Ano'] == ano_int]
                print(f"  ✅ Matriz - Filtro ANO aplicado: {ano_int}")
            except:
                print(f"  ⚠️ Erro ao filtrar matriz por ano: {ano}")
        
        if mes != 'todos' and 'Mes' in df_risco_filtrado.columns:
            try:
                mes_int = int(mes)
                df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Mes'] == mes_int]
                print(f"  ✅ Matriz - Filtro MÊS aplicado: {mes_int}")
            except:
                print(f"  ⚠️ Erro ao filtrar matriz por mês: {mes}")
        
        if unidade != 'todas' and 'Unidade' in df_risco_filtrado.columns:
            df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade]
            print(f"  ✅ Matriz - Filtro UNIDADE aplicado: '{unidade}'")

        # CORREÇÃO: Criar Mes_Ano de forma robusta - VERIFICAR SE TEM DADOS
        print(f"📋 Matriz: {len(df_risco_filtrado)} registros após filtros")
        
        if len(df_risco_filtrado) > 0:
            # Garantir que temos Ano e Mes
            if 'Ano' not in df_risco_filtrado.columns or 'Mes' not in df_risco_filtrado.columns:
                print("⚠️ Matriz: Colunas Ano ou Mes não encontradas")
            
            # Criar Mes_Ano se não existir
            if 'Mes_Ano' not in df_risco_filtrado.columns:
                df_risco_filtrado['Mes_Ano'] = df_risco_filtrado.apply(
                    lambda row: f"{int(row['Mes']):02d}/{int(row['Ano'])}" 
                    if pd.notna(row['Mes']) and pd.notna(row['Ano']) and row['Mes'] != 0 and row['Ano'] != 0
                    else "Sem Data", 
                    axis=1
                )
            
            # Agrupar dados por Unidade e Mes_Ano
            unidades = sorted(df_risco_filtrado['Unidade'].dropna().unique())
            meses_anos = sorted(df_risco_filtrado['Mes_Ano'].dropna().unique())
            
            print(f"📊 Matriz: {len(unidades)} unidades, {len(meses_anos)} meses_anos")
            print(f"📅 Meses_Anos: {meses_anos}")
            
            # Criar estrutura de dados para a matriz
            matriz_data = []
            
            for unidade_nome in unidades:
                linha = {'Unidade': unidade_nome}
                df_unidade = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade_nome]
                
                for mes_ano in meses_anos:
                    if pd.isna(mes_ano) or mes_ano == "Sem Data":
                        continue
                    
                    df_mes = df_unidade[df_unidade['Mes_Ano'] == mes_ano]
                    
                    if len(df_mes) > 0:
                        # Criar lista de relatórios com cores
                        relatorios_html = []
                        for _, row in df_mes.iterrows():
                            relatorio = str(row['Relatorio']) if 'Relatorio' in row else "Sem Relatório"
                            status = str(row['Status']) if 'Status' in row else "Sem Status"
                            cores = get_status_color(status)
                            
                            relatorio_item = html.Span(
                                relatorio,
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
                if pd.isna(mes_ano) or mes_ano == "Sem Data":
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
                    if pd.isna(mes_ano) or mes_ano == "Sem Data":
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
                html.P(f"Total de registros: {len(df_risco_filtrado)} | Período: {ano if ano != 'todos' else 'Todos'} {f'Mês: {mes}' if mes != 'todos' else ''}", 
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
        # Aplicar formatação de data nas colunas de data
        colunas_data_melhorias = [col for col in df_melhorias.columns 
                                 if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
        
        df_melhorias_display = df_melhorias.copy()
        for coluna_data in colunas_data_melhorias:
            if coluna_data in df_melhorias_display.columns:
                df_melhorias_display[coluna_data] = df_melhorias_display[coluna_data].apply(formatar_data)
        
        tabela_melhorias = dash_table.DataTable(
            df_melhorias_display.to_dict('records'),
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
        # Aplicar formatação de data nas colunas de data
        colunas_data_politicas = [col for col in df_politicas.columns 
                                 if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
        
        df_politicas_display = df_politicas.copy()
        for coluna_data in colunas_data_politicas:
            if coluna_data in df_politicas_display.columns:
                df_politicas_display[coluna_data] = df_politicas_display[coluna_data].apply(formatar_data)
        
        tabela_politicas = dash_table.DataTable(
            df_politicas_display.to_dict('records'),
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

# ========== EXECUÇÃO DO APP ==========
if __name__ == '__main__':
    print("🌐 DASHBOARD RODANDO: http://localhost:8050")
    app.run(debug=True, host='0.0.0.0', port=8050)

# ========== SERVER PARA O RENDER ==========
server = app.server

