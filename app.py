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

        # Leitura das planilhas - IMPORTANTE: Não usar dtype=str para permitir conversão de tipos
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
                
                elif i == 2:  # df_risco - CORREÇÃO COMPLETA AQUI
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
                    
                    # 2. Encontrar e processar coluna de Data - CORREÇÃO CRÍTICA
                    coluna_data = None
                    for col in df.columns:
                        if col.lower() == 'data':
                            coluna_data = col
                            break
                    
                    if coluna_data:
                        print(f"\n  ✅ Coluna de Data encontrada: '{coluna_data}'")
                        print(f"  Tipo da coluna Data: {df[coluna_data].dtype}")
                        print(f"  Amostra de 10 datas brutas:")
                        datas_amostra = df[coluna_data].head(10).tolist()
                        for j, data in enumerate(datas_amostra):
                            print(f"    {j+1:2d}. '{data}' (tipo: {type(data)})")
                        
                        # CORREÇÃO: Converter datas para datetime de forma mais agressiva
                        print(f"\n  🔍 Convertendo datas para datetime...")
                        
                        # Primeiro, converter tudo para string para análise
                        df['Data_Str'] = df[coluna_data].astype(str)
                        
                        # Tentar extrair datas de diferentes formatos
                        def converter_data_agressiva(data_str):
                            if pd.isna(data_str) or data_str in ['nan', 'NaT', 'None', '']:
                                return pd.NaT
                            
                            # Remover espaços extras
                            data_str = str(data_str).strip()
                            
                            # Padrões comuns
                            padroes = [
                                # dd/mm/aaaa
                                r'(\d{1,2})/(\d{1,2})/(\d{4})',
                                # dd-mm-aaaa
                                r'(\d{1,2})-(\d{1,2})-(\d{4})',
                                # dd.mm.aaaa
                                r'(\d{1,2})\.(\d{1,2})\.(\d{4})',
                                # aaaa-mm-dd (formato ISO)
                                r'(\d{4})-(\d{1,2})-(\d{1,2})',
                            ]
                            
                            for padrao in padroes:
                                match = re.search(padrao, data_str)
                                if match:
                                    grupos = match.groups()
                                    if len(grupos) == 3:
                                        try:
                                            # Formato dd/mm/aaaa ou dd-mm-aaaa
                                            if '/' in data_str or '-' in data_str:
                                                if int(grupos[0]) <= 31:  # Provavelmente dd/mm/aaaa
                                                    dia, mes, ano = grupos
                                                    # Garantir que temos números
                                                    dia = int(dia)
                                                    mes = int(mes)
                                                    ano = int(ano)
                                                    return pd.Timestamp(year=ano, month=mes, day=dia)
                                                else:  # Provavelmente aaaa-mm-dd
                                                    ano, mes, dia = grupos
                                                    dia = int(dia)
                                                    mes = int(mes)
                                                    ano = int(ano)
                                                    return pd.Timestamp(year=ano, month=mes, day=dia)
                                        except:
                                            continue
                            
                            # Se não encontrou padrão, tentar pandas diretamente
                            try:
                                # Primeiro formato brasileiro
                                data_dt = pd.to_datetime(data_str, dayfirst=True, errors='coerce')
                                if pd.notna(data_dt):
                                    return data_dt
                                
                                # Tentar formato americano
                                data_dt = pd.to_datetime(data_str, dayfirst=False, errors='coerce')
                                if pd.notna(data_dt):
                                    return data_dt
                                
                                # Tentar qualquer formato
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
                        
                        # Criar Mes_Ano
                        df['Mes_Ano'] = df.apply(
                            lambda row: f"{int(row['Mes']):02d}/{int(row['Ano'])}" 
                            if pd.notna(row['Mes']) and pd.notna(row['Ano']) 
                            else "Sem Data", 
                            axis=1
                        )
                        
                        # Mostrar distribuição de Mes_Ano
                        print(f"\n  📊 DISTRIBUIÇÃO DE MES_ANO:")
                        distribuicao = df['Mes_Ano'].value_counts().sort_index()
                        for mes_ano, contagem in distribuicao.items():
                            print(f"     {mes_ano}: {contagem} registros")
                        
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
                    else:
                        print(f"  ⚠️ Coluna de Relatório não encontrada")
                        df['Relatorio'] = df.get('ID', 'Sem Relatório').astype(str)
                    
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
            if 'Mes_Ano' in df_risco.columns:
                periodos_unicos = df_risco['Mes_Ano'].dropna().unique()
                print(f"  Períodos únicos encontrados ({len(periodos_unicos)}):")
                for periodo in sorted(periodos_unicos):
                    contagem = len(df_risco[df_risco['Mes_Ano'] == periodo])
                    print(f"    {periodo}: {contagem} registros")
        
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

anos_disponiveis = obter_anos_disponiveis(df_checklist)
print(f"\nDEBUG: Anos disponíveis no filtro: {anos_disponiveis}")

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
        df_nao_conforme_display = df_nao_conforme.copy()
        
        colunas_para_remover = ['Ano', 'Mes', 'Mes_Ano']
        for col in colunas_para_remover:
            if col in df_nao_conforme_display.columns:
                df_nao_conforme_display = df_nao_conforme_display.drop(columns=[col])
        
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
            style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': '#f9e6e6'},
                                    {'if': {'row_index': 'even'}, 'backgroundColor': '#fdecea'}]
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
        print(f"\n📋 PROCESSANDO MATRIZ DE RISCO:")
        print(f"  Total de registros: {len(df_risco)}")
        
        df_risco_filtrado = df_risco.copy()
        
        # Aplicar filtros
        filtros_aplicados = False
        
        if 'Ano' in df_risco_filtrado.columns:
            df_risco_filtrado['Ano'] = pd.to_numeric(df_risco_filtrado['Ano'], errors='coerce')
        
        if 'Mes' in df_risco_filtrado.columns:
            df_risco_filtrado['Mes'] = pd.to_numeric(df_risco_filtrado['Mes'], errors='coerce')
        
        if ano != 'todos' and 'Ano' in df_risco_filtrado.columns:
            try:
                ano_int = int(ano)
                df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Ano'] == ano_int]
                print(f"  ✅ Filtro ANO aplicado: {ano_int}")
                filtros_aplicados = True
            except:
                pass
        
        if mes != 'todos' and 'Mes' in df_risco_filtrado.columns:
            try:
                mes_int = int(mes)
                df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Mes'] == mes_int]
                print(f"  ✅ Filtro MÊS aplicado: {mes_int}")
                filtros_aplicados = True
            except:
                pass
        
        if unidade != 'todas' and 'Unidade' in df_risco_filtrado.columns:
            df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade]
            print(f"  ✅ Filtro UNIDADE aplicado: '{unidade}'")
            filtros_aplicados = True

        print(f"\n📋 Matriz após filtros: {len(df_risco_filtrado)} registros")
        
        if len(df_risco_filtrado) > 0:
            # Garantir que temos coluna Mes_Ano
            if 'Mes_Ano' not in df_risco_filtrado.columns:
                if 'Mes' in df_risco_filtrado.columns and 'Ano' in df_risco_filtrado.columns:
                    df_risco_filtrado['Mes_Ano'] = df_risco_filtrado.apply(
                        lambda row: f"{int(row['Mes']):02d}/{int(row['Ano'])}" 
                        if pd.notna(row['Mes']) and pd.notna(row['Ano']) 
                        else "Sem Data", 
                        axis=1
                    )
                else:
                    df_risco_filtrado['Mes_Ano'] = "Sem Data"
            
            # Agrupar dados por Unidade e Mes_Ano
            unidades = sorted(df_risco_filtrado['Unidade'].dropna().unique())
            meses_anos = sorted(df_risco_filtrado['Mes_Ano'].dropna().unique())
            
            print(f"  📊 Unidades encontradas: {len(unidades)}")
            print(f"  📅 Períodos encontrados: {len(meses_anos)}")
            print(f"  Períodos: {meses_anos}")
            
            # Filtrar apenas períodos válidos (não "Sem Data")
            meses_anos_validos = [ma for ma in meses_anos if ma != "Sem Data" and pd.notna(ma)]
            
            if len(meses_anos_validos) == 0:
                abas_extra.append(html.Div([
                    html.H3("📋 Matriz Auditoria Risco"),
                    html.P("⚠️ Não foi possível extrair períodos (mês/ano) das datas.", 
                           style={'textAlign':'center', 'color':'#e74c3c', 'padding': '20px'}),
                    html.P("As datas na planilha podem estar em formato incorreto.", 
                           style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '10px'})
                ], style={'marginTop':'30px'}))
            else:
                # Criar estrutura de dados para a matriz
                matriz_data = []
                
                for unidade_nome in unidades:
                    linha = {'Unidade': unidade_nome}
                    df_unidade = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade_nome]
                    
                    for mes_ano in meses_anos_validos:
                        df_mes = df_unidade[df_unidade['Mes_Ano'] == mes_ano]
                        
                        if len(df_mes) > 0:
                            relatorios_html = []
                            for _, row in df_mes.iterrows():
                                relatorio = str(row.get('Relatorio', 'Sem Relatório'))
                                status = str(row.get('Status', 'Sem Status'))
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
                
                # Criar tabela HTML
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
                
                for mes_ano in meses_anos_validos:
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
                    
                    for mes_ano in meses_anos_validos:
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
                
                titulo_matriz = f"📋 Matriz Auditoria Risco ({len(df_risco_filtrado)} registros)"
                
                abas_extra.append(html.Div([
                    html.H3(titulo_matriz, style={'marginBottom': '20px'}),
                    html.P(f"Período: {ano if ano != 'todos' else 'Todos'} {f'Mês: {mes}' if mes != 'todos' else ''}", 
                           style={'color': '#7f8c8d', 'marginBottom': '5px'}),
                    html.P(f"Unidades: {len(unidades)} | Períodos: {len(meses_anos_validos)}", 
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

    # ---------- Melhorias e Políticas ----------
    if df_melhorias is not None and len(df_melhorias) > 0:
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

    if df_politicas is not None and len(df_politicas) > 0:
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
