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

def criar_sigla_relatorio(relatorio, index):
    """Cria uma sigla única para o relatório baseada no nome ou índice"""
    if pd.isna(relatorio) or str(relatorio).strip() == '':
        return f"R{index:03d}"
    
    relatorio_str = str(relatorio).strip().upper()
    
    # Tenta extrair sigla de padrões comuns
    # Padrão: "AUD-2023-001" -> "A23001"
    padrao_aud = re.match(r'([A-Z]{2,})-(\d{2,4})-(\d{2,})', relatorio_str)
    if padrao_aud:
        prefixo = padrao_aud.group(1)[:2]  # Primeiras 2 letras
        ano = padrao_aud.group(2)[-2:]     # Últimos 2 dígitos do ano
        numero = padrao_aud.group(3)[:3]   # Primeiros 3 dígitos do número
        return f"{prefixo}{ano}{numero}"
    
    # Padrão: "RELATORIO 2023-001" -> "R23001"
    padrao_rel = re.match(r'RELATORIO\s*(\d{2,4})-?(\d{2,})', relatorio_str, re.IGNORECASE)
    if padrao_rel:
        ano = padrao_rel.group(1)[-2:]
        numero = padrao_rel.group(2)[:3]
        return f"R{ano}{numero}"
    
    # Padrão: "AUDITORIA JANEIRO 2023" -> "AJAN23"
    palavras = re.findall(r'[A-Z]{2,}', relatorio_str)
    if palavras and len(palavras[0]) >= 2:
        primeira_palavra = palavras[0][:2]
        
        # Tenta extrair ano
        anos = re.findall(r'\b(20\d{2})\b', relatorio_str)
        if anos:
            ano_short = anos[0][-2:]
            return f"{primeira_palavra}{ano_short}"
    
    # Se não encontrou padrão, cria sigla baseada nas primeiras letras
    palavras = relatorio_str.split()
    if palavras:
        # Pega primeira letra de cada palavra (até 3 palavras)
        sigla = ''.join([p[0] for p in palavras[:3] if len(p) > 0])
        if len(sigla) >= 2:
            # Adiciona índice para garantir unicidade
            return f"{sigla[:3]}{index:02d}"
    
    # Último recurso: usa as primeiras letras do relatório
    if len(relatorio_str) >= 3:
        return f"{relatorio_str[:3]}{index:02d}"
    
    # Se tudo falhar, usa o índice
    return f"REL{index:03d}"

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
                        
                        # Criar Mes_Ano para exibição (usado apenas no checklist)
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
                        # Criar siglas para os relatórios
                        print(f"  🔤 Criando siglas para os relatórios...")
                        siglas = []
                        for idx, relatorio in enumerate(df['Relatorio']):
                            sigla = criar_sigla_relatorio(relatorio, idx)
                            siglas.append(sigla)
                        df['Sigla'] = siglas
                        print(f"  ✅ Siglas criadas. Exemplos: {siglas[:10]}")
                    else:
                        print(f"  ⚠️ Coluna de Relatório não encontrada")
                        df['Relatorio'] = df.get('ID', 'Sem Relatório').astype(str)
                        # Criar siglas padrão
                        siglas = []
                        for idx in range(len(df)):
                            siglas.append(f"REL{idx:03d}")
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
    # Lista completa de todos os meses do ano
    nomes_meses = {
        1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 
        5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago', 
        9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
    }
    
    # Retorna todos os meses de 1 a 12
    return [{'label': f'{nomes_meses[m]}', 'value': m} for m in range(1, 13)]

def criar_matriz_risco_anual(df_risco_filtrado, ano_filtro):
    """Cria matriz de risco com todos os meses do ano - SIGLAS VISÍVEIS"""
    
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
                for _, row in df_mes.iterrows():
                    relatorio = str(row.get('Relatorio', ''))
                    sigla = str(row.get('Sigla', 'REL'))
                    status = str(row.get('Status', 'Sem Status'))
                    cores = get_status_color(status)
                    
                    # Armazenar mapeamento sigla -> relatório completo
                    mapeamento_siglas[sigla] = relatorio
                    
                    # Criar elemento com a SIGLA visível
                    relatorio_item = html.Div([
                        html.Div(
                            sigla,
                            style={
                                'fontSize': '9px',
                                'fontWeight': 'bold',
                                'color': cores['text_color'],
                                'textAlign': 'center',
                                'overflow': 'hidden',
                                'textOverflow': 'ellipsis',
                                'whiteSpace': 'nowrap'
                            }
                        ),
                        html.Div(
                            "●",
                            style={
                                'fontSize': '8px',
                                'color': cores['border_color'],
                                'textAlign': 'center',
                                'marginTop': '1px'
                            }
                        )
                    ],
                    title=f"Sigla: {sigla}\nRelatório: {relatorio}\nStatus: {status}\nMês: {nomes_completos[mes]}",
                    style={
                        'display': 'flex',
                        'flexDirection': 'column',
                        'justifyContent': 'center',
                        'alignItems': 'center',
                        'backgroundColor': cores['bg_color'],
                        'padding': '3px 2px',
                        'margin': '1px',
                        'borderRadius': '2px',
                        'border': f'1px solid {cores["border_color"]}',
                        'cursor': 'help',
                        'minWidth': '40px',
                        'maxWidth': '45px',
                        'height': '35px'
                    })
                    relatorios_info.append(relatorio_item)
                
                if len(relatorios_info) > 0:
                    # Se houver mais de 1 relatório, mostrar contagem
                    if len(relatorios_info) > 1:
                        contador = html.Div(
                            f"+{len(relatorios_info)-1}",
                            style={
                                'fontSize': '8px',
                                'backgroundColor': '#e74c3c',
                                'color': 'white',
                                'padding': '1px 3px',
                                'borderRadius': '2px',
                                'marginLeft': '2px',
                                'fontWeight': 'bold'
                            },
                            title=f"Total: {len(relatorios_info)} relatórios neste mês"
                        )
                        
                        # Mostrar apenas o primeiro relatório + contador
                        linha[mes] = html.Div([
                            relatorios_info[0],
                            contador
                        ], style={
                            'display': 'flex',
                            'alignItems': 'center',
                            'justifyContent': 'center',
                            'gap': '2px'
                        })
                    else:
                        # Apenas um relatório
                        linha[mes] = relatorios_info[0]
                else:
                    linha[mes] = html.Div("-", style={'color': '#ccc', 'fontSize': '10px', 'padding': '10px 0'})
            else:
                linha[mes] = html.Div("-", style={'color': '#ccc', 'fontSize': '10px', 'padding': '10px 0'})
        
        matriz_data.append(linha)
    
    # Criar lista de siglas para legenda
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
                'minWidth': '50px',
                'maxWidth': '60px',
                'fontSize': '11px'
            },
            title=nomes_completos[mes]
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
                    'minHeight': '45px',
                    'minWidth': '50px',
                    'maxWidth': '60px'
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
    
    # Criar legenda de siglas (agrupadas)
    linhas_legenda = []
    linha_atual = []
    for i, sigla in enumerate(siglas_unicas):
        relatorio_completo = mapeamento_siglas[sigla]
        
        item_legenda = html.Div([
            html.Span(f"{sigla}: ", style={
                'fontWeight': 'bold',
                'color': '#2c3e50',
                'fontSize': '10px',
                'minWidth': '50px',
                'display': 'inline-block'
            }),
            html.Span(relatorio_completo[:30] + ("..." if len(relatorio_completo) > 30 else ""), {
                'color': '#7f8c8d',
                'fontSize': '10px',
                'overflow': 'hidden',
                'textOverflow': 'ellipsis',
                'whiteSpace': 'nowrap',
                'maxWidth': '200px',
                'display': 'inline-block'
            })
        ], style={
            'display': 'flex',
            'alignItems': 'center',
            'marginBottom': '3px',
            'padding': '2px 5px',
            'borderBottom': '1px solid #f0f0f0'
        })
        
        linha_atual.append(item_legenda)
        
        # Cada linha da legenda terá 2 colunas
        if len(linha_atual) == 2 or i == len(siglas_unicas) - 1:
            linhas_legenda.append(html.Div(
                linha_atual,
                style={'display': 'flex', 'justifyContent': 'space-between', 'gap': '10px', 'marginBottom': '5px'}
            ))
            linha_atual = []
    
    # Container da legenda com scroll
    legenda_container = html.Div(
        linhas_legenda,
        style={
            'maxHeight': '150px',
            'overflowY': 'auto',
            'padding': '5px',
            'border': '1px solid #dee2e6',
            'borderRadius': '3px',
            'backgroundColor': '#f8f9fa',
            'marginBottom': '10px'
        }
    )
    
    # Legenda completa
    legenda_completa = html.Div([
        html.P("📋 LEGENDA DE SIGLAS - MATRIZ DE RISCO", style={
            'marginBottom': '5px', 
            'color': '#2c3e50', 
            'fontSize': '12px', 
            'fontWeight': 'bold',
            'display': 'flex',
            'alignItems': 'center',
            'gap': '5px'
        }),
        html.P(f"Total de {len(siglas_unicas)} relatórios encontrados", 
               style={'color': '#7f8c8d', 'marginBottom': '8px', 'fontSize': '10px'}),
        html.Div([
            html.Div([
                html.Div("A23001", style={
                    'fontSize': '9px',
                    'fontWeight': 'bold',
                    'color': '#c0392b',
                    'backgroundColor': '#fdecea',
                    'padding': '3px 5px',
                    'borderRadius': '2px',
                    'border': '1px solid #c0392b',
                    'marginRight': '5px',
                    'minWidth': '45px',
                    'textAlign': 'center'
                }),
                html.Span("Não Iniciado", style={'color': '#2c3e50', 'fontSize': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '15px'}),
            
            html.Div([
                html.Div("R23002", style={
                    'fontSize': '9px',
                    'fontWeight': 'bold',
                    'color': '#f39c12',
                    'backgroundColor': '#fff8e1',
                    'padding': '3px 5px',
                    'borderRadius': '2px',
                    'border': '1px solid #f39c12',
                    'marginRight': '5px',
                    'minWidth': '45px',
                    'textAlign': 'center'
                }),
                html.Span("Pendente", style={'color': '#2c3e50', 'fontSize': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginRight': '15px'}),
            
            html.Div([
                html.Div("A23003", style={
                    'fontSize': '9px',
                    'fontWeight': 'bold',
                    'color': '#27ae60',
                    'backgroundColor': '#eafaf1',
                    'padding': '3px 5px',
                    'borderRadius': '2px',
                    'border': '1px solid #27ae60',
                    'marginRight': '5px',
                    'minWidth': '45px',
                    'textAlign': 'center'
                }),
                html.Span("Finalizado", style={'color': '#2c3e50', 'fontSize': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center'})
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '10px', 'alignItems': 'center', 'marginBottom': '10px'}),
        
        html.P("Lista Completa de Siglas:", style={'marginBottom': '5px', 'color': '#2c3e50', 'fontSize': '11px', 'fontWeight': 'bold'}),
        legenda_container,
        
        html.P("ℹ️ Passe o mouse sobre as siglas na matriz para ver detalhes completos", 
               style={'color': '#7f8c8d', 'marginTop': '8px', 'fontSize': '9px', 'fontStyle': 'italic'})
    ], style={
        'backgroundColor': '#f8f9fa',
        'padding': '12px 15px',
        'borderRadius': '3px',
        'marginBottom': '12px',
        'border': '1px solid #dee2e6',
        'fontSize': '11px'
    })
    
    titulo_matriz = f"📋 Matriz de Risco - {ano_filtro} ({len(df_risco_filtrado)} registros)"
    
    return html.Div([
        html.H4(titulo_matriz, style={
            'marginBottom': '10px', 
            'color': '#2c3e50',
            'fontSize': '16px',
            'display': 'flex',
            'alignItems': 'center',
            'gap': '8px'
        }),
        legenda_completa,
        tabela_container
    ], style={
        'marginTop': '20px',
        'padding': '15px',
        'backgroundColor': 'white',
        'borderRadius': '5px',
        'boxShadow': '0 1px 3px rgba(0,0,0,0.1)',
        'overflow': 'hidden'
    })

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
    html.Div([
        html.H1("📊 DASHBOARD DE AUDITORIA", 
                style={
                    'textAlign':'center', 
                    'marginBottom':'15px',
                    'fontSize': '22px',
                    'color': '#2c3e50'
                })
    ]),
    html.Div([
        html.Div([
            html.Label("Ano:", style={'fontSize': '12px', 'fontWeight': 'bold', 'marginBottom': '3px'}),
            dcc.Dropdown(
                id='filtro-ano',
                options=[{'label':'Todos','value':'todos'}]+
                       [{'label':str(a),'value':a} for a in anos_disponiveis],
                value='todos',
                style={'fontSize': '12px'}
            )
        ], style={'marginRight':'15px','width':'180px'}),
        html.Div([
            html.Label("Mês:", style={'fontSize': '12px', 'fontWeight': 'bold', 'marginBottom': '3px'}),
            dcc.Dropdown(
                id='filtro-mes',
                options=[{'label':'Todos','value':'todos'}]+
                       [{'label':f'{m}','value':m} for m in range(1, 13)],
                value='todos',
                style={'fontSize': '12px'}
            )
        ], style={'marginRight':'15px','width':'150px'}),
        html.Div([
            html.Label("Unidade:", style={'fontSize': '12px', 'fontWeight': 'bold', 'marginBottom': '3px'}),
            dcc.Dropdown(
                id='filtro-unidade',
                options=[{'label':'Todas','value':'todas'}]+
                        [{'label':str(u),'value':str(u)} for u in sorted(df_checklist['Unidade'].dropna().unique())],
                value='todas',
                style={'fontSize': '12px'}
            )
        ], style={'width':'200px'})
    ], style={'display':'flex','justifyContent':'center','marginBottom':'25px','flexWrap':'wrap'}),
    html.Div(id='conteudo-principal', style={'padding':'15px', 'maxWidth': '1400px', 'margin': '0 auto'})
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

    # ---------- KPIs GERAIS ----------
    kpis = html.Div([
        html.Div([
            html.H4("Conforme", style={'color':'#27ae60','margin':'0', 'fontSize': '14px'}),
            html.H2(f"{conforme}", style={'color':'#27ae60','margin':'0', 'fontSize': '28px'}),
            html.P(f"{(conforme/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#27ae60', 'fontSize': '12px'})
        ], style={'borderLeft':'5px solid #27ae60','borderRadius':'5px','padding':'15px','margin':'8px','flex':'1',
                  'backgroundColor':'#eafaf1','textAlign':'center','boxShadow':'1px 1px 3px rgba(0,0,0,0.1)',
                  'minWidth': '150px'}),

        html.Div([
            html.H4("Conforme Parcialmente", style={'color':'#f39c12','margin':'0', 'fontSize': '14px'}),
            html.H2(f"{parcial}", style={'color':'#f39c12','margin':'0', 'fontSize': '28px'}),
            html.P(f"{(parcial/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#f39c12', 'fontSize': '12px'})
        ], style={'borderLeft':'5px solid #f39c12','borderRadius':'5px','padding':'15px','margin':'8px','flex':'1',
                  'backgroundColor':'#fff8e1','textAlign':'center','boxShadow':'1px 1px 3px rgba(0,0,0,0.1)',
                  'minWidth': '150px'}),

        html.Div([
            html.H4("Não Conforme", style={'color':'#e74c3c','margin':'0', 'fontSize': '14px'}),
            html.H2(f"{nao}", style={'color':'#e74c3c','margin':'0', 'fontSize': '28px'}),
            html.P(f"{(nao/total*100 if total>0 else 0):.1f}%", style={'margin':'0','color':'#e74c3c', 'fontSize': '12px'})
        ], style={'borderLeft':'5px solid #e74c3c','borderRadius':'5px','padding':'15px','margin':'8px','flex':'1',
                  'backgroundColor':'#fdecea','textAlign':'center','boxShadow':'1px 1px 3px rgba(0,0,0,0.1)',
                  'minWidth': '150px'})
    ], style={'display':'flex','justifyContent':'center','flexWrap':'wrap','marginBottom':'25px', 'gap': '5px'})

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
            
            # Criar tabela com cores condicionais - FONTE MENOR
            tabela_nao_conforme = dash_table.DataTable(
                columns=[{"name": col, "id": col} for col in df_nao_conforme_display.columns],
                data=df_nao_conforme_display.to_dict('records'),
                page_size=10,
                style_table={'overflowX':'auto', 'fontSize': '11px'},
                style_header={
                    'backgroundColor': '#c0392b',
                    'color': 'white',
                    'fontWeight': 'bold',
                    'textAlign':'center',
                    'fontSize': '11px',
                    'padding': '6px 8px'
                },
                style_cell={
                    'textAlign': 'center',
                    'padding': '4px 6px',
                    'whiteSpace':'normal',
                    'height':'auto',
                    'fontSize': '10px',
                    'minWidth': '50px'
                },
                style_data_conditional=[
                    # Linhas pares
                    {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9e6e6'},
                    # Linhas ímpares
                    {'if': {'row_index': 'even'}, 'backgroundColor': '#fdecea'},
                    # Cores condicionais para Status do Prazo
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
            # Se não encontrou as colunas, mostrar tabela normal
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
            
            tabela_nao_conforme = dash_table.DataTable(
                columns=[{"name": col, "id": col} for col in df_nao_conforme_display.columns],
                data=df_nao_conforme_display.to_dict('records'),
                page_size=10,
                style_table={'overflowX':'auto', 'fontSize': '11px'},
                style_header={
                    'backgroundColor': '#c0392b',
                    'color': 'white',
                    'fontWeight': 'bold',
                    'textAlign':'center',
                    'fontSize': '11px',
                    'padding': '6px 8px'
                },
                style_cell={
                    'textAlign': 'center',
                    'padding': '4px 6px',
                    'whiteSpace':'normal',
                    'height':'auto',
                    'fontSize': '10px',
                    'minWidth': '50px'
                },
                style_data_conditional=[
                    {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9e6e6'},
                    {'if': {'row_index': 'even'}, 'backgroundColor': '#fdecea'}
                ]
            )
    else:
        tabela_nao_conforme = html.Div([
            html.P("✅ Nenhum item não conforme encontrado com os filtros atuais.", 
                   style={'textAlign': 'center', 'padding': '15px', 'color': '#27ae60', 'fontSize': '12px'})
        ])
    
    tabela_titulo = html.H3(f"❌ Itens Não Conformes ({len(df_nao_conforme)} itens)", 
                           style={'marginTop': '25px', 'marginBottom': '10px', 'color': '#c0392b', 'fontSize': '16px'})
    
    # ---------- KPIs de PRAZOS dos Itens Não Conformes ----------
    if len(df_nao_conforme) > 0 and (coluna_prazo and coluna_finalizacao):
        kpis_prazos = html.Div([
            html.Div([
                html.H4("Dentro do Prazo", style={'color':'#27ae60','margin':'0', 'fontSize': '12px'}),
                html.H2(f"{dentro_prazo}", style={'color':'#27ae60','margin':'0', 'fontSize': '24px'}),
                html.P(f"{(dentro_prazo/len(df_nao_conforme)*100 if len(df_nao_conforme)>0 else 0):.1f}%", 
                       style={'margin':'0','color':'#27ae60', 'fontSize': '10px'})
            ], style={'borderLeft':'4px solid #27ae60','borderRadius':'4px','padding':'10px','margin':'5px','flex':'1',
                      'backgroundColor':'#d4edda','textAlign':'center','boxShadow':'1px 1px 3px rgba(0,0,0,0.1)',
                      'minWidth': '120px', 'maxWidth': '140px'}),

            html.Div([
                html.H4("Fora do Prazo", style={'color':'#e74c3c','margin':'0', 'fontSize': '12px'}),
                html.H2(f"{fora_prazo}", style={'color':'#e74c3c','margin':'0', 'fontSize': '24px'}),
                html.P(f"{(fora_prazo/len(df_nao_conforme)*100 if len(df_nao_conforme)>0 else 0):.1f}%", 
                       style={'margin':'0','color':'#e74c3c', 'fontSize': '10px'})
            ], style={'borderLeft':'4px solid #e74c3c','borderRadius':'4px','padding':'10px','margin':'5px','flex':'1',
                      'backgroundColor':'#f8d7da','textAlign':'center','boxShadow':'1px 1px 3px rgba(0,0,0,0.1)',
                      'minWidth': '120px', 'maxWidth': '140px'}),

            html.Div([
                html.H4("Não Concluídos", style={'color':'#f39c12','margin':'0', 'fontSize': '12px'}),
                html.H2(f"{nao_concluido}", style={'color':'#f39c12','margin':'0', 'fontSize': '24px'}),
                html.P(f"{(nao_concluido/len(df_nao_conforme)*100 if len(df_nao_conforme)>0 else 0):.1f}%", 
                       style={'margin':'0','color':'#f39c12', 'fontSize': '10px'})
            ], style={'borderLeft':'4px solid #f39c12','borderRadius':'4px','padding':'10px','margin':'5px','flex':'1',
                      'backgroundColor':'#fff3cd','textAlign':'center','boxShadow':'1px 1px 3px rgba(0,0,0,0.1)',
                      'minWidth': '120px', 'maxWidth': '140px'})
        ], style={'display':'flex','justifyContent':'center','flexWrap':'wrap','marginBottom':'15px', 'gap': '5px'})
        
        # Adicionar legenda para KPIs de prazo
        legenda_prazo = html.Div([
            html.P("📊 Status dos Prazos dos Itens Não Conformes:", 
                   style={'fontWeight':'bold','marginBottom':'3px', 'fontSize': '11px', 'textAlign': 'center'}),
            html.Div([
                html.Span("🟢 ", style={'color':'#27ae60','marginRight':'3px', 'fontSize': '10px'}),
                html.Span("Dentro do Prazo", style={'marginRight':'10px','color':'#27ae60', 'fontSize': '10px'}),
                html.Span("🔴 ", style={'color':'#e74c3c','marginRight':'3px', 'fontSize': '10px'}),
                html.Span("Fora do Prazo", style={'marginRight':'10px','color':'#e74c3c', 'fontSize': '10px'}),
                html.Span("🟡 ", style={'color':'#f39c12','marginRight':'3px', 'fontSize': '10px'}),
                html.Span("Não Concluído", style={'color':'#f39c12', 'fontSize': '10px'})
            ], style={'backgroundColor':'#f8f9fa','padding':'6px 8px','borderRadius':'3px','marginBottom':'8px', 
                      'fontSize': '10px', 'textAlign': 'center'})
        ], style={'marginBottom':'15px', 'textAlign': 'center'})
    else:
        kpis_prazos = html.Div()
        legenda_prazo = html.Div()

    # ---------- Matriz de Risco (APENAS ANO) ----------
    abas_extra = []

    if df_risco is not None and len(df_risco) > 0:
        print(f"\n📋 PROCESSANDO MATRIZ DE RISCO:")
        print(f"  Total de registros: {len(df_risco)}")
        
        df_risco_filtrado = df_risco.copy()
        
        # Aplicar APENAS filtro de ano para a matriz de risco
        filtros_aplicados = False
        
        if 'Ano' in df_risco_filtrado.columns:
            df_risco_filtrado['Ano'] = pd.to_numeric(df_risco_filtrado['Ano'], errors='coerce')
        
        # Aplicar filtro de ano se não for 'todos'
        if ano != 'todos' and 'Ano' in df_risco_filtrado.columns:
            try:
                ano_int = int(ano)
                df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Ano'] == ano_int]
                print(f"  ✅ Filtro ANO aplicado para matriz de risco: {ano_int}")
                filtros_aplicados = True
            except:
                pass
        
        # Aplicar filtro de unidade se não for 'todas'
        if unidade != 'todas' and 'Unidade' in df_risco_filtrado.columns:
            df_risco_filtrado = df_risco_filtrado[df_risco_filtrado['Unidade'] == unidade]
            print(f"  ✅ Filtro UNIDADE aplicado para matriz de risco: '{unidade}'")
            filtros_aplicados = True
        
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
                html.H3("📋 Matriz Auditoria Risco", style={'fontSize': '16px'}),
                html.P("Nenhum dado encontrado para o ano selecionado.", 
                       style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '20px', 'fontSize': '12px'})
            ], style={'marginTop':'20px'}))
    else:
        abas_extra.append(html.Div([
            html.H3("📋 Matriz Auditoria Risco", style={'fontSize': '16px'}),
            html.P("Não há dados de risco disponíveis.", 
                   style={'textAlign':'center', 'color':'#7f8c8d', 'padding': '20px', 'fontSize': '12px'})
        ], style={'marginTop':'20px'}))

    # ---------- Melhorias e Políticas ----------
    if df_melhorias is not None and len(df_melhorias) > 0:
        colunas_data_melhorias = [col for col in df_melhorias.columns 
                                 if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
        
        df_melhorias_display = df_melhorias.copy()
        for coluna_data in colunas_data_melhorias:
            if coluna_data in df_melhorias_display.columns:
                df_melhorias_display[coluna_data] = df_melhorias_display[coluna_data].apply(formatar_data)
        
        tabela_melhorias = dash_table.DataTable(
            columns=[{"name": col, "id": col} for col in df_melhorias_display.columns],
            data=df_melhorias_display.to_dict('records'),
            page_size=10,
            style_table={'overflowX':'auto','marginTop':'8px', 'fontSize': '11px'},
            style_header={
                'backgroundColor': '#34495e',
                'color': 'white',
                'fontWeight': 'bold',
                'textAlign':'center',
                'fontSize': '11px',
                'padding': '6px 8px'
            },
            style_cell={
                'textAlign': 'center',
                'padding': '4px 6px',
                'whiteSpace':'normal',
                'height':'auto',
                'fontSize': '10px'
            },
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#ecf0f1'},
                {'if': {'row_index': 'even'}, 'backgroundColor': 'white'}
            ]
        )
        abas_extra.append(html.Div([
            html.H3(f"📈 Melhorias ({len(df_melhorias)} registros)", style={'fontSize': '16px', 'marginBottom': '8px'}),
            html.P("Todos os registros de melhorias (filtros não aplicados)", 
                   style={'color': '#7f8c8d', 'marginBottom': '8px', 'fontSize': '11px'}),
            tabela_melhorias
        ], style={'marginTop':'20px'}))

    if df_politicas is not None and len(df_politicas) > 0:
        colunas_data_politicas = [col for col in df_politicas.columns 
                                 if any(termo in col.lower() for termo in ['data', 'prazo', 'vencimento', 'limite', 'criacao', 'conclusao'])]
        
        df_politicas_display = df_politicas.copy()
        for coluna_data in colunas_data_politicas:
            if coluna_data in df_politicas_display.columns:
                df_politicas_display[coluna_data] = df_politicas_display[coluna_data].apply(formatar_data)
        
        tabela_politicas = dash_table.DataTable(
            columns=[{"name": col, "id": col} for col in df_politicas_display.columns],
            data=df_politicas_display.to_dict('records'),
            page_size=10,
            style_table={'overflowX':'auto','marginTop':'8px', 'fontSize': '11px'},
            style_header={
                'backgroundColor': '#34495e',
                'color': 'white',
                'fontWeight': 'bold',
                'textAlign':'center',
                'fontSize': '11px',
                'padding': '6px 8px'
            },
            style_cell={
                'textAlign': 'center',
                'padding': '4px 6px',
                'whiteSpace':'normal',
                'height':'auto',
                'fontSize': '10px'
            },
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#ecf0f1'},
                {'if': {'row_index': 'even'}, 'backgroundColor': 'white'}
            ]
        )
        abas_extra.append(html.Div([
            html.H3(f"📑 Políticas ({len(df_politicas)} registros)", style={'fontSize': '16px', 'marginBottom': '8px'}),
            html.P("Todos os registros de políticas (filtros não aplicados)", 
                   style={'color': '#7f8c8d', 'marginBottom': '8px', 'fontSize': '11px'}),
            tabela_politicas
        ], style={'marginTop':'20px'}))

    # ---------- Layout Final ----------
    return html.Div([
        html.Div([
            html.H4(f"📊 Resumo - {len(df)} itens auditados", 
                    style={'textAlign':'center', 'color':'#2c3e50', 'marginBottom':'15px', 'fontSize': '18px'})
        ]),
        kpis,
        tabela_titulo,
        # KPIs de PRAZOS dos Itens Não Conformes
        kpis_prazos,
        legenda_prazo,
        tabela_nao_conforme,
        *abas_extra
    ])

# ========== EXECUÇÃO DO APP ==========
if __name__ == '__main__':
    print("🌐 DASHBOARD RODANDO: http://localhost:8050")
    print("📊 DASHBOARD OTIMIZADO:")
    print("  - ✅ Gráfico de pizza REMOVIDO")
    print("  - ✅ KPIs de PRAZOS para itens não conformes")
    print("  - ✅ Matriz de risco COM SIGLAS VISÍVEIS")
    print("  - ✅ Legenda completa com lista de siglas")
    print("  - ✅ Passe o mouse sobre as siglas para ver detalhes")
    app.run(debug=True, host='0.0.0.0', port=8050)

# ========== SERVER PARA O RENDER ==========
server = app.server
