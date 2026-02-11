import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import msal
from datetime import datetime, timedelta
import time
import pytz

# =========================================================
# 0. CONFIGURAÇÕES INICIAIS E DEBUG
# =========================================================

# Configurar pandas para mostrar TUDO
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 100)

# Configurar página
st.set_page_config(
    page_title="Dashboard de Campanhas - SICOOB COCRED",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========== BANNER DE VERIFICAÇÃO ==========
st.markdown("""
<div style="background: linear-gradient(90deg, #4CAF50 0%, #2196F3 100%); 
            padding: 15px; 
            border-radius: 8px; 
            color: white; 
            text-align: center; 
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
    <h2 style="margin: 0; font-size: 24px;">🚀 DASHBOARD COCRED - VERSÃO ATUALIZADA</h2>
    <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">
        Atualizado em: 11/02/2026 | Streamlit v""" + st.__version__ + """ | Conectado ao SharePoint
    </p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 1. CONFIGURAÇÕES DA API E SHAREPOINT
# =========================================================

# Credenciais (via secrets.toml no Streamlit Cloud)
MS_CLIENT_ID = st.secrets.get("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = st.secrets.get("MS_CLIENT_SECRET", "")
MS_TENANT_ID = st.secrets.get("MS_TENANT_ID", "")

# Informações do arquivo Excel
USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"
SHAREPOINT_FILE_ID = "01S7YQRRWMBXCV3AAHYZEIZGL55EPOZULE"
SHEET_NAME = "Demandas ID"
EXCEL_ONLINE_URL = "https://agenciaideatore-my.sharepoint.com/:x:/g/personal/cristini_cordesco_ideatoreamericas_com/IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY"

# =========================================================
# 2. FUNÇÕES DE AUTENTICAÇÃO
# =========================================================

@st.cache_resource(ttl=3500)  # Cache de ~58 minutos
def get_msal_app():
    """Configura a aplicação MSAL para autenticação"""
    if not all([MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID]):
        st.error("❌ Credenciais da API não configuradas no Streamlit Secrets!")
        return None
    
    try:
        authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
        app = msal.ConfidentialClientApplication(
            client_id=MS_CLIENT_ID,
            authority=authority,
            client_credential=MS_CLIENT_SECRET
        )
        return app
    except Exception as e:
        st.error(f"❌ Erro ao configurar MSAL: {str(e)}")
        return None

@st.cache_data(ttl=3300)  # Cache de 55 minutos
def get_access_token():
    """Obtém token de acesso para Microsoft Graph API"""
    app = get_msal_app()
    if not app:
        return None
    
    try:
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        
        if "access_token" in result:
            return result["access_token"]
        else:
            st.error(f"❌ Erro ao obter token: {result.get('error_description', 'Unknown error')}")
            return None
    except Exception as e:
        st.error(f"❌ Erro na autenticação: {str(e)}")
        return None

# =========================================================
# 3. FUNÇÃO PRINCIPAL PARA CARREGAR DADOS
# =========================================================

@st.cache_data(ttl=300, show_spinner="🔄 Carregando dados do Excel Online...")  # 5 minutos cache
def carregar_dados_excel():
    """
    Carrega dados do Excel no SharePoint usando Microsoft Graph API
    """
    # Obter token de acesso
    access_token = get_access_token()
    if not access_token:
        st.error("Não foi possível obter token de acesso. Verifique as credenciais.")
        return pd.DataFrame()
    
    # URL para baixar o arquivo
    file_url = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{SHAREPOINT_FILE_ID}/content"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/octet-stream"
    }
    
    try:
        with st.spinner("📥 Conectando ao SharePoint..."):
            response = requests.get(file_url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            # Converter resposta para BytesIO
            excel_file = BytesIO(response.content)
            
            # Tentar ler a aba especificada
            try:
                df = pd.read_excel(
                    excel_file, 
                    sheet_name=SHEET_NAME,
                    engine='openpyxl'
                )
                
                # Log de sucesso (apenas em debug)
                if st.session_state.get('debug_mode', False):
                    st.sidebar.success(f"✅ {len(df)} linhas carregadas da aba '{SHEET_NAME}'")
                
                return df
                
            except Exception as sheet_error:
                # Se falhar na aba específica, tentar primeira aba
                st.warning(f"⚠️ Não foi possível ler a aba '{SHEET_NAME}': {str(sheet_error)[:100]}")
                excel_file.seek(0)
                df = pd.read_excel(excel_file, engine='openpyxl')
                return df
                
        else:
            st.error(f"❌ Erro {response.status_code} ao acessar o arquivo")
            if response.status_code == 401:
                st.error("Token expirado ou inválido")
            elif response.status_code == 404:
                st.error("Arquivo não encontrado. Verifique o FILE_ID.")
            elif response.status_code == 403:
                st.error("Permissão negada. Verifique as permissões da aplicação.")
            
            return pd.DataFrame()
            
    except requests.exceptions.Timeout:
        st.error("⏰ Timeout ao conectar com SharePoint. Tente novamente.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erro inesperado: {str(e)}")
        return pd.DataFrame()

# =========================================================
# 4. FUNÇÕES AUXILIARES E UTILITÁRIOS
# =========================================================

def calcular_altura_tabela(num_linhas, max_altura=800):
    """Calcula altura ideal para tabela baseada no número de linhas"""
    altura_por_linha = 35
    altura_minima = 300
    altura_calculada = altura_minima + (num_linhas * altura_por_linha)
    return min(altura_calculada, max_altura)

def formatar_data(data):
    """Formata data para exibição amigável"""
    if pd.isna(data):
        return "N/A"
    try:
        if isinstance(data, str):
            data = pd.to_datetime(data)
        return data.strftime('%d/%m/%Y')
    except:
        return str(data)

def verificar_credenciais():
    """Verifica se as credenciais estão configuradas"""
    if not MS_CLIENT_ID or not MS_CLIENT_SECRET or not MS_TENANT_ID:
        return False, "Credenciais não configuradas no secrets.toml"
    
    token = get_access_token()
    if not token:
        return False, "Falha na autenticação com Microsoft"
    
    return True, "Credenciais OK"

# =========================================================
# 5. SIDEBAR - CONTROLES E CONFIGURAÇÕES
# =========================================================

with st.sidebar:
    st.title("⚙️ Controles")
    
    # Modo Debug
    if 'debug_mode' not in st.session_state:
        st.session_state.debug_mode = False
    
    st.session_state.debug_mode = st.checkbox("🐛 Modo Debug", value=st.session_state.debug_mode)
    
    st.markdown("---")
    
    # Configurações de Visualização
    st.subheader("👁️ Visualização")
    
    linhas_por_pagina = st.selectbox(
        "Linhas por página:",
        ["50", "100", "200", "500", "Todas"],
        index=1,
        help="Quantidade de linhas a serem exibidas por vez"
    )
    
    mostrar_filtros = st.checkbox("🎛️ Mostrar Filtros Avançados", value=True)
    
    st.markdown("---")
    
    # Atualização de Dados
    st.subheader("🔄 Atualização")
    
    col_atualiza1, col_atualiza2 = st.columns(2)
    
    with col_atualiza1:
        if st.button("🔄 Atualizar", use_container_width=True, type="primary"):
            st.cache_data.clear()
            st.rerun()
    
    with col_atualiza2:
        if st.button("🗑️ Limpar Cache", use_container_width=True, type="secondary"):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.success("Cache limpo!")
            time.sleep(1)
            st.rerun()
    
    # Teste de Conexão
    st.markdown("---")
    st.subheader("🔗 Conexão")
    
    if st.button("🔍 Testar Conexão SharePoint", use_container_width=True):
        with st.spinner("Testando conexão..."):
            token = get_access_token()
            if token:
                st.success("✅ Conectado à API Microsoft Graph")
            else:
                st.error("❌ Falha na conexão")
    
    # Link para Excel Online
    st.markdown("---")
    st.subheader("📝 Editar Dados")
    
    st.markdown(f"""
    [✏️ Abrir Excel Online]({EXCEL_ONLINE_URL})
    
    **Instruções:**
    1. Edite os dados no Excel
    2. Salve as alterações (Ctrl+S)
    3. Clique em **"Atualizar"** ao lado
    4. Aguarde alguns segundos
    """)
    
    # Informações Técnicas
    if st.session_state.debug_mode:
        st.markdown("---")
        st.subheader("🐛 Debug Info")
        
        with st.expander("Detalhes Técnicos"):
            st.write(f"**Streamlit:** v{st.__version__}")
            st.write(f"**Pandas:** v{pd.__version__}")
            st.write(f"**Hora Local:** {datetime.now().strftime('%H:%M:%S')}")
            
            # Verificar credenciais
            cred_ok, msg = verificar_credenciais()
            if cred_ok:
                st.success("✅ " + msg)
            else:
                st.error("❌ " + msg)

# =========================================================
# 6. INTERFACE PRINCIPAL
# =========================================================

# Título Principal
st.title("📊 Dashboard de Campanhas – SICOOB COCRED")
st.caption(f"🔗 Conectado ao Excel Online | Aba: {SHEET_NAME} | Atualizado em: {datetime.now().strftime('%H:%M:%S')}")

# =========================================================
# 7. CARREGAR DADOS
# =========================================================

# Carregar dados com indicador de progresso
with st.spinner("📥 Carregando dados do Excel Online..."):
    df = carregar_dados_excel()

# Verificar se os dados foram carregados
if df.empty:
    st.error("""
    ❌ Não foi possível carregar os dados. Verifique:
    
    1. **Credenciais:** Configure MS_CLIENT_ID, MS_CLIENT_SECRET e MS_TENANT_ID no Streamlit Secrets
    2. **Permissões:** A aplicação precisa ter permissão para acessar o SharePoint
    3. **Arquivo:** Verifique se o arquivo existe e está acessível
    4. **Conexão:** Sua rede pode estar bloqueando a conexão com Microsoft Graph
    
    Clique em **"Testar Conexão SharePoint"** na sidebar para diagnosticar.
    """)
    
    # Mostrar dados de exemplo para teste
    st.info("📋 **Enquanto isso, aqui estão dados de exemplo:**")
    
    dados_exemplo = {
        'ID': list(range(1, 11)),
        'Campanha': [f'Campanha {i}' for i in range(1, 11)],
        'Status': ['Aprovado', 'Em Produção', 'Aguardando', 'Aprovado', 'Em Produção', 
                  'Aguardando', 'Aprovado', 'Concluído', 'Em Produção', 'Aguardando'],
        'Prioridade': ['Alta', 'Média', 'Baixa', 'Alta', 'Média', 'Baixa', 'Alta', 'Média', 'Baixa', 'Alta'],
        'Produção': ['Cocred', 'Ideatore', 'Cocred', 'Ideatore', 'Cocred', 
                    'Ideatore', 'Cocred', 'Ideatore', 'Cocred', 'Ideatore'],
        'Data Solicitação': pd.date_range(start='2024-01-01', periods=10, freq='D'),
        'Prazo (dias)': [5, 3, 10, 2, 7, 14, 1, 0, 5, 3]
    }
    
    df = pd.DataFrame(dados_exemplo)
    st.warning("⚠️ Mostrando dados de exemplo. Os dados reais não foram carregados.")

# Informações do DataFrame
total_linhas = len(df)
total_colunas = len(df.columns)

# Mostrar resumo
col_res1, col_res2, col_res3, col_res4 = st.columns(4)

with col_res1:
    st.metric("📈 Total de Registros", total_linhas)

with col_res2:
    st.metric("📊 Total de Colunas", total_colunas)

with col_res3:
    # Última data disponível
    col_data = None
    for col in ['Data Solicitação', 'Data', 'Data de Solicitação', 'Criado em']:
        if col in df.columns:
            col_data = col
            break
    
    if col_data and col_data in df.columns:
        ultima_data = df[col_data].max()
        st.metric("📅 Última Solicitação", formatar_data(ultima_data))
    else:
        st.metric("📅 Última Atualização", datetime.now().strftime('%d/%m/%Y'))

with col_res4:
    st.metric("🔄 Atualiza em", "5 minutos")

st.divider()

# =========================================================
# 8. VISUALIZAÇÃO DOS DADOS COM PAGINAÇÃO
# =========================================================

st.header("📋 Dados Completos")

# Criar abas para diferentes visualizações
tab1, tab2, tab3, tab4 = st.tabs(["📊 Dados", "📈 Estatísticas", "🔍 Pesquisa", "📤 Exportar"])

with tab1:
    if linhas_por_pagina == "Todas":
        # Mostrar TODOS os dados de uma vez
        altura = calcular_altura_tabela(total_linhas, max_altura=1500)
        
        st.subheader(f"📋 Todos os {total_linhas} registros")
        
        # Mostrar DataFrame completo
        st.dataframe(
            df,
            height=altura,
            width='stretch',
            hide_index=False,
            use_container_width=True
        )
        
        if altura >= 1500:
            st.info(f"ℹ️ Mostrando {total_linhas} linhas. Use o scroll para navegar.")
            
    else:
        # Paginação manual
        linhas_por_pagina = int(linhas_por_pagina)
        total_paginas = max(1, (total_linhas - 1) // linhas_por_pagina + 1)
        
        # Gerenciar estado da página
        if 'pagina_atual' not in st.session_state:
            st.session_state.pagina_atual = 1
        
        # Controles de navegação
        col_nav1, col_nav2, col_nav3, col_nav4, col_nav5 = st.columns([1, 2, 2, 2, 1])
        
        with col_nav1:
            if st.session_state.pagina_atual > 1:
                if st.button("⏮️", help="Primeira página", use_container_width=True):
                    st.session_state.pagina_atual = 1
                    st.rerun()
        
        with col_nav2:
            if st.session_state.pagina_atual > 1:
                if st.button("◀️ Anterior", use_container_width=True):
                    st.session_state.pagina_atual -= 1
                    st.rerun()
        
        with col_nav3:
            st.markdown(f"**Página {st.session_state.pagina_atual} de {total_paginas}**", unsafe_allow_html=True)
        
        with col_nav4:
            if st.session_state.pagina_atual < total_paginas:
                if st.button("Próxima ▶️", use_container_width=True):
                    st.session_state.pagina_atual += 1
                    st.rerun()
        
        with col_nav5:
            if st.session_state.pagina_atual < total_paginas:
                if st.button("⏭️", help="Última página", use_container_width=True):
                    st.session_state.pagina_atual = total_paginas
                    st.rerun()
        
        # Seletor de página direto
        pagina_selecionada = st.number_input(
            "Ir para página:",
            min_value=1,
            max_value=total_paginas,
            value=st.session_state.pagina_atual,
            key="pagina_seletor",
            label_visibility="collapsed"
        )
        
        if pagina_selecionada != st.session_state.pagina_atual:
            st.session_state.pagina_atual = pagina_selecionada
            st.rerun()
        
        # Calcular índices
        inicio = (st.session_state.pagina_atual - 1) * linhas_por_pagina
        fim = min(inicio + linhas_por_pagina, total_linhas)
        
        st.write(f"**Mostrando registros {inicio + 1} a {fim} de {total_linhas}**")
        
        # Mostrar dados paginados
        altura_pagina = calcular_altura_tabela(linhas_por_pagina)
        
        st.dataframe(
            df.iloc[inicio:fim],
            height=altura_pagina,
            width='stretch',
            hide_index=False,
            use_container_width=True
        )

with tab2:
    # Estatísticas dos dados
    st.subheader("📈 Análise Estatística")
    
    col_stat1, col_stat2 = st.columns(2)
    
    with col_stat1:
        st.write("**Resumo Numérico:**")
        
        # Filtrar apenas colunas numéricas
        colunas_numericas = df.select_dtypes(include=['number']).columns
        
        if len(colunas_numericas) > 0:
            st.dataframe(
                df[colunas_numericas].describe(),
                width='stretch',
                height=300,
                use_container_width=True
            )
        else:
            st.info("ℹ️ Não há colunas numéricas para análise estatística.")
    
    with col_stat2:
        st.write("**Informações das Colunas:**")
        
        info_data = []
        for coluna in df.columns:
            tipo = str(df[coluna].dtype)
            unicos = df[coluna].nunique()
            nulos = df[coluna].isnull().sum()
            percent_preenchido = ((total_linhas - nulos) / total_linhas * 100) if total_linhas > 0 else 0
            
            info_data.append({
                'Coluna': coluna,
                'Tipo': tipo,
                'Valores Únicos': unicos,
                'Valores Nulos': nulos,
                '% Preenchido': f"{percent_preenchido:.1f}%"
            })
        
        info_df = pd.DataFrame(info_data)
        st.dataframe(
            info_df,
            width='stretch',
            height=400,
            use_container_width=True,
            hide_index=True
        )
    
    # Distribuições
    st.subheader("📊 Distribuições")
    
    # Encontrar colunas categóricas para análise
    colunas_categoricas = []
    for coluna in df.columns:
        if df[coluna].dtype == 'object' or df[coluna].nunique() < 20:
            colunas_categoricas.append(coluna)
    
    if len(colunas_categoricas) > 0:
        col_dist1, col_dist2 = st.columns(2)
        
        # Selecionar colunas para análise
        coluna_analise1 = col_dist1.selectbox(
            "Selecione uma coluna para análise:",
            colunas_categoricas,
            key="dist1"
        )
        
        coluna_analise2 = col_dist2.selectbox(
            "Selecione outra coluna para análise:",
            [c for c in colunas_categoricas if c != coluna_analise1],
            key="dist2"
        )
        
        # Mostrar distribuições
        with col_dist1:
            if coluna_analise1 in df.columns:
                contagem1 = df[coluna_analise1].value_counts()
                st.write(f"**Distribuição de {coluna_analise1}:**")
                st.bar_chart(contagem1)
                
                # Tabela de distribuição
                st.dataframe(
                    contagem1.reset_index().rename(
                        columns={'index': coluna_analise1, coluna_analise1: 'Contagem'}
                    ),
                    width='stretch',
                    height=200,
                    use_container_width=True,
                    hide_index=True
                )
        
        with col_dist2:
            if coluna_analise2 in df.columns:
                contagem2 = df[coluna_analise2].value_counts()
                st.write(f"**Distribuição de {coluna_analise2}:**")
                st.bar_chart(contagem2)
                
                # Tabela de distribuição
                st.dataframe(
                    contagem2.reset_index().rename(
                        columns={'index': coluna_analise2, coluna_analise2: 'Contagem'}
                    ),
                    width='stretch',
                    height=200,
                    use_container_width=True,
                    hide_index=True
                )

with tab3:
    # Pesquisa e filtros
    st.subheader("🔍 Pesquisa Avançada")
    
    # Pesquisa por texto
    pesquisa_texto = st.text_input(
        "🔎 Digite um termo para pesquisar em todas as colunas:",
        placeholder="Ex: Campanha, Aprovado, Urgente...",
        key="pesquisa_geral"
    )
    
    if pesquisa_texto:
        # Criar máscara de pesquisa
        mascara = pd.Series(False, index=df.index)
        
        for coluna in df.columns:
            try:
                # Tentar pesquisar em todas as colunas (convertendo para string)
                mascara_coluna = df[coluna].astype(str).str.contains(
                    pesquisa_texto, 
                    case=False, 
                    na=False
                )
                mascara = mascara | mascara_coluna
            except:
                continue
        
        resultados = df[mascara]
        
        if len(resultados) > 0:
            st.success(f"✅ Encontrados {len(resultados)} resultado(s) para '{pesquisa_texto}'")
            
            altura_resultados = calcular_altura_tabela(len(resultados), max_altura=600)
            
            st.dataframe(
                resultados,
                height=altura_resultados,
                width='stretch',
                use_container_width=True
            )
            
            # Exportar resultados
            col_exp1, col_exp2 = st.columns(2)
            
            with col_exp1:
                csv_resultados = resultados.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Exportar Resultados (CSV)",
                    data=csv_resultados,
                    file_name=f"resultados_pesquisa_{pesquisa_texto}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            with col_exp2:
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    resultados.to_excel(writer, index=False, sheet_name='Resultados')
                excel_data = excel_buffer.getvalue()
                
                st.download_button(
                    label="📥 Exportar Resultados (Excel)",
                    data=excel_data,
                    file_name=f"resultados_pesquisa_{pesquisa_texto}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.warning(f"⚠️ Nenhum resultado encontrado para '{pesquisa_texto}'")
    
    # Filtros avançados
    if mostrar_filtros:
        st.subheader("🎛️ Filtros Avançados")
        
        col_filt1, col_filt2, col_filt3 = st.columns(3)
        
        filtros_aplicados = {}
        
        # Filtro por Status
        if 'Status' in df.columns:
            with col_filt1:
                opcoes_status = ['Todos'] + sorted(df['Status'].dropna().unique().tolist())
                status_selecionado = st.selectbox("Filtrar por Status:", opcoes_status)
                if status_selecionado != 'Todos':
                    filtros_aplicados['Status'] = status_selecionado
        
        # Filtro por Prioridade
        if 'Prioridade' in df.columns:
            with col_filt2:
                opcoes_prioridade = ['Todos'] + sorted(df['Prioridade'].dropna().unique().tolist())
                prioridade_selecionada = st.selectbox("Filtrar por Prioridade:", opcoes_prioridade)
                if prioridade_selecionada != 'Todos':
                    filtros_aplicados['Prioridade'] = prioridade_selecionada
        
        # Filtro por Produção
        if 'Produção' in df.columns:
            with col_filt3:
                opcoes_producao = ['Todos'] + sorted(df['Produção'].dropna().unique().tolist())
                producao_selecionada = st.selectbox("Filtrar por Produção:", opcoes_producao)
                if producao_selecionada != 'Todos':
                    filtros_aplicados['Produção'] = producao_selecionada
        
        # Aplicar filtros
        if filtros_aplicados:
            df_filtrado = df.copy()
            
            for coluna, valor in filtros_aplicados.items():
                df_filtrado = df_filtrado[df_filtrado[coluna] == valor]
            
            st.write(f"**📊 Resultados Filtrados: {len(df_filtrado)} de {total_linhas} registros**")
            
            if len(df_filtrado) > 0:
                altura_filtrada = calcular_altura_tabela(len(df_filtrado), max_altura=500)
                
                st.dataframe(
                    df_filtrado,
                    height=altura_filtrada,
                    width='stretch',
                    use_container_width=True
                )
            else:
                st.info("ℹ️ Nenhum registro corresponde aos filtros aplicados.")
            
            # Botão para limpar filtros
            if st.button("🧹 Limpar Todos os Filtros", type="secondary"):
                st.rerun()

with tab4:
    # Exportação de dados
    st.subheader("📤 Exportar Dados")
    
    st.info("""
    **Instruções de Exportação:**
    1. Selecione o formato desejado
    2. Clique no botão de download
    3. O arquivo será baixado automaticamente
    """)
    
    col_exp1, col_exp2, col_exp3, col_exp4 = st.columns(4)
    
    with col_exp1:
        # CSV
        csv_data = df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 CSV",
            data=csv_data,
            file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            help="Baixar em formato CSV (compatível com Excel)",
            use_container_width=True
        )
    
    with col_exp2:
        # Excel
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dados')
            
            # Adicionar aba de metadados
            metadados = pd.DataFrame({
                'Campo': ['Total de Registros', 'Total de Colunas', 'Data de Exportação', 'Fonte'],
                'Valor': [total_linhas, total_colunas, 
                         datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                         f'SharePoint - {SHEET_NAME}']
            })
            metadados.to_excel(writer, index=False, sheet_name='Metadados')
        
        excel_data = excel_buffer.getvalue()
        
        st.download_button(
            label="📥 Excel",
            data=excel_data,
            file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Baixar em formato Excel com múltiplas abas",
            use_container_width=True
        )
    
    with col_exp3:
        # JSON
        json_data = df.to_json(orient='records', force_ascii=False, indent=2)
        st.download_button(
            label="📥 JSON",
            data=json_data,
            file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d')}.json",
            mime="application/json",
            help="Baixar em formato JSON para integrações",
            use_container_width=True
        )
    
    with col_exp4:
        # Clipboard
        if st.button("📋 Copiar para Área de Transferência", use_container_width=True):
            df.to_clipboard(index=False)
            st.success("✅ Dados copiados para área de transferência!")
    
    # Opções de exportação avançada
    with st.expander("⚙️ Opções Avançadas de Exportação"):
        col_adv1, col_adv2 = st.columns(2)
        
        with col_adv1:
            # Filtrar colunas para exportação
            colunas_selecionadas = st.multiselect(
                "Selecionar colunas para exportar:",
                df.columns.tolist(),
                default=df.columns.tolist()[:10] if len(df.columns) > 10 else df.columns.tolist()
            )
        
        with col_adv2:
            # Formato de data
            formato_data = st.selectbox(
                "Formato de datas:",
                ["DD/MM/YYYY", "YYYY-MM-DD", "MM/DD/YYYY"]
            )
        
        # Exportar com filtros
        if colunas_selecionadas:
            df_exportar = df[colunas_selecionadas].copy()
            
            # Aplicar formato de data
            for coluna in df_exportar.columns:
                if pd.api.types.is_datetime64_any_dtype(df_exportar[coluna]):
                    if formato_data == "DD/MM/YYYY":
                        df_exportar[coluna] = df_exportar[coluna].dt.strftime('%d/%m/%Y')
                    elif formato_data == "YYYY-MM-DD":
                        df_exportar[coluna] = df_exportar[coluna].dt.strftime('%Y-%m-%d')
                    else:
                        df_exportar[coluna] = df_exportar[coluna].dt.strftime('%m/%d/%Y')
            
            csv_personalizado = df_exportar.to_csv(index=False, encoding='utf-8-sig')
            
            st.download_button(
                label="📥 Exportar Configuração Personalizada",
                data=csv_personalizado,
                file_name=f"dados_cocred_personalizado_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )

# =========================================================
# 9. ANÁLISE DE PRAZOS E STATUS
# =========================================================

st.divider()
st.header("⏱️ Análise de Prazos e Status")

# Verificar se temos colunas relevantes
colunas_relevantes = []
for col in ['Prazo', 'Prazo (dias)', 'Prazo em dias', 'Dias Restantes', 'Deadline']:
    if col in df.columns:
        colunas_relevantes.append(col)

if colunas_relevantes:
    coluna_prazo = colunas_relevantes[0]
    
    col_analise1, col_analise2 = st.columns(2)
    
    with col_analise1:
        # Classificar prazos
        def classificar_prazo(valor):
            try:
                dias = int(float(str(valor)))
                if dias < 0:
                    return "🟥 Atrasado"
                elif dias == 0:
                    return "🟨 Vence Hoje"
                elif dias <= 3:
                    return "🟧 Urgente (1-3 dias)"
                elif dias <= 7:
                    return "🟦 Próxima Semana"
                else:
                    return "🟩 Em Prazo"
            except:
                return "⚪ Sem Prazo"
        
        df['Classificação Prazo'] = df[coluna_prazo].apply(classificar_prazo)
        
        # Gráfico de prazos
        contagem_prazos = df['Classificação Prazo'].value_counts()
        st.write("**Distribuição de Prazos:**")
        st.bar_chart(contagem_prazos)
    
    with col_analise2:
        # Tabela de resumo
        st.write("**Resumo por Status de Prazo:**")
        
        resumo_prazos = []
        for classificacao, contagem in contagem_prazos.items():
            percentual = (contagem / total_linhas * 100) if total_linhas > 0 else 0
            resumo_prazos.append({
                'Status': classificacao,
                'Quantidade': contagem,
                'Percentual': f"{percentual:.1f}%"
            })
        
        resumo_df = pd.DataFrame(resumo_prazos)
        st.dataframe(
            resumo_df,
            width='stretch',
            height=200,
            use_container_width=True,
            hide_index=True
        )
    
    # Mostrar itens críticos
    itens_criticos = df[df['Classificação Prazo'].str.contains('Atrasado|Vence Hoje|Urgente')]
    
    if len(itens_criticos) > 0:
        st.warning(f"🚨 **{len(itens_criticos)} itens com prazos críticos:**")
        
        colunas_exibir = ['ID', 'Campanha', 'Status', coluna_prazo, 'Classificação Prazo']
        colunas_disponiveis = [c for c in colunas_exibir if c in itens_criticos.columns]
        
        st.dataframe(
            itens_criticos[colunas_disponiveis],
            width='stretch',
            height=min(300, 100 + len(itens_criticos) * 35),
            use_container_width=True
        )
    else:
        st.success("✅ Nenhum prazo crítico identificado!")

# =========================================================
# 10. RESUMO POR CAMPANHA (se houver coluna Campanha)
# =========================================================

if 'Campanha' in df.columns:
    st.divider()
    st.header("📊 Resumo por Campanha")
    
    # Criar resumo
    resumo_campanhas = df.groupby('Campanha').agg({
        'ID': 'count'
    }).rename(columns={'ID': 'Total Jobs'}).reset_index()
    
    # Adicionar contagem por status se disponível
    if 'Status' in df.columns:
        status_pivot = df.pivot_table(
            index='Campanha',
            columns='Status',
            values='ID',
            aggfunc='count',
            fill_value=0
        ).reset_index()
        
        resumo_campanhas = pd.merge(resumo_campanhas, status_pivot, on='Campanha', how='left')
    
    # Ordenar
    resumo_campanhas = resumo_campanhas.sort_values('Total Jobs', ascending=False)
    
    st.dataframe(
        resumo_campanhas,
        width='stretch',
        height=calcular_altura_tabela(len(resumo_campanhas), max_altura=400),
        use_container_width=True
    )
    
    # Gráfico de campanhas
    st.bar_chart(resumo_campanhas.set_index('Campanha')['Total Jobs'])

# =========================================================
# 11. RODAPÉ E INFORMAÇÕES FINAIS
# =========================================================

st.divider()

col_rodape1, col_rodape2, col_rodape3 = st.columns(3)

with col_rodape1:
    st.caption(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    st.caption("Brasília - GMT-3")

with col_rodape2:
    st.caption(f"📊 {total_linhas} registros | {total_colunas} colunas")
    if 'df' in locals() and not df.empty:
        tamanho_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
        st.caption(f"💾 {tamanho_mb:.2f} MB em memória")

with col_rodape3:
    st.caption("🔗 Conectado ao SharePoint Online")
    st.caption(f"📧 {USUARIO_PRINCIPAL}")

# =========================================================
# 12. AUTO-REFRESH OPCIONAL
# =========================================================

# Auto-refresh (opcional - descomente se quiser)
# auto_refresh = st.sidebar.checkbox("🔄 Auto-refresh a cada 5 minutos", value=False)
# 
# if auto_refresh:
#     time.sleep(300)  # 5 minutos
#     st.rerun()

# =========================================================
# FIM DO CÓDIGO
# =========================================================

# Mensagem de sucesso
st.sidebar.success("✅ Aplicação carregada com sucesso!")