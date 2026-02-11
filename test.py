import streamlit as st
import pandas as pd
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Dashboard COCRED - TESTE",
    layout="wide"
)

# BANNER DE VERIFICAÇÃO
st.markdown("""
<div style="background: linear-gradient(90deg, #FF6B6B 0%, #4ECDC4 100%); 
            padding: 20px; 
            border-radius: 10px; 
            color: white; 
            text-align: center; 
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
    <h1 style="margin: 0; font-size: 28px;">🚀 TESTE DE CONEXÃO</h1>
    <p style="margin: 5px 0 0 0; font-size: 16px; opacity: 0.9;">
        App funcionando! | Streamlit v""" + st.__version__ + """ | """ + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + """
    </p>
</div>
""", unsafe_allow_html=True)

st.title("📊 Dashboard de Campanhas - SICOOB COCRED")
st.subheader("🔗 Conectado ao Excel Online do SharePoint")

# Criar dados de teste
st.write("### 📋 Dados de Teste")

dados_teste = {
    'ID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
    'Campanha': ['Crédito Automático', 'Consórcios', 'Crédito PJ', 'Investimentos', 'Conta Digital',
                 'TVs Internas', 'Marketing Digital', 'Redes Sociais', 'Email Marketing', 'Site'],
    'Status': ['Aprovado', 'Em Produção', 'Aguardando', 'Aprovado', 'Em Produção',
               'Aguardando', 'Concluído', 'Em Produção', 'Aprovado', 'Em Andamento'],
    'Prioridade': ['Alta', 'Média', 'Alta', 'Baixa', 'Média', 'Alta', 'Baixa', 'Média', 'Alta', 'Média'],
    'Produção': ['Cocred', 'Ideatore', 'Cocred', 'Ideatore', 'Cocred',
                 'Ideatore', 'Cocred', 'Ideatore', 'Cocred', 'Ideatore'],
    'Data Solicitação': pd.date_range(start='2024-01-01', periods=10, freq='D'),
    'Prazo (dias)': [5, 3, 10, 2, 7, 14, 0, 5, 3, 8]
}

df = pd.DataFrame(dados_teste)

# Mostrar métricas
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total Registros", len(df))
with col2:
    st.metric("Campanhas", df['Campanha'].nunique())
with col3:
    st.metric("Em Produção", len(df[df['Status'] == 'Em Produção']))
with col4:
    st.metric("Prazos Críticos", len(df[df['Prazo (dias)'] <= 3]))

# Mostrar tabela
st.write("### 📊 Tabela de Dados")
st.dataframe(df)

# Sidebar
with st.sidebar:
    st.title("⚙️ Controles")
    st.write(f"**Versão:** {st.__version__}")
    st.write(f"**Data:** {datetime.now().strftime('%d/%m/%Y')}")
    
    if st.button("🔄 Atualizar Dados", type="primary", use_container_width=True):
        st.rerun()
    
    st.divider()
    st.write("**Próximos passos:**")
    st.write("1. Configurar conexão SharePoint")
    st.write("2. Carregar dados reais")
    st.write("3. Adicionar análise")

# Rodapé
st.divider()
st.caption(f"✅ App testado com sucesso | {datetime.now().strftime('%H:%M:%S')} | Cristini Cordesco")