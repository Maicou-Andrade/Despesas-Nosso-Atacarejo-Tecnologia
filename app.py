import streamlit as st
import pandas as pd
import io
from sheets_extractor import SheetsExtractor
from color_extractor import get_logo_colors

def main():
    st.set_page_config(page_title="Controle de Despesas", page_icon="💰", layout="wide")
    
    # Obter cores da logo
    colors = get_logo_colors()
    primary_color = colors['primary']['hex']
    secondary_color = colors['secondary']['hex']
    
    # CSS personalizado com as cores da logo
    st.markdown(f"""
    <style>
     /* Cores principais da aplicação baseadas na logo */
     :root {{
         --primary-color: {primary_color};
         --secondary-color: {secondary_color};
         --primary-rgb: {colors['primary']['rgb'][0]}, {colors['primary']['rgb'][1]}, {colors['primary']['rgb'][2]};
         --secondary-rgb: {colors['secondary']['rgb'][0]}, {colors['secondary']['rgb'][1]}, {colors['secondary']['rgb'][2]};
     }}
     
     /* Fundo preto e contraste geral */
     html, body, .stApp {{
         background-color: #000000 !important;
         color: #f0f0f0 !important;
     }}
     
     /* Deixar containers principais transparentes para o fundo aparecer */
     .main .block-container {{
         background: transparent !important;
     }}

     /* Ajustar cores do st.metric para fundo escuro */
     div[data-testid="stMetricValue"] {{
         color: #f0f0f0 !important;
     }}
     div[data-testid="stMetricLabel"] {{
         color: #cccccc !important;
     }}
     div[data-testid="stMetricDelta"] {{
         color: #cccccc !important;
     }}
     
     /* Estilo do título principal */
     .main-title {{
         background: linear-gradient(90deg, {primary_color}, {secondary_color});
         -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 3rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 2rem;
    }}
    
    /* Botões personalizados */
    .stButton > button {{
        background: linear-gradient(45deg, {primary_color}, {secondary_color});
        color: white;
        border: none;
        border-radius: 10px;
        font-weight: bold;
        transition: all 0.3s ease;
    }}
    
    .stButton > button:hover {{
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(var(--primary-rgb), 0.3);
    }}
    
    /* Ajuste da fonte dos valores nos metrics para evitar truncamento */
    [data-testid="metric-container"] [data-testid="metric-value"] {{
        font-size: 1rem !important;
        white-space: nowrap !important;
        overflow: visible !important;
        text-overflow: clip !important;
    }}
    
    [data-testid="metric-container"] {{
        min-width: 180px !important;
        width: auto !important;
    }}
    
    /* Ajuste específico para colunas com metrics */
    .stColumn {{
        min-width: 200px !important;
    }}
    
    /* Métricas personalizadas */
    .metric-card {{
        background: linear-gradient(135deg, rgba(var(--primary-rgb), 0.1), rgba(var(--secondary-rgb), 0.1));
        border: 2px solid {primary_color};
        border-radius: 15px;
        padding: 1rem;
        text-align: center;
        margin: 0.5rem 0;
    }}
    
    .metric-value {{
        font-size: 2rem;
        font-weight: bold;
        color: {primary_color};
    }}
    
    .metric-label {{
        font-size: 1rem;
        color: {primary_color};
        font-weight: 600;
    }}
    
    /* Calendário personalizado */
    .calendar-month {{
        background: linear-gradient(135deg, rgba(var(--primary-rgb), 0.05), rgba(var(--secondary-rgb), 0.05));
        border: 1px solid {primary_color};
        border-radius: 10px;
        padding: 1rem;
        margin: 0.5rem;
        transition: all 0.3s ease;
    }}
    
    .calendar-month:hover {{
        transform: translateY(-3px);
        box-shadow: 0 5px 20px rgba(var(--primary-rgb), 0.2);
    }}
    
    .month-title {{
        color: {primary_color};
        font-weight: bold;
        font-size: 1.1rem;
        text-align: center;
        margin-bottom: 0.5rem;
    }}
    
    .positive-diff {{
        color: #28a745;
        font-weight: bold;
    }}
    
    .negative-diff {{
        color: #dc3545;
        font-weight: bold;
    }}
    
    .neutral-diff {{
        color: #6c757d;
        font-weight: bold;
    }}
    
    /* Seletor de ano */
    .year-selector {{
        background: linear-gradient(45deg, {primary_color}, {secondary_color});
        color: white;
        border-radius: 10px;
        padding: 0.5rem;
        text-align: center;
        font-weight: bold;
        margin: 1rem 0;
    }}
    </style>
    """, unsafe_allow_html=True)
    
    # Logo e título
    col_logo, col_title = st.columns([1, 4])
    with col_logo:
        try:
            st.image("LOGO.jpg", width=120)
        except:
            st.markdown("🏢", unsafe_allow_html=True)
    
    with col_title:
        st.markdown('<h1 class="main-title">💰 Controle de Despesas</h1>', unsafe_allow_html=True)
    
    # URL e ID padrão da planilha
    default_url = "https://docs.google.com/spreadsheets/d/1mKtICE1MaSKr60CcDHmPKBm7z35aiE1SsJ4EUw04CgQ/edit?gid=467594733#gid=467594733"
    default_sheet_id = "1mKtICE1MaSKr60CcDHmPKBm7z35aiE1SsJ4EUw04CgQ"

    # Permite configurar via segredos da Streamlit Cloud
    sheet_url = default_url
    sheet_id = default_sheet_id
    try:
        if hasattr(st, "secrets"):
            if "SHEET_URL" in st.secrets:
                sheet_url = st.secrets["SHEET_URL"]
            if "SHEET_ID" in st.secrets:
                sheet_id = st.secrets["SHEET_ID"]
    except Exception:
        pass
    
    # Botão de refresh simples
    col1, col2, col3 = st.columns([1, 1, 4])
    with col2:
        if st.button("🔄 Atualizar Dados", type="primary", use_container_width=True):
            with st.spinner("Atualizando dados..."):
                extractor = SheetsExtractor()
                success = extractor.extract_data_from_sheet(sheet_url)
                if success:
                    # Também carrega dados de contratos para projeções
                    extractor.extract_contracts_data(sheet_id)
                    st.session_state['extractor'] = extractor
                    st.session_state['data_loaded'] = True
                    st.success("✅ Dados atualizados!")
                    st.rerun()
                else:
                    st.error("❌ Erro ao carregar dados da planilha.")
    
    # Carrega dados automaticamente na primeira vez
    if 'data_loaded' not in st.session_state:
        with st.spinner("Carregando dados iniciais..."):
            extractor = SheetsExtractor()
            success = extractor.extract_data_from_sheet(sheet_url)
            if success:
                # Também carrega dados de contratos para projeções
                extractor.extract_contracts_data(sheet_id)
                st.session_state['extractor'] = extractor
                st.session_state['data_loaded'] = True
            else:
                st.error("❌ Erro ao carregar dados iniciais da planilha.")
    
    # Exibe a visualização do calendário
    if 'data_loaded' in st.session_state and st.session_state['data_loaded']:
        extractor = st.session_state['extractor']
        display_calendar_view(extractor)

def display_calendar_view(extractor):
    """Exibe visualização em formato calendário com comparação entre Proposta e Boleto"""
    st.header("📅 Calendário Financeiro - Proposta vs Boleto")
    
    # Obtém o resumo mensal incluindo projeções automáticas para meses futuros ausentes
    monthly_data = extractor.get_monthly_data_with_auto_projections()
    
    if 'error' in monthly_data:
        st.error(f"❌ {monthly_data['error']}")
        return
    
    resumo_mensal = monthly_data.get('resumo_mensal', {})
    
    if not resumo_mensal:
        st.warning("⚠️ Nenhum dado encontrado")
        return
    
    # Extrai anos disponíveis
    anos_disponiveis = sorted(list(set([month_year.split('-')[0] for month_year in resumo_mensal.keys()])))
    
    if not anos_disponiveis:
        st.warning("⚠️ Nenhum ano encontrado nos dados")
        return
    
    # Seletor de ano
    col1, col2 = st.columns([1, 3])
    with col1:
        st.markdown('<div class="year-selector">📅 Seleção de Ano</div>', unsafe_allow_html=True)
        # Define 2025 como ano padrão se disponível, senão usa o primeiro ano
        default_index = 0
        if "2025" in anos_disponiveis:
            default_index = anos_disponiveis.index("2025")
        ano_selecionado = st.selectbox("Selecione o Ano:", anos_disponiveis, index=default_index, label_visibility="collapsed")
    
    with col2:
        # Métricas do ano selecionado
        dados_ano = {k: v for k, v in resumo_mensal.items() if k.startswith(ano_selecionado)}
        if dados_ano:
            total_proposta_ano = sum(month_data.get('total_proposta', 0) for month_data in dados_ano.values())
            total_boleto_ano = sum(month_data.get('total_boleto', 0) for month_data in dados_ano.values())
            diferenca_ano = sum(month_data.get('total_diferenca', 0) for month_data in dados_ano.values())
            
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-value">R$ {total_proposta_ano:,.2f}</div>
                    <div class="metric-label">💰 Proposta Total</div>
                </div>
                """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)
            with col_b:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-value">R$ {total_boleto_ano:,.2f}</div>
                    <div class="metric-label">🧾 Boleto Total</div>
                </div>
                """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)
            with col_c:
                percentual_ano = (diferenca_ano / total_proposta_ano * 100) if total_proposta_ano != 0 else 0
                diff_class = "positive-diff" if diferenca_ano > 0 else "negative-diff" if diferenca_ano < 0 else "neutral-diff"
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-value {diff_class}">R$ {diferenca_ano:,.2f}</div>
                    <div class="metric-label">📊 Diferença ({percentual_ano:+.1f}%)</div>
                </div>
                """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Grid de calendário (3x4 para 12 meses)
    st.subheader(f"📅 Calendário {ano_selecionado}")
    
    meses_nomes = [
        "Janeiro", "Fevereiro", "Março", "Abril", 
        "Maio", "Junho", "Julho", "Agosto",
        "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    
    # Cria grid 3x4
    for linha in range(4):
        cols = st.columns(3)
        for coluna in range(3):
            mes_index = linha * 3 + coluna
            if mes_index < 12:
                mes_numero = str(mes_index + 1).zfill(2)
                mes_key = f"{ano_selecionado}-{mes_numero}"
                mes_nome = meses_nomes[mes_index]
                
                with cols[coluna]:
                    if mes_key in resumo_mensal:
                        # Mês com dados
                        month_data = resumo_mensal[mes_key]
                        proposta = month_data.get('total_proposta', 0)
                        boleto = month_data.get('total_boleto', 0)
                        diferenca = month_data.get('total_diferenca', boleto - proposta)
                        percentual = month_data.get('diferenca_percentual_media', 0)
                        count = month_data['count']
                        is_projection = month_data.get('is_projection', False)
                        
                        # Cor baseada no tipo de dados
                        if is_projection:
                            cor = "🔴"  # Vermelho para projeções
                        else:
                            cor = "🟢"  # Verde para dados reais da aba Despesas
                        
                        # Indicador de projeção
                        projection_indicator = ""
                        if is_projection:
                            projection_indicator = '<div style="background-color: #ff4444; color: white; font-size: 10px; font-weight: bold; padding: 2px 6px; border-radius: 4px; margin-bottom: 5px; text-align: center;">PROJEÇÃO</div>'
                        
                        # Formata valores monetários
                        proposta_fmt = f"R$ {proposta:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        boleto_fmt = f"R$ {boleto:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        diferenca_fmt = f"R$ {diferenca:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        
                        # Usa container com borda para simular o card
                        with st.container():
                            # Indicador de projeção se necessário
                            if is_projection:
                                st.markdown("🔮 **PROJEÇÃO**", help="Dados baseados em contratos")
                            
                            # Título do mês com emoji indicador
                            st.markdown(f"### {cor} {mes_nome}")
                            st.caption(f"{count} registros")
                            
                            # Métricas em colunas
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("💰 Proposta", proposta_fmt)
                            with col2:
                                st.metric("🧾 Boleto", boleto_fmt)
                            
                            # Diferença com cor
                            if diferenca > 0:
                                st.error(f"📊 Diferença: {diferenca_fmt} (+{percentual:.1f}%)")
                            elif diferenca < 0:
                                st.success(f"📊 Diferença: {diferenca_fmt} ({percentual:.1f}%)")
                            else:
                                st.info(f"📊 Diferença: {diferenca_fmt} ({percentual:.1f}%)")
                            
                            st.markdown("---")
                        
                        # Botão de detalhamento
                        if st.button(f"➕ Detalhar {mes_nome}", key=f"detail_{mes_key}"):
                            display_month_details(extractor, mes_key, mes_nome)
                    else:
                        # Mês sem dados
                        st.markdown(f"""
                        <div style="border: 2px dashed #ccc; border-radius: 10px; padding: 15px; margin: 5px; background-color: #f5f5f5; opacity: 0.6;">
                            <h4 style="margin: 0; color: #999;">⚪ {mes_nome}</h4>
                            <p style="margin: 5px 0; font-size: 12px; color: #999;">Sem dados</p>
                        </div>
                        """, unsafe_allow_html=True)
    

    


def display_summary(processed_data):
    """Exibe resumo das despesas"""
    st.header("📊 Resumo das Despesas")
    
    summary = processed_data['resumo']
    
    # Métricas principais
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📋 Total de Registros", processed_data.get('total_registros', 0))
    
    with col2:
        st.metric("✅ Despesas Válidas", summary.get('total_despesas', 0))
    
    with col3:
        st.metric("💰 Valor Total", f"R$ {summary.get('valor_total', 0):.2f}")
    
    with col4:
        st.metric("📊 Valor Médio", f"R$ {summary.get('valor_medio', 0):.2f}")
    
    # Informações adicionais
    if summary.get('total_despesas', 0) > 0:
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric("📉 Menor Valor", f"R$ {summary.get('valor_minimo', 0):.2f}")
        
        with col2:
            st.metric("📈 Maior Valor", f"R$ {summary.get('valor_maximo', 0):.2f}")
    
    # Informações sobre colunas
    st.markdown("---")
    st.subheader("🔍 Colunas Identificadas")
    
    columns_info = processed_data['colunas_identificadas']
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**💰 Colunas de Despesa:**")
        if columns_info['valores']:
            for col in columns_info['valores']:
                st.markdown(f"• {col}")
        else:
            st.markdown("• Nenhuma identificada")
    
    with col2:
        st.markdown("**📅 Colunas de Data:**")
        if columns_info['datas']:
            for col in columns_info['datas']:
                st.markdown(f"• {col}")
        else:
            st.markdown("• Nenhuma identificada")
    
    with col3:
        st.markdown("**📝 Colunas de Descrição:**")
        if columns_info['descricoes']:
            for col in columns_info['descricoes']:
                st.markdown(f"• {col}")
        else:
            st.markdown("• Nenhuma identificada")

def display_data_preview(processed_data):
    """Exibe prévia dos dados"""
    st.header("👀 Prévia dos Dados")
    
    # Usa os dados brutos do primeiro registro ou dados originais
    raw_data = []
    if 'despesas' in processed_data and processed_data['despesas']:
        # Extrai dados brutos das despesas processadas
        for despesa in processed_data['despesas']:
            if 'dados_brutos' in despesa:
                raw_data.append(despesa['dados_brutos'])
    
    # Se não tem dados brutos, usa os dados originais do extractor
    if not raw_data:
        extractor = SheetsExtractor()
        raw_data = extractor.data if hasattr(extractor, 'data') else []
    
    if raw_data:
        st.markdown(f"**Mostrando os primeiros 10 registros de {len(raw_data)} total:**")
        
        # Converte para formato de tabela
        preview_data = raw_data[:10]
        st.dataframe(preview_data, use_container_width=True)
        
        # Informações sobre as colunas
        st.markdown("---")
        st.subheader("📋 Informações das Colunas")
        
        columns = list(raw_data[0].keys()) if raw_data else []
        st.markdown(f"**Total de colunas:** {len(columns)}")
        
        # Mostra todas as colunas
        cols_per_row = 3
        for i in range(0, len(columns), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, col_name in enumerate(columns[i:i+cols_per_row]):
                with cols[j]:
                    st.markdown(f"• {col_name}")
    
    else:
        st.warning("⚠️ Nenhum dado disponível para prévia")

def analyze_expenses(processed_data):
    """Analisa e exibe gráficos das despesas"""
    st.header("📈 Análise das Despesas")
    
    processed_expenses = processed_data['despesas']
    valid_expenses = [exp for exp in processed_expenses if exp.get('valor') is not None]
    
    if not valid_expenses:
        st.warning("⚠️ Nenhuma despesa válida encontrada para análise")
        return
    
    # Estatísticas básicas
    st.subheader("📊 Estatísticas")
    
    values = [exp['valor'] for exp in valid_expenses]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**💰 Distribuição de Valores:**")
        st.markdown(f"• Menor: R$ {min(values):.2f}")
        st.markdown(f"• Maior: R$ {max(values):.2f}")
        st.markdown(f"• Média: R$ {sum(values)/len(values):.2f}")
        st.markdown(f"• Total: R$ {sum(values):.2f}")
    
    with col2:
        st.markdown("**📈 Análise Quantitativa:**")
        st.markdown(f"• Total de despesas: {len(valid_expenses)}")
        st.markdown(f"• Registros válidos: {len(valid_expenses)}")
        if len(processed_expenses) > 0:
            st.markdown(f"• Taxa de sucesso: {len(valid_expenses)/len(processed_expenses)*100:.1f}%")
        else:
            st.markdown("• Taxa de sucesso: 0%")
    
    # Lista das despesas processadas
    st.markdown("---")
    st.subheader("📋 Despesas Processadas")
    
    # Cria tabela com despesas válidas
    expense_table = []
    for exp in valid_expenses:
        expense_table.append({
            'Valor': f"R$ {exp['valor']:.2f}",
            'Data': exp.get('data', 'N/A'),
            'Descrição': exp.get('descricao', 'N/A')
        })
    
    if expense_table:
        st.dataframe(expense_table, use_container_width=True)

def show_raw_data(processed_data):
    """Exibe dados brutos com opção de download"""
    st.header("📋 Dados Brutos")
    
    # Extrai dados brutos das despesas processadas
    raw_data = []
    if 'despesas' in processed_data and processed_data['despesas']:
        for despesa in processed_data['despesas']:
            if 'dados_brutos' in despesa:
                raw_data.append(despesa['dados_brutos'])
    
    if raw_data:
        st.markdown(f"**Total de registros:** {len(raw_data)}")
        
        # Exibe os dados
        st.dataframe(raw_data, use_container_width=True)
        
        # Opção de download
        st.markdown("---")
        st.subheader("💾 Download")
        
        # Converte para CSV
        csv_buffer = io.StringIO()
        if raw_data:
            import csv
            fieldnames = raw_data[0].keys()
            writer = csv.DictWriter(csv_buffer, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(raw_data)
        
        csv_data = csv_buffer.getvalue()
        
        st.download_button(
            label="📥 Baixar dados em CSV",
            data=csv_data,
            file_name="despesas_extraidas.csv",
            mime="text/csv"
        )
        
        # Informações sobre o arquivo
        st.markdown(f"**Tamanho do arquivo:** ~{len(csv_data)} caracteres")
        st.markdown(f"**Colunas:** {len(raw_data[0].keys()) if raw_data else 0}")
    
    else:
        st.warning("⚠️ Nenhum dado bruto disponível")

def display_month_details(extractor, mes_key, mes_nome):
    """Exibe detalhes do mês com hierarquia: Tipo -> Empresa -> Propostas"""
    st.markdown("---")
    st.subheader(f"📊 Detalhes de {mes_nome}")
    
    # Obtém dados detalhados
    detailed_data = extractor.get_detailed_monthly_data()
    
    if 'error' in detailed_data:
        st.error(f"❌ {detailed_data['error']}")
        return
    
    month_data = detailed_data.get('detailed_data', {}).get(mes_key)
    if not month_data:
        st.warning(f"⚠️ Nenhum dado detalhado encontrado para {mes_nome}")
        return
    
    # Exibe totais do mês
    st.markdown(f"""
    **📅 Resumo de {mes_nome}:**
    - 📝 **Total de registros:** {month_data['total_registros']}
    - 💰 **Proposta Total:** R$ {month_data['total_proposta']:,.2f}
    - 🧾 **Boleto Total:** R$ {month_data['total_boleto']:,.2f}
    - 📊 **Diferença:** R$ {month_data['total_diferenca']:,.2f}
    """.replace(',', 'X').replace('.', ',').replace('X', '.'))

    # Monta lista de propostas a partir da estrutura hierárquica (tipo -> empresa -> propostas)
    items_detalhado = []
    for tipo_nome, tipo_data in month_data.get('tipos', {}).items():
        for empresa_nome, empresa_data in tipo_data.get('empresas', {}).items():
            for prop in empresa_data.get('propostas', []):
                items_detalhado.append(prop)

    # Validação cruzada: compara totais do resumo mensal vs detalhado
    resumo_all = extractor.get_monthly_summary_by_columns()
    resumo_mes = resumo_all.get('resumo_mensal', {}).get(mes_key, {})
    total_boleto_resumo = resumo_mes.get('total_boleto', 0)
    total_proposta_resumo = resumo_mes.get('total_proposta', 0)
    total_boleto_detalhado = sum(it.get('valor_boleto', 0) for it in items_detalhado)
    total_proposta_detalhado = sum(it.get('valor_proposta', 0) for it in items_detalhado)

    diff_boleto = total_boleto_resumo - total_boleto_detalhado
    diff_proposta = total_proposta_resumo - total_proposta_detalhado

    col_x, col_y = st.columns(2)
    with col_x:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">R$ {total_boleto_resumo:,.2f}</div>
            <div class="metric-label">🧾 Boleto (Resumo)</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">R$ {total_proposta_resumo:,.2f}</div>
            <div class="metric-label">💰 Proposta (Resumo)</div>
        </div>
        """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)
    with col_y:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">R$ {total_boleto_detalhado:,.2f}</div>
            <div class="metric-label">🧾 Boleto (Detalhado)</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">R$ {total_proposta_detalhado:,.2f}</div>
            <div class="metric-label">💰 Proposta (Detalhado)</div>
        </div>
        """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)

    # Painel de consistência
    if abs(diff_boleto) > 0.01 or abs(diff_proposta) > 0.01:
        with st.expander("🔎 Validação de consistência do mês"):
            st.markdown(f"- 🧾 Diferença Boleto (Resumo - Detalhado): R$ {diff_boleto:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            st.markdown(f"- 💰 Diferença Proposta (Resumo - Detalhado): R$ {diff_proposta:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            st.markdown(f"- 📄 Itens detalhados: {len(items_detalhado)} | Registros no resumo: {resumo_mes.get('count', 0)}")

            # Casos suspeitos
            boleto_only = [it for it in items_detalhado if it.get('valor_boleto', 0) > 0 and (it.get('valor_proposta', 0) == 0)]
            boleto_gt = [it for it in items_detalhado if it.get('valor_boleto', 0) > it.get('valor_proposta', 0)]
            invalid_date = [it for it in items_detalhado if extractor._extract_month_year_from_date(it.get('data', '')) != mes_key]
            sum_boleto_only = sum(it.get('valor_boleto', 0) for it in boleto_only)
            sum_diff_positive = sum((it.get('valor_boleto', 0) - it.get('valor_proposta', 0)) for it in boleto_gt)
            
            st.markdown(f"- 📌 Boleto SEM proposta: {len(boleto_only)} | Soma: R$ {sum_boleto_only:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            st.markdown(f"- 📌 Boleto > Proposta: {len(boleto_gt)} | Soma diferença: R$ {sum_diff_positive:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            st.markdown(f"- 📆 Datas fora do mês {mes_nome}: {len(invalid_date)}")

            # Top 10 maiores diferenças
            top_diffs = sorted(boleto_gt, key=lambda it: it.get('valor_boleto', 0) - it.get('valor_proposta', 0), reverse=True)[:10]
            if top_diffs:
                st.markdown("---")
                st.markdown("**Top 10 diferenças (R$ boleto - R$ proposta):**")
                for it in top_diffs:
                    diff_val = (it.get('valor_boleto', 0) - it.get('valor_proposta', 0))
                    data_str = it.get('data', 'N/A')
                    desc = it.get('proposta', it.get('descricao', '')) or ''
                    st.markdown(f"• {data_str} — {desc} — Dif: R$ {diff_val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    
    # Detalhamento por tipo
    tipos = month_data.get('tipos', {})
    for tipo_nome, tipo_data in tipos.items():
        with st.expander(f"🔍 {tipo_nome} ({tipo_data['total_registros']} registros) - R$ {tipo_data['total_proposta']:,.2f} (Proposta) | R$ {tipo_data['total_boleto']:,.2f} (Boleto)".replace(',', 'X').replace('.', ',').replace('X', '.')):
            # Totais do tipo
            st.markdown(f"""
            **📋 Resumo {tipo_nome}:**
            - 📝 **Total de registros:** {tipo_data['total_registros']}
            - 💰 **Proposta:** R$ {tipo_data['total_proposta']:,.2f}
            - 🧾 **Boleto:** R$ {tipo_data['total_boleto']:,.2f}
            - 📊 **Diferença:** R$ {tipo_data['total_diferenca']:,.2f}
            """.replace(',', 'X').replace('.', ',').replace('X', '.'))
            
            # Detalhamento por empresa (já expandido)
            empresas = tipo_data.get('empresas', {})
            for empresa_nome, empresa_data in empresas.items():
                empresa_nome_display = empresa_nome if empresa_nome else "Empresa não informada"
                
                # Mostra informações da empresa diretamente
                st.markdown(f"**🏢 {empresa_nome_display}** ({empresa_data['total_registros']} registros)")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("💰 Proposta", f"R$ {empresa_data['total_proposta']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                with col2:
                    st.metric("🧾 Boleto", f"R$ {empresa_data['total_boleto']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                with col3:
                    st.metric("📊 Diferença", f"R$ {empresa_data['total_diferenca']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                
                # Botão para ver propostas individuais usando fragment
                propostas_key = f"show_propostas_{mes_key}_{tipo_nome}_{empresa_nome}".replace(" ", "_").replace("/", "_")
                
                # Função fragment para isolar re-rendering
                @st.fragment
                def render_propostas_section():
                    # Inicializa session state se não existir
                    if propostas_key not in st.session_state:
                        st.session_state[propostas_key] = False
                    
                    # Botão toggle
                    button_text = "➖ Ocultar propostas individuais" if st.session_state[propostas_key] else f"➕ Ver propostas individuais de {empresa_nome_display}"
                    
                    if st.button(button_text, key=f"btn_propostas_{mes_key}_{tipo_nome}_{empresa_nome}".replace(" ", "_").replace("/", "_")):
                        st.session_state[propostas_key] = not st.session_state[propostas_key]
                        st.rerun(scope="fragment")
                    
                    # Mostra propostas se ativado
                    if st.session_state[propostas_key]:
                        st.markdown("### 📋 Propostas Individuais")
                        display_individual_proposals(empresa_data['propostas'])
                
                # Chama o fragment
                render_propostas_section()
                
                st.markdown("---")



def display_individual_proposals(propostas):
    """Exibe propostas individuais com colunas C,G,H,E,I,J,K"""
    
    if not propostas:
        st.warning("⚠️ Nenhuma proposta encontrada")
        return
    
    st.info(f"📊 Encontradas {len(propostas)} propostas")
    
    # Cria tabela com as propostas
    proposals_table = []
    for i, proposta in enumerate(propostas):
        proposals_table.append({
            '#': i + 1,
            'Empresa (C)': proposta.get('empresa', ''),
            'Coluna G': proposta.get('col_g', ''),
            'Coluna H': proposta.get('col_h', ''),
            'Data (E)': proposta.get('data', ''),
            'Valor Proposta (I)': f"R$ {proposta.get('valor_proposta', 0):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
            'Coluna J': proposta.get('col_j', ''),
            'Valor Boleto (K)': f"R$ {proposta.get('valor_boleto', 0):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
            'Diferença': f"R$ {proposta.get('diferenca', 0):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        })
    
    # Exibe a tabela
    st.dataframe(proposals_table, use_container_width=True, hide_index=True)
    
    # Estatísticas
    total_propostas = len(propostas)
    total_proposta_valor = sum(p.get('valor_proposta', 0) for p in propostas)
    total_boleto_valor = sum(p.get('valor_boleto', 0) for p in propostas)
    total_diferenca = sum(p.get('diferenca', 0) for p in propostas)
    
    # Exibe estatísticas em colunas
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📝 Total", total_propostas)
    with col2:
        st.metric("💰 Propostas", f"R$ {total_proposta_valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    with col3:
        st.metric("🧾 Boletos", f"R$ {total_boleto_valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    with col4:
        st.metric("📊 Diferença", f"R$ {total_diferenca:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

if __name__ == "__main__":
    main()