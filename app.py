import streamlit as st
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
    
    /* Removidos estilos de métricas customizadas não utilizados */
    
    /* Removidos estilos obsoletos não utilizados */
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
                # Overrides de colunas: Data (E), Proposta (I), Boleto (K)
                try:
                    extractor.set_column_overrides({
                        'data': 'Data de Vencimento Boleto',
                        'proposta': 'Valor Proposta',
                        'boleto': 'Valor do Boleto (R$)'
                    })
                except Exception:
                    pass
                success = extractor.extract_data_from_sheet(sheet_url)
                if success:
                    # Também carrega dados de contratos para projeções
                    extractor.extract_contracts_data(sheet_id)
                    st.session_state['extractor'] = extractor
                    st.session_state['data_loaded'] = True
                    st.success("✅ Dados atualizados!")
                    st.rerun()
                else:
                    # Exibe detalhes do último erro coletado
                    err = getattr(extractor, 'last_error', '')
                    if err:
                        st.error(f"❌ Erro ao carregar dados da planilha: {err}")
                    else:
                        st.error("❌ Erro ao carregar dados da planilha.")
    
    # Carrega dados automaticamente na primeira vez
    if 'data_loaded' not in st.session_state:
        with st.spinner("Carregando dados iniciais..."):
            extractor = SheetsExtractor()
            # Overrides de colunas: Data (E), Proposta (I), Boleto (K)
            try:
                extractor.set_column_overrides({
                    'data': 'Data de Vencimento Boleto',
                    'proposta': 'Valor Proposta',
                    'boleto': 'Valor do Boleto (R$)'
                })
            except Exception:
                pass
            success = extractor.extract_data_from_sheet(sheet_url)
            if success:
                # Também carrega dados de contratos para projeções
                extractor.extract_contracts_data(sheet_id)
                st.session_state['extractor'] = extractor
                st.session_state['data_loaded'] = True
            else:
                err = getattr(extractor, 'last_error', '')
                if err:
                    st.error(f"❌ Erro ao carregar dados iniciais da planilha: {err}")
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
    # Obtém dados detalhados para validação
    detailed_resp = extractor.get_detailed_monthly_data()
    detailed_map = detailed_resp.get('detailed_data', {}) if isinstance(detailed_resp, dict) else {}
    
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
        # Define 2025 como ano padrão se disponível, senão usa o primeiro ano
        default_index = 0
        if "2025" in anos_disponiveis:
            default_index = anos_disponiveis.index("2025")
        ano_selecionado = st.selectbox("📅 Seleção de Ano", anos_disponiveis, index=default_index)
    
    with col2:
        # Métricas apenas para Junho do ano selecionado
        mes_key_junho = f"{ano_selecionado}-06"
        dados_junho = resumo_mensal.get(mes_key_junho)
        if dados_junho:
            total_proposta_junho = dados_junho.get('total_proposta', 0)
            total_boleto_junho = dados_junho.get('total_boleto', 0)
            diferenca_junho = dados_junho.get('total_diferenca', total_boleto_junho - total_proposta_junho)
            
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("💰 Proposta Junho", f"R$ {total_proposta_junho:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            with col_b:
                st.metric("🧾 Boleto Junho", f"R$ {total_boleto_junho:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            with col_c:
                percentual_junho = (diferenca_junho / total_proposta_junho * 100) if total_proposta_junho != 0 else 0
                st.metric("📊 Diferença Junho", f"R$ {diferenca_junho:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), delta=f"{percentual_junho:+.1f}%")
    
    st.markdown("---")
    
    # Calendário de Junho a Dezembro
    st.subheader(f"📅 Calendário {ano_selecionado} (Junho–Dezembro)")
    
    meses_nomes = [
        "Janeiro", "Fevereiro", "Março", "Abril", 
        "Maio", "Junho", "Julho", "Agosto",
        "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    
    # Cria grid com meses de Junho (5) a Dezembro (11)
    visible_months = list(range(5, 12))
    for i in range(0, len(visible_months), 3):
        cols = st.columns(3)
        for j in range(3):
            if i + j < len(visible_months):
                mes_index = visible_months[i + j]
                mes_numero = str(mes_index + 1).zfill(2)
                mes_key = f"{ano_selecionado}-{mes_numero}"
                mes_nome = meses_nomes[mes_index]

                with cols[j]:
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
                        cor = "🔴" if is_projection else "🟢"

                        # Formata valores monetários
                        proposta_fmt = f"R$ {proposta:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        boleto_fmt = f"R$ {boleto:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        diferenca_fmt = f"R$ {diferenca:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

                        with st.container():
                            if is_projection:
                                st.markdown("🔮 **PROJEÇÃO**", help="Dados baseados em contratos")

                            st.markdown(f"### {cor} {mes_nome}")
                            st.caption(f"{count} registros")

                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("💰 Proposta", proposta_fmt)
                            with col2:
                                st.metric("🧾 Boleto", boleto_fmt)

                            if diferenca > 0:
                                st.error(f"📊 Diferença: {diferenca_fmt} (+{percentual:.1f}%)")
                            elif diferenca < 0:
                                st.success(f"📊 Diferença: {diferenca_fmt} ({percentual:.1f}%)")
                            else:
                                st.info(f"📊 Diferença: {diferenca_fmt} ({percentual:.1f}%)")

                            # Removido: bloco de validação cruzada com dados detalhados
                             
                            st.markdown("---")

                        if st.button(f"➕ Detalhar {mes_nome}", key=f"detail_{mes_key}"):
                            display_month_details(extractor, mes_key, mes_nome)
                    else:
                        # Mês sem dados
                        st.info(f"⚪ {mes_nome} — Sem dados")
            else:
                with cols[j]:
                    st.empty()
    


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
        st.metric("🧾 Boleto (Resumo)", f"R$ {total_boleto_resumo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        st.metric("💰 Proposta (Resumo)", f"R$ {total_proposta_resumo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    with col_y:
        st.metric("🧾 Boleto (Detalhado)", f"R$ {total_boleto_detalhado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        st.metric("💰 Proposta (Detalhado)", f"R$ {total_proposta_detalhado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

    # Removido: painel de validação de consistência do mês
    
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
