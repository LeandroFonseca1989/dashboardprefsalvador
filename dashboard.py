"""
Dashboard de Produtividade - USF São Cristóvão
"""
import pandas as pd
import streamlit as st
import altair as alt
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Dashboard de Produtividade - USF São Cristóvão",
    page_icon="📊",
    layout="wide"
)

# Header com botão de recarregar
col_header1, col_header2 = st.columns([10, 1])
with col_header1:
    st.title("📊 Dashboard de Produtividade - USF São Cristóvão")
with col_header2:
    if st.button("🔄 Recarregar Arquivo", key="btn_recarregar", help="Clique para carregar um novo arquivo"):
        # Limpar session_state relacionado ao arquivo
        if 'arquivo_carregado' in st.session_state:
            del st.session_state.arquivo_carregado
        if 'arquivo_nome' in st.session_state:
            del st.session_state.arquivo_nome
        st.rerun()

st.markdown("---")

# Função para carregar e consolidar dados e extrair mês
@st.cache_data
def carregar_dados(uploaded_file):
    """
    Carrega e consolida dados de todas as abas do Excel
    """
    try:
        # Ler todas as abas
        xls = pd.ExcelFile(uploaded_file)
        
        # Lista de nomes de meses em português
        meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        
        # Identificar aba de mês e extrair o mês
        mes_do_arquivo = None
        abas_conhecidas = ['Consolidado']
        
        for aba in xls.sheet_names:
            # Se não começa com "Dia" e não é "Consolidado", pode ser uma aba de mês
            if not aba.startswith("Dia") and aba not in abas_conhecidas:
                try:
                    # Ler apenas a primeira linha da coluna A
                    df_mes = pd.read_excel(xls, sheet_name=aba, nrows=1, usecols=[0], header=None)
                    if len(df_mes) > 0:
                        valor_celula_a1 = df_mes.iloc[0, 0]
                        
                        # Tentar extrair o mês
                        if pd.notna(valor_celula_a1):
                            # Se for uma data (Timestamp)
                            if isinstance(valor_celula_a1, pd.Timestamp) or hasattr(valor_celula_a1, 'month'):
                                mes_numero = valor_celula_a1.month
                                mes_do_arquivo = meses[mes_numero - 1]
                            else:
                                # Tentar converter string para data
                                try:
                                    data = pd.to_datetime(valor_celula_a1)
                                    mes_numero = data.month
                                    mes_do_arquivo = meses[mes_numero - 1]
                                except:
                                    # Verificar se é um mês em texto
                                    valor_str = str(valor_celula_a1).strip()
                                    if valor_str in meses:
                                        mes_do_arquivo = valor_str
                except:
                    pass  # Ignorar erros na leitura da aba de mês
        
        dados_consolidados = []
        
        for aba in xls.sheet_names:
            # Ignorar abas que não são de dias (ex: "Consolidado")
            if aba.startswith("Dia"):
                df = pd.read_excel(xls, sheet_name=aba)
                
                # Adicionar coluna Dia
                df['Dia'] = aba
                
                # Adicionar coluna Mês se foi identificado
                if mes_do_arquivo:
                    df['Mês'] = mes_do_arquivo
                else:
                    df['Mês'] = 'Não informado'
                
                # Remover coluna Unnamed: 0 se existir
                if 'Unnamed: 0' in df.columns:
                    df = df.drop(columns=['Unnamed: 0'])
                
                dados_consolidados.append(df)
        
        # Concatenar todos os DataFrames
        df_consolidado = pd.concat(dados_consolidados, ignore_index=True)
        
        # Limpar dados: tratar valores NaN
        df_consolidado['Profissional'] = df_consolidado['Profissional'].fillna('Não informado')
        df_consolidado['Especialidade'] = df_consolidado['Especialidade'].fillna('Não informado')
        
        # Consolidar status: criar nova coluna Status_Consolidado
        def consolidar_status(status):
            if pd.isna(status):
                return 'Não informado'
            status_upper = str(status).upper()
            
            # Status que devem ser consolidados em "Atendimento realizado"
            if status_upper in ['AGENDADO', 'AGUARDANDO ATENDIMENTO', 'ATENDIMENTO FINALIZADO', 
                                'REALIZANDO PROCEDIMENTO/EXAME']:
                return 'Atendimento realizado'
            # Status que permanecem separados
            elif status_upper == 'EVADIDO':
                return 'Evadido'
            elif status_upper == 'FALTOSO':
                return 'Faltoso'
            else:
                return 'Atendimento realizado'  # Por padrão, outros status também consolidados
        
        df_consolidado['Status_Consolidado'] = df_consolidado['Status'].apply(consolidar_status)
        
        return df_consolidado, xls.sheet_names
    
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {str(e)}")
        return None, None

# Controlar se mostra o upload ou não
mostrar_upload = True
if 'arquivo_carregado' in st.session_state:
    mostrar_upload = False

# Widget para upload de arquivo (só mostra se não houver arquivo carregado)
uploaded_file = None
if mostrar_upload:
    uploaded_file = st.file_uploader(
        "📁 Carregue o arquivo Excel com os dados de produtividade",
        type=['xlsx', 'xls'],
        key="file_uploader"
    )
    
    if uploaded_file is not None:
        # Marcar que arquivo foi carregado (armazenar em bytes para persistir)
        st.session_state.arquivo_carregado = uploaded_file.read()
        st.session_state.arquivo_nome = uploaded_file.name
        # Resetar posição do arquivo
        uploaded_file.seek(0)
        st.rerun()
elif 'arquivo_carregado' in st.session_state:
    # Recriar objeto UploadedFile a partir dos bytes armazenados
    from io import BytesIO
    uploaded_file = BytesIO(st.session_state.arquivo_carregado)
    uploaded_file.name = st.session_state.get('arquivo_nome', 'arquivo.xlsx')

if uploaded_file is not None:
    # Carregar dados
    df, todas_abas = carregar_dados(uploaded_file)
    
    if df is not None:
        # ========== EXIBIR MÊS DE REFERÊNCIA ==========
        # Extrair o mês do DataFrame (pegar o primeiro mês único que não seja "Não informado")
        meses_no_df = [m for m in df['Mês'].unique() if pd.notna(m) and m != 'Não informado']
        if len(meses_no_df) > 0:
            mes_referencia = sorted(meses_no_df)[0]  # Pega o primeiro mês alfabeticamente
            st.subheader(f"📅 Mês de Referência: {mes_referencia}")
        else:
            st.subheader("📅 Mês de Referência: Não informado")
        
        # ========== FILTROS NA SIDEBAR ==========
        st.sidebar.header("🔍 Filtros")
        
        # Seção de Filtros Temporais
        st.sidebar.subheader("📅 Período")
        dias_disponiveis = sorted([aba for aba in todas_abas if aba.startswith("Dia")])
        
        # Inicializar session_state para dias se não existir
        if 'dias_selecionados' not in st.session_state:
            st.session_state.dias_selecionados = dias_disponiveis.copy()
        
        # Botões de seleção rápida para dias
        col_btn1, col_btn2 = st.sidebar.columns(2)
        with col_btn1:
            if st.button("✅ Todos", key="todos_dias", use_container_width=True):
                st.session_state.dias_selecionados = dias_disponiveis.copy()
                st.session_state.multiselect_dias = dias_disponiveis.copy()
                st.rerun()
        with col_btn2:
            if st.button("❌ Limpar", key="limpar_dias", use_container_width=True):
                st.session_state.dias_selecionados = []
                st.session_state.multiselect_dias = []
                st.rerun()
        
        # Sincronizar session_state do multiselect se não existir
        if 'multiselect_dias' not in st.session_state:
            st.session_state.multiselect_dias = st.session_state.dias_selecionados.copy()
        
        # Filtrar valores padrão para garantir que estejam nas opções disponíveis
        multiselect_dias_default = [
            dia for dia in st.session_state.multiselect_dias 
            if dia in dias_disponiveis
        ]
        
        dias_selecionados = st.sidebar.multiselect(
            "Escolha os dias:",
            options=dias_disponiveis,
            default=multiselect_dias_default,
            key="multiselect_dias",
            placeholder="Selecione os dias..."
        )
        
        # Atualizar session_state quando o multiselect mudar
        st.session_state.dias_selecionados = dias_selecionados
        
        # Informação sobre dias selecionados
        st.sidebar.caption(f"📅 {len(dias_selecionados)} de {len(dias_disponiveis)} dias selecionados")
        
        st.sidebar.markdown("---")
        
        # ========== FILTRO POR MÊS ==========
        st.sidebar.subheader("📆 Mês")
        meses_disponiveis = sorted([m for m in df['Mês'].unique() if pd.notna(m) and m != 'Não informado'])
        
        if len(meses_disponiveis) > 0:
            # Inicializar session_state para meses se não existir
            if 'meses_selecionados' not in st.session_state:
                st.session_state.meses_selecionados = meses_disponiveis.copy()
            
            # Sincronizar session_state do multiselect se não existir
            if 'multiselect_meses' not in st.session_state:
                st.session_state.multiselect_meses = st.session_state.meses_selecionados.copy()
            
            # Filtrar valores padrão para garantir que estejam nas opções disponíveis
            multiselect_meses_default = [
                mes for mes in st.session_state.multiselect_meses 
                if mes in meses_disponiveis
            ]
            
            meses_selecionados = st.sidebar.multiselect(
                "Escolha os meses:",
                options=meses_disponiveis,
                default=multiselect_meses_default,
                key="multiselect_meses",
                placeholder="Selecione os meses..."
            )
            
            # Atualizar session_state quando o multiselect mudar
            st.session_state.meses_selecionados = meses_selecionados
            
            # Informação sobre meses selecionados
            st.sidebar.caption(f"📆 {len(meses_selecionados)} de {len(meses_disponiveis)} mês(es) selecionado(s)")
        else:
            meses_selecionados = []
            st.sidebar.info("📆 Nenhum mês identificado no arquivo")
        
        st.sidebar.markdown("---")
        
        # Lista de profissionais e equipes disponíveis
        profissionais_disponiveis = sorted([p for p in df['Profissional'].unique() if pd.notna(p)])
        equipes_disponiveis = sorted([e for e in df['Especialidade'].unique() if pd.notna(e)])
        
        # Criar mapeamento de Profissional -> Especialidade
        mapeamento_prof_equipe = df[['Profissional', 'Especialidade']].drop_duplicates(subset='Profissional')
        mapeamento_prof_equipe = dict(zip(mapeamento_prof_equipe['Profissional'], mapeamento_prof_equipe['Especialidade']))
        
        # Inicializar session_state para profissionais se não existir
        if 'profissionais_selecionados' not in st.session_state:
            st.session_state.profissionais_selecionados = profissionais_disponiveis.copy()
        
        # Inicializar session_state para equipes se não existir
        if 'equipes_selecionadas' not in st.session_state:
            st.session_state.equipes_selecionadas = equipes_disponiveis.copy()
        
        # ========== SEÇÃO DE EQUIPES (ANTES DOS PROFISSIONAIS) ==========
        st.sidebar.subheader("🏥 Equipes")
        
        # Botões de seleção rápida para equipes
        col_btn_eq1, col_btn_eq2 = st.sidebar.columns(2)
        with col_btn_eq1:
            if st.button("✅ Todas", key="todas_equipes", use_container_width=True):
                st.session_state.equipes_selecionadas = equipes_disponiveis.copy()
                st.session_state.multiselect_equipes = equipes_disponiveis.copy()
                st.session_state.equipes_selecionadas_anteriores = equipes_disponiveis.copy()
                st.rerun()
        with col_btn_eq2:
            if st.button("❌ Limpar", key="limpar_equipes", use_container_width=True):
                st.session_state.equipes_selecionadas = []
                st.session_state.multiselect_equipes = []
                st.session_state.profissionais_selecionados = []
                st.session_state.equipes_selecionadas_anteriores = []
                st.rerun()
        
        # Sincronizar session_state do multiselect se não existir
        if 'multiselect_equipes' not in st.session_state:
            st.session_state.multiselect_equipes = st.session_state.equipes_selecionadas.copy()
        
        # Filtrar valores padrão para garantir que estejam nas opções disponíveis
        multiselect_equipes_default = [
            eq for eq in st.session_state.multiselect_equipes 
            if eq in equipes_disponiveis
        ]
        
        equipes_selecionadas = st.sidebar.multiselect(
            "Escolha as equipes:",
            options=equipes_disponiveis,
            default=multiselect_equipes_default,
            key="multiselect_equipes",
            placeholder="Selecione as equipes..."
        )
        
        # Atualizar session_state quando o multiselect mudar
        st.session_state.equipes_selecionadas = equipes_selecionadas
        
        # Informação sobre equipes selecionadas
        st.sidebar.caption(f"🏥 {len(equipes_selecionadas)} de {len(equipes_disponiveis)} equipes selecionadas")
        
        st.sidebar.markdown("---")
        
        # ========== SINCRONIZAÇÃO ENTRE EQUIPES E PROFISSIONAIS ==========
        # Verificar se a seleção de equipes mudou
        equipes_anteriores = st.session_state.get('equipes_selecionadas_anteriores', [])
        equipes_mudaram = set(equipes_selecionadas) != set(equipes_anteriores)
        
        # Se equipes foram selecionadas e mudaram, marcar automaticamente todos os profissionais dessas equipes
        if len(equipes_selecionadas) > 0 and equipes_mudaram:
            # Identificar quais profissionais pertencem às equipes selecionadas
            profissionais_das_equipes = [
                prof for prof, equipe in mapeamento_prof_equipe.items()
                if equipe in equipes_selecionadas
            ]
            
            # Verificar quais equipes foram adicionadas e quais foram removidas
            equipes_adicionadas = set(equipes_selecionadas) - set(equipes_anteriores)
            equipes_removidas = set(equipes_anteriores) - set(equipes_selecionadas)
            
            # Identificar profissionais das equipes removidas (para desmarcar)
            profissionais_das_equipes_removidas = [
                prof for prof, equipe in mapeamento_prof_equipe.items()
                if equipe in equipes_removidas
            ]
            
            # Remover profissionais das equipes removidas
            profissionais_atualizados = [
                p for p in st.session_state.profissionais_selecionados
                if p not in profissionais_das_equipes_removidas
            ]
            
            # Adicionar profissionais das equipes adicionadas (se ainda não estiverem marcados)
            profissionais_das_equipes_adicionadas = [
                prof for prof, equipe in mapeamento_prof_equipe.items()
                if equipe in equipes_adicionadas
            ]
            
            for prof in profissionais_das_equipes_adicionadas:
                if prof not in profissionais_atualizados:
                    profissionais_atualizados.append(prof)
            
            # Se não havia equipes anteriores, marcar todos os profissionais das equipes selecionadas
            if len(equipes_anteriores) == 0:
                profissionais_atualizados = profissionais_das_equipes.copy()
            
            # Atualizar session_state
            st.session_state.profissionais_selecionados = profissionais_atualizados
            
            # Atualizar checkboxes
            for prof in profissionais_disponiveis:
                checkbox_key = f"checkbox_{prof}"
                if prof in st.session_state.profissionais_selecionados:
                    st.session_state[checkbox_key] = True
                else:
                    st.session_state[checkbox_key] = False
            
            # Guardar o estado atual das equipes para comparar na próxima vez
            st.session_state.equipes_selecionadas_anteriores = equipes_selecionadas.copy()
            
            # Forçar rerun para atualizar a interface
            st.rerun()
        elif len(equipes_selecionadas) == 0 and equipes_mudaram:
            # Se nenhuma equipe foi selecionada, limpar todos os profissionais
            st.session_state.profissionais_selecionados = []
            for prof in profissionais_disponiveis:
                checkbox_key = f"checkbox_{prof}"
                st.session_state[checkbox_key] = False
            st.session_state.equipes_selecionadas_anteriores = []
            st.rerun()
        else:
            # Guardar o estado atual das equipes para comparar na próxima vez
            if 'equipes_selecionadas_anteriores' not in st.session_state:
                st.session_state.equipes_selecionadas_anteriores = equipes_selecionadas.copy()
        
        # ========== SEÇÃO DE PROFISSIONAIS (DEPOIS DA SINCRONIZAÇÃO) ==========
        st.sidebar.subheader("👥 Profissionais")
        
        # Aplicar CSS para diminuir fonte dos checkboxes de profissionais
        st.sidebar.markdown("""
        <style>
        .stCheckbox label {
            font-size: 0.85em !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Botões de seleção rápida
        col_btn_prof1, col_btn_prof2 = st.sidebar.columns(2)
        
        with col_btn_prof1:
            if st.button("✅ Todos", key="todos_profissionais", use_container_width=True):
                # Adicionar todos os profissionais das equipes selecionadas
                if len(equipes_selecionadas) > 0:
                    profissionais_das_equipes = [
                        prof for prof, equipe in mapeamento_prof_equipe.items()
                        if equipe in equipes_selecionadas
                    ]
                    st.session_state.profissionais_selecionados = profissionais_das_equipes.copy()
                else:
                    st.session_state.profissionais_selecionados = profissionais_disponiveis.copy()
                st.rerun()
        
        with col_btn_prof2:
            if st.button("❌ Limpar", key="limpar_profissionais", use_container_width=True):
                st.session_state.profissionais_selecionados = []
                st.rerun()
        
        # Filtrar profissionais disponíveis baseado nas equipes selecionadas
        if len(equipes_selecionadas) > 0:
            profissionais_disponiveis_filtrados = [
                prof for prof in profissionais_disponiveis
                if mapeamento_prof_equipe.get(prof) in equipes_selecionadas
            ]
        else:
            profissionais_disponiveis_filtrados = profissionais_disponiveis.copy()
        
        # Criar checkboxes de forma mais compacta
        if len(profissionais_disponiveis_filtrados) > 0:
            # Usar um container com scroll para não ocupar muito espaço
            with st.sidebar.container():
                for profissional in profissionais_disponiveis_filtrados:
                    # Verificar se está selecionado no session_state
                    is_selected = profissional in st.session_state.profissionais_selecionados
                    
                    # Criar checkbox com key única
                    checkbox_key = f"checkbox_{profissional}"
                    
                    # Inicializar o checkbox no session_state se não existir
                    if checkbox_key not in st.session_state:
                        st.session_state[checkbox_key] = is_selected
                    
                    # Criar checkbox - o Streamlit mantém o estado automaticamente através da key
                    checkbox_value = st.sidebar.checkbox(
                        profissional,
                        value=st.session_state[checkbox_key],
                        key=checkbox_key
                    )
                    
                    # Atualizar lista de profissionais selecionados baseado no checkbox
                    if checkbox_value:
                        if profissional not in st.session_state.profissionais_selecionados:
                            st.session_state.profissionais_selecionados.append(profissional)
                    else:
                        if profissional in st.session_state.profissionais_selecionados:
                            st.session_state.profissionais_selecionados.remove(profissional)
        
        # Usar session_state para profissionais selecionados
        profissionais_selecionados = [
            p for p in st.session_state.profissionais_selecionados 
            if p in profissionais_disponiveis_filtrados
        ]
        
        # Informação sobre seleção
        st.sidebar.caption(f"👥 {len(profissionais_selecionados)} de {len(profissionais_disponiveis_filtrados)} profissionais selecionados")
        
        st.sidebar.markdown("---")
        
        # Seção de Filtros de Status de Atendimento
        st.sidebar.subheader("📊 Status de Atendimento")
        status_disponiveis = sorted([s for s in df['Status_Consolidado'].unique() if pd.notna(s)])
        
        # Garantir que os valores padrão estejam nas opções disponíveis
        status_default = status_disponiveis.copy()
        if 'multiselect_status' in st.session_state:
            status_default = [
                s for s in st.session_state.multiselect_status 
                if s in status_disponiveis
            ]
            if len(status_default) == 0:
                status_default = status_disponiveis.copy()
        
        status_selecionados = st.sidebar.multiselect(
            "Selecione os Status:",
            options=status_disponiveis,
            default=status_default,
            key="multiselect_status"
        )
        
        # Aplicar filtros
        condicao = (
            (df['Dia'].isin(dias_selecionados)) &
            (df['Profissional'].isin(profissionais_selecionados)) &
            (df['Especialidade'].isin(equipes_selecionadas)) &
            (df['Status_Consolidado'].isin(status_selecionados))
        )
        
        # Adicionar filtro de mês se houver meses selecionados
        if len(meses_selecionados) > 0:
            condicao = condicao & (df['Mês'].isin(meses_selecionados))
        
        df_filtrado = df[condicao].copy()
        
        st.markdown("---")
        
        # ========== MÉTRICAS KPIs ==========
        st.header("📈 Métricas Principais")
        
        col1, col2, col3, col4 = st.columns(4)
        
        # Calcular métricas usando Status_Consolidado
        total_atendimentos_realizados = len(df_filtrado[df_filtrado['Status_Consolidado'] == 'Atendimento realizado'])
        total_faltosos = len(df_filtrado[df_filtrado['Status_Consolidado'] == 'Faltoso'])
        total_evadidos = len(df_filtrado[df_filtrado['Status_Consolidado'] == 'Evadido'])
        total_registros = len(df_filtrado)
        
        # Percentuais
        percentual_faltosos = (total_faltosos / total_registros * 100) if total_registros > 0 else 0
        percentual_evadidos = (total_evadidos / total_registros * 100) if total_registros > 0 else 0
        
        # Média de atendimentos por dia
        dias_unicos = df_filtrado['Dia'].nunique()
        media_por_dia = (total_atendimentos_realizados / dias_unicos) if dias_unicos > 0 else 0
        
        with col1:
            st.metric(
                "Total de Atendimentos Realizados",
                total_atendimentos_realizados
            )
        
        with col2:
            st.metric(
                "Média de Atendimentos/Dia",
                f"{media_por_dia:.1f}"
            )
        
        with col3:
            st.metric(
                "Percentual de Faltosos",
                f"{percentual_faltosos:.2f}%",
                delta=f"{total_faltosos} faltosos"
            )
        
        with col4:
            st.metric(
                "Percentual de Evadidos",
                f"{percentual_evadidos:.2f}%",
                delta=f"{total_evadidos} evadidos"
            )
        
        # Informações de Faltosos e Evadidos
        st.markdown("---")
        col_info1, col_info2 = st.columns(2)
        
        with col_info1:
            st.info(f"📋 **Quantidade de Faltosos:** {total_faltosos} | **Percentual:** {percentual_faltosos:.2f}%")
            
            # Expander com detalhes por profissional
            with st.expander("📊 Ver percentual de faltosos por profissional"):
                # Calcular faltosos por profissional
                df_faltosos = df_filtrado[df_filtrado['Status_Consolidado'] == 'Faltoso']
                
                if len(df_faltosos) > 0:
                    # Contagem de faltosos por profissional
                    faltosos_por_prof = df_faltosos.groupby('Profissional').size().reset_index(name='Qtd Faltosos')
                    
                    # Contagem total de registros por profissional (todos os status) - apenas para calcular percentual
                    total_por_prof = df_filtrado.groupby('Profissional').size().reset_index(name='Total Registros')
                    
                    # Calcular percentual
                    percentual_por_prof = faltosos_por_prof.merge(total_por_prof, on='Profissional', how='left')
                    percentual_por_prof['Percentual'] = (percentual_por_prof['Qtd Faltosos'] / percentual_por_prof['Total Registros'] * 100).round(2)
                    percentual_por_prof = percentual_por_prof.sort_values('Percentual', ascending=False)
                    
                    # Exibir tabela - apenas Profissional, Quantidade de Faltosos e Percentual
                    st.dataframe(
                        percentual_por_prof[['Profissional', 'Qtd Faltosos', 'Percentual']].rename(
                            columns={
                                'Profissional': 'Profissional',
                                'Qtd Faltosos': 'Quantidade de Faltosos',
                                'Percentual': 'Percentual (%)'
                            }
                        ),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("Nenhum registro de faltosos encontrado para os filtros selecionados.")
        
        with col_info2:
            st.info(f"📋 **Quantidade de Evadidos:** {total_evadidos} | **Percentual:** {percentual_evadidos:.2f}%")
            
            # Expander com detalhes por profissional
            with st.expander("📊 Ver percentual de evadidos por profissional"):
                # Calcular evadidos por profissional
                df_evadidos = df_filtrado[df_filtrado['Status_Consolidado'] == 'Evadido']
                
                if len(df_evadidos) > 0:
                    # Contagem de evadidos por profissional
                    evadidos_por_prof = df_evadidos.groupby('Profissional').size().reset_index(name='Qtd Evadidos')
                    
                    # Contagem total de registros por profissional (todos os status) - apenas para calcular percentual
                    total_por_prof = df_filtrado.groupby('Profissional').size().reset_index(name='Total Registros')
                    
                    # Calcular percentual
                    percentual_por_prof = evadidos_por_prof.merge(total_por_prof, on='Profissional', how='left')
                    percentual_por_prof['Percentual'] = (percentual_por_prof['Qtd Evadidos'] / percentual_por_prof['Total Registros'] * 100).round(2)
                    percentual_por_prof = percentual_por_prof.sort_values('Percentual', ascending=False)
                    
                    # Exibir tabela - apenas Profissional, Quantidade de Evadidos e Percentual
                    st.dataframe(
                        percentual_por_prof[['Profissional', 'Qtd Evadidos', 'Percentual']].rename(
                            columns={
                                'Profissional': 'Profissional',
                                'Qtd Evadidos': 'Quantidade de Evadidos',
                                'Percentual': 'Percentual (%)'
                            }
                        ),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("Nenhum registro de evadidos encontrado para os filtros selecionados.")
        
        st.markdown("---")
        
        # ========== GRÁFICOS ==========
        st.header("📊 Visualizações")
        
        # Preparar dados para gráficos (usando Status_Consolidado)
        df_finalizados = df_filtrado[df_filtrado['Status_Consolidado'] == 'Atendimento realizado']
        
        # Gráfico 1: Atendimentos por Profissional (ocupando toda a largura)
        # Opção de tipo de gráfico acima do gráfico
        tipo_grafico_profissional = st.selectbox(
            "Tipo de gráfico:",
            ["Barras", "Pizza", "Linhas"],
            key="tipo_graf_prof",
            index=0
        )
        
        st.subheader("Atendimentos por Profissional")
        
        # Contagem por profissional
        atendimentos_profissional = df_finalizados.groupby('Profissional').size().reset_index(name='Qtd Atendimentos')
        atendimentos_profissional = atendimentos_profissional.sort_values('Qtd Atendimentos', ascending=False)
        
        # Criar campo combinado com profissional e quantidade para a legenda (todos os tipos de gráfico)
        atendimentos_profissional_com_legenda = atendimentos_profissional.copy()
        atendimentos_profissional_com_legenda['Profissional_Completo'] = atendimentos_profissional_com_legenda.apply(
            lambda row: f"{row['Profissional']} ({row['Qtd Atendimentos']} atendimentos)", 
            axis=1
        )
        
        # Criar gráfico baseado na seleção
        if tipo_grafico_profissional == "Barras":
            chart_profissional = alt.Chart(atendimentos_profissional_com_legenda).mark_bar().encode(
                x=alt.X('Qtd Atendimentos:Q', title='Quantidade de Atendimentos'),
                y=alt.Y('Profissional:N', sort='-x', title='Profissional'),
                color=alt.Color('Qtd Atendimentos:Q', scale=alt.Scale(scheme='blues')),
                tooltip=[alt.Tooltip('Profissional:N', title='Profissional'), 
                        alt.Tooltip('Qtd Atendimentos:Q', title='Atendimentos')]
            ).properties(height=400)
        elif tipo_grafico_profissional == "Pizza":
            chart_profissional = alt.Chart(atendimentos_profissional_com_legenda).mark_arc(innerRadius=0).encode(
                theta=alt.Theta('Qtd Atendimentos:Q', stack=True),
                color=alt.Color('Profissional_Completo:N', 
                              scale=alt.Scale(scheme='category20'),
                              legend=alt.Legend(title='Profissional', 
                                              orient='right',
                                              labelLimit=500,  # Valor alto para evitar truncamento
                                              labelFontSize=14,  # Fonte maior
                                              titleFontSize=16,  # Título da legenda maior
                                              offset=10,  # Espaçamento próximo ao gráfico
                                              padding=10,  # Espaçamento interno
                                              columnPadding=5)),  # Espaçamento entre itens
                tooltip=[alt.Tooltip('Profissional:N', title='Profissional'), 
                        alt.Tooltip('Qtd Atendimentos:Q', title='Atendimentos')]
            ).properties(height=400, width=500).configure_view(strokeWidth=0)
        else:  # Linhas
            chart_profissional = alt.Chart(atendimentos_profissional_com_legenda).mark_line(point=True).encode(
                x=alt.X('Profissional:N', sort='-y', title='Profissional'),
                y=alt.Y('Qtd Atendimentos:Q', title='Quantidade de Atendimentos'),
                tooltip=[alt.Tooltip('Profissional:N', title='Profissional'), 
                        alt.Tooltip('Qtd Atendimentos:Q', title='Atendimentos')]
            ).properties(height=400)
        
        st.altair_chart(chart_profissional, use_container_width=True)
        
        # Estatísticas abaixo do gráfico
        if len(atendimentos_profissional) > 0:
            st.caption(f"📊 **Total:** {atendimentos_profissional['Qtd Atendimentos'].sum()} atendimentos | "
                      f"**Média:** {atendimentos_profissional['Qtd Atendimentos'].mean():.1f} | "
                      f"**Máximo:** {atendimentos_profissional['Qtd Atendimentos'].max()}")
        
        st.markdown("---")
        
        # Gráfico 2: Atendimentos por Especialidades (ocupando toda a largura, abaixo do anterior)
        # Opção de tipo de gráfico acima do gráfico
        tipo_grafico_equipe = st.selectbox(
            "Tipo de gráfico:",
            ["Barras", "Pizza", "Linhas"],
            key="tipo_graf_equipe",
            index=0
        )
        
        st.subheader("Atendimentos por Especialidades")
        
        # Contagem por equipe (Especialidade)
        atendimentos_equipe = df_finalizados.groupby('Especialidade').size().reset_index(name='Qtd Atendimentos')
        atendimentos_equipe = atendimentos_equipe.sort_values('Qtd Atendimentos', ascending=False)
        
        # Criar campo combinado com especialidade e quantidade para a legenda (todos os tipos de gráfico)
        atendimentos_equipe_com_legenda = atendimentos_equipe.copy()
        atendimentos_equipe_com_legenda['Especialidade_Completa'] = atendimentos_equipe_com_legenda.apply(
            lambda row: f"{row['Especialidade']} ({row['Qtd Atendimentos']} atendimentos)", 
            axis=1
        )
        
        # Criar gráfico baseado na seleção
        if tipo_grafico_equipe == "Barras":
            chart_equipe = alt.Chart(atendimentos_equipe_com_legenda).mark_bar().encode(
                x=alt.X('Qtd Atendimentos:Q', title='Quantidade de Atendimentos'),
                y=alt.Y('Especialidade:N', sort='-x', title='Especialidade'),
                color=alt.Color('Qtd Atendimentos:Q', scale=alt.Scale(scheme='greens')),
                tooltip=[alt.Tooltip('Especialidade:N', title='Especialidade'), 
                        alt.Tooltip('Qtd Atendimentos:Q', title='Atendimentos')]
            ).properties(height=400)
        elif tipo_grafico_equipe == "Pizza":
            chart_equipe = alt.Chart(atendimentos_equipe_com_legenda).mark_arc(innerRadius=0).encode(
                theta=alt.Theta('Qtd Atendimentos:Q', stack=True),
                color=alt.Color('Especialidade_Completa:N', 
                              scale=alt.Scale(scheme='category10'),
                              legend=alt.Legend(title='Especialidade', 
                                              orient='right',
                                              labelLimit=500,  # Valor alto para evitar truncamento
                                              labelFontSize=14,  # Fonte maior
                                              titleFontSize=16,  # Título da legenda maior
                                              offset=10,  # Espaçamento próximo ao gráfico
                                              padding=10,  # Espaçamento interno
                                              columnPadding=5)),  # Espaçamento entre itens
                tooltip=[alt.Tooltip('Especialidade:N', title='Especialidade'), 
                        alt.Tooltip('Qtd Atendimentos:Q', title='Atendimentos')]
            ).properties(height=400, width=500).configure_view(strokeWidth=0)
        else:  # Linhas
            chart_equipe = alt.Chart(atendimentos_equipe_com_legenda).mark_line(point=True).encode(
                x=alt.X('Especialidade:N', sort='-y', title='Especialidade'),
                y=alt.Y('Qtd Atendimentos:Q', title='Quantidade de Atendimentos'),
                tooltip=[alt.Tooltip('Especialidade:N', title='Especialidade'), 
                        alt.Tooltip('Qtd Atendimentos:Q', title='Atendimentos')]
            ).properties(height=400)
        
        st.altair_chart(chart_equipe, use_container_width=True)
        
        # Estatísticas abaixo do gráfico
        if len(atendimentos_equipe) > 0:
            st.caption(f"📊 **Total:** {atendimentos_equipe['Qtd Atendimentos'].sum()} atendimentos")
        
        # Gráfico de Evolução dos Atendimentos por Dia e Profissional
        st.markdown("---")
        st.subheader("📈 Evolução dos Atendimentos por Dia")
        
        # Seletor de tipo de gráfico
        tipo_grafico_temporal = st.selectbox(
            "Tipo de gráfico:",
            ["Linhas", "Barras"],
            key="tipo_graf_temporal",
            index=0
        )
        
        # Obter todos os dias disponíveis nos dados filtrados (não apenas finalizados)
        todos_dias_disponiveis = sorted(df_filtrado['Dia'].unique())
        
        # Contagem por dia e profissional (apenas finalizados)
        atendimentos_por_dia_prof = df_finalizados.groupby(['Dia', 'Profissional']).size().reset_index(name='Qtd Atendimentos')
        
        # Criar estrutura completa: todos os dias x todos os profissionais
        # Isso garante que todos os dias apareçam, mesmo sem atendimentos
        from itertools import product
        
        if len(profissionais_selecionados) > 0 and len(todos_dias_disponiveis) > 0:
            # Criar todas as combinações de dia e profissional
            combinacoes = pd.DataFrame(
                list(product(todos_dias_disponiveis, profissionais_selecionados)),
                columns=['Dia', 'Profissional']
            )
            
            # Fazer merge com os dados reais, preenchendo com 0 onde não houver dados
            atendimentos_completo = combinacoes.merge(
                atendimentos_por_dia_prof,
                on=['Dia', 'Profissional'],
                how='left'
            ).fillna(0)
            
            # Garantir que Qtd Atendimentos seja inteiro
            atendimentos_completo['Qtd Atendimentos'] = atendimentos_completo['Qtd Atendimentos'].astype(int)
            
            # Ordenar por dia
            atendimentos_completo = atendimentos_completo.sort_values('Dia')
            
            # Calcular média para exibição
            media_atendimentos = atendimentos_completo['Qtd Atendimentos'].mean() if len(atendimentos_completo) > 0 else 0
            
            # Criar gráfico baseado na seleção
            if tipo_grafico_temporal == "Linhas":
                # Criar gráfico base
                chart_base = alt.Chart(atendimentos_completo).encode(
                    x=alt.X('Dia:N', sort='x', title='Dia'),
                    y=alt.Y('Qtd Atendimentos:Q', title='Quantidade de Atendimentos'),
                    tooltip=['Dia', 'Profissional', 'Qtd Atendimentos']
                )
                
                # Linha para cada profissional (com cores diferentes)
                chart_temporal = chart_base.mark_line(
                    point=True,
                    strokeWidth=3
                ).encode(
                    color=alt.Color(
                        'Profissional:N',
                        scale=alt.Scale(scheme='category20'),
                        legend=alt.Legend(title='Profissional', orient='right')
                    )
                ).properties(
                    height=400,
                    width=800
                )
            else:  # Barras
                # Gráfico de barras agrupadas por dia (lado a lado)
                # No Altair, barras agrupadas são criadas usando x para categoria principal
                # e color para subcategoria, o que automaticamente cria barras lado a lado
                chart_temporal = alt.Chart(atendimentos_completo).mark_bar(
                    cornerRadiusTopLeft=3,
                    cornerRadiusTopRight=3
                ).encode(
                    x=alt.X('Dia:N', sort='x', title='Dia', axis=alt.Axis(labelAngle=-45)),
                    y=alt.Y('Qtd Atendimentos:Q', title='Quantidade de Atendimentos', scale=alt.Scale(domain=[0, None])),
                    color=alt.Color(
                        'Profissional:N',
                        scale=alt.Scale(scheme='category20'),
                        legend=alt.Legend(title='Profissional', orient='right')
                    ),
                    tooltip=['Dia', 'Profissional', 'Qtd Atendimentos']
                ).properties(
                    height=400,
                    width=800
                )
            
            st.altair_chart(chart_temporal, use_container_width=True)
            
            # Informações sobre os profissionais
            num_profissionais = len(profissionais_selecionados)
            if num_profissionais > 0:
                st.caption(f"📊 **{num_profissionais} profissional(is) selecionado(s)** | "
                          f"**Média de atendimentos por dia:** {media_atendimentos:.1f} | "
                          f"**Total de dias:** {len(todos_dias_disponiveis)}")
        else:
            st.warning("Nenhum dado disponível para exibir o gráfico temporal.")
        
        # Gráfico de Status
        st.markdown("---")
        st.subheader("Distribuição de Status")
        
        status_counts = df_filtrado.groupby('Status_Consolidado').size().reset_index(name='Quantidade')
        status_counts = status_counts.sort_values('Quantidade', ascending=False)
        
        # Criar campo combinado com status e quantidade para a legenda
        status_counts_com_legenda = status_counts.copy()
        status_counts_com_legenda['Status_Completo'] = status_counts_com_legenda.apply(
            lambda row: f"{row['Status_Consolidado']} ({row['Quantidade']} atendimentos)", 
            axis=1
        )
        
        # Gráfico de Pizza para Distribuição de Status
        chart_status_pizza = alt.Chart(status_counts_com_legenda).mark_arc(innerRadius=0).encode(
            theta=alt.Theta('Quantidade:Q', stack=True),
            color=alt.Color('Status_Completo:N', 
                          scale=alt.Scale(scheme='set2'),
                          legend=alt.Legend(title='Status de Atendimento', 
                                          orient='right',
                                          labelLimit=500,
                                          labelFontSize=14,
                                          titleFontSize=16,
                                          offset=10,
                                          padding=10,
                                          columnPadding=5)),
            tooltip=[alt.Tooltip('Status_Consolidado:N', title='Status'), 
                    alt.Tooltip('Quantidade:Q', title='Quantidade')]
        ).properties(height=400, width=500).configure_view(strokeWidth=0)
        
        st.altair_chart(chart_status_pizza, use_container_width=True)
        
        # Estatísticas abaixo do gráfico de pizza
        if len(status_counts) > 0:
            st.caption(f"📊 **Total:** {status_counts['Quantidade'].sum()} atendimentos")
        
        st.markdown("---")
        
        # Gráfico de Status por Profissional (em linha completa)
        st.subheader("Status por Profissional (Top 10)")
        status_prof = df_filtrado.groupby(['Profissional', 'Status_Consolidado']).size().reset_index(name='Quantidade')
        
        # Pegar os top 10 profissionais por quantidade total
        total_por_prof = status_prof.groupby('Profissional')['Quantidade'].sum().reset_index(name='Total')
        top_10_profissionais = total_por_prof.nlargest(10, 'Total')['Profissional'].tolist()
        
        # Filtrar apenas os top 10 profissionais
        status_prof_top10 = status_prof[status_prof['Profissional'].isin(top_10_profissionais)]
        
        # Ordenar alfabeticamente pelos nomes dos profissionais
        status_prof_top10 = status_prof_top10.sort_values('Profissional', ascending=True)
        
        chart_status_prof = alt.Chart(status_prof_top10).mark_bar().encode(
            x=alt.X('Quantidade:Q', title='Quantidade'),
            y=alt.Y('Profissional:N', sort='y', title='Profissional'),  # Ordenar alfabeticamente
            color=alt.Color('Status_Consolidado:N', 
                          scale=alt.Scale(scheme='set2'),
                          legend=alt.Legend(title='Status de Atendimento',
                                          orient='right',
                                          labelFontSize=12,
                                          titleFontSize=14)),
            tooltip=['Profissional', 'Status_Consolidado', 'Quantidade']
        ).properties(height=400)
        
        st.altair_chart(chart_status_prof, use_container_width=True)
        
        # ========== TABELA DE DADOS ==========
        st.markdown("---")
        with st.expander("📋 Visualizar Dados Filtrados"):
            st.dataframe(
                df_filtrado,
                use_container_width=True,
                height=400
            )
            
            # Botão para download
            csv = df_filtrado.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Baixar dados filtrados (CSV)",
                data=csv,
                file_name="dados_filtrados.csv",
                mime="text/csv"
            )
        
        # Informações sobre o dataset
        st.sidebar.markdown("---")
        st.sidebar.info(f"📊 **Total de registros:** {len(df)}")
        st.sidebar.info(f"📅 **Dias disponíveis:** {len(dias_disponiveis)}")
        st.sidebar.info(f"👥 **Profissionais:** {len(profissionais_disponiveis)}")
        st.sidebar.info(f"🏥 **Equipes:** {len(equipes_disponiveis)}")

else:
    st.info("👆 Por favor, carregue o arquivo Excel para começar a análise.")

