"""
Módulo para cruzar atendimentos: identificar pacientes que foram ao médico
sem passar pelo técnico no mesmo dia
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Para não precisar de interface gráfica
import seaborn as sns
from datetime import datetime
import re
import os

# Configurar estilo dos gráficos
try:
    plt.style.use('seaborn-v0_8-darkgrid')
except:
    try:
        plt.style.use('seaborn-darkgrid')
    except:
        plt.style.use('default')
sns.set_palette("husl")


def extrair_dia_aba(nome_aba):
    """Extrai o número do dia do nome da aba (ex: 'Dia 01' -> 1, 'Dia 24' -> 24)"""
    match = re.search(r'[Dd]ia\s*(\d+)', nome_aba)
    if match:
        return int(match.group(1))
    return None


def carregar_dados(arquivo):
    """
    Carrega e consolida dados de todas as abas 'Dia' do arquivo Excel
    
    Args:
        arquivo: Caminho para o arquivo Excel ou nome do arquivo
        
    Returns:
        DataFrame com todos os dados consolidados
    """
    try:
        # Ler todas as abas
        xls = pd.ExcelFile(arquivo)
        
        # Consolidar dados de todas as abas "Dia"
        dados_consolidados = []
        
        for aba in xls.sheet_names:
            if aba.startswith("Dia"):
                try:
                    df = pd.read_excel(xls, sheet_name=aba)
                    # Adicionar coluna com o nome completo da aba
                    df['Aba'] = aba
                    # Extrair o número do dia
                    dia_numero = extrair_dia_aba(aba)
                    if dia_numero:
                        df['Dia_Numero'] = dia_numero
                        df['Dia_Atendimento'] = f"Dia {dia_numero:02d}"
                    else:
                        df['Dia_Numero'] = None
                        df['Dia_Atendimento'] = aba
                    dados_consolidados.append(df)
                except Exception as e:
                    print(f"⚠️ Erro ao ler {aba}: {e}")
        
        if not dados_consolidados:
            raise ValueError("Nenhuma aba 'Dia' encontrada no arquivo!")
        
        # Concatenar todos
        df_consolidado = pd.concat(dados_consolidados, ignore_index=True)
        
        # Remover coluna Unnamed: 0 se existir
        if 'Unnamed: 0' in df_consolidado.columns:
            df_consolidado = df_consolidado.drop(columns=['Unnamed: 0'])
        
        return df_consolidado
    
    except Exception as e:
        raise Exception(f"Erro ao carregar arquivo: {e}")


def cruzar_atendimentos(df):
    """
    Cruza os atendimentos para identificar quais pacientes foram ao médico
    sem passar pelo técnico no mesmo dia
    
    IMPORTANTE: Considera apenas atendimentos REALIZADOS (status: 
    'ATENDIMENTO FINALIZADO' ou 'REALIZANDO PROCEDIMENTO/EXAME')
    
    Args:
        df: DataFrame com todos os dados consolidados
        
    Returns:
        DataFrame com atendimentos médicos e informação se passou pelo técnico
    """
    # Definir especialidades de técnico e médico
    especialidade_tecnico = 'TÉCNICO DE ENFERMAGEM DA ESTRATÉGIA DE SAÚDE DA FAMÍLIA'
    especialidade_medico = 'MÉDICO DA ESTRATÉGIA DE SAÚDE DA FAMÍLIA'
    
    # Status que indicam atendimento realizado
    status_realizados = ['ATENDIMENTO FINALIZADO', 'REALIZANDO PROCEDIMENTO/EXAME']
    
    # Filtrar apenas atendimentos de médico que foram REALIZADOS
    df_medicos = df[
        (df['Especialidade'] == especialidade_medico) &
        (df['Status'].isin(status_realizados))
    ].copy()
    
    if len(df_medicos) == 0:
        raise ValueError("Nenhum atendimento médico realizado encontrado!")
    
    # Filtrar atendimentos de técnico que foram REALIZADOS
    df_tecnicos = df[
        (df['Especialidade'] == especialidade_tecnico) &
        (df['Status'].isin(status_realizados))
    ].copy()
    
    # Criar chave de identificação: Prontuário + Dia_Atendimento
    # Função para verificar se passou pelo técnico
    def verificar_passou_tecnico(row):
        prontuario = row['Número Prontuário']
        dia_atendimento = row['Dia_Atendimento']
        
        # Verificar se existe atendimento do técnico para o mesmo paciente no mesmo dia
        passou = len(df_tecnicos[
            (df_tecnicos['Número Prontuário'] == prontuario) &
            (df_tecnicos['Dia_Atendimento'] == dia_atendimento)
        ]) > 0
        
        return passou
    
    # Aplicar verificação
    df_medicos['Passou_Pelo_Tecnico'] = df_medicos.apply(verificar_passou_tecnico, axis=1)
    
    return df_medicos


def gerar_estatisticas_por_medico(df_medicos_cruzados):
    """
    Gera estatísticas por médico
    
    Args:
        df_medicos_cruzados: DataFrame com atendimentos médicos já cruzados
        
    Returns:
        DataFrame com estatísticas por médico
    """
    # Agrupar por médico
    stats = df_medicos_cruzados.groupby('Profissional').agg({
        'Número Prontuário': 'count',  # Total de atendimentos
        'Passou_Pelo_Tecnico': lambda x: (x == True).sum(),  # Quantos passaram pelo técnico
    }).rename(columns={
        'Número Prontuário': 'Total_Atendimentos',
        'Passou_Pelo_Tecnico': 'Passou_Pelo_Tecnico'
    })
    
    # Calcular quantos NÃO passaram pelo técnico
    stats['Nao_Passou_Pelo_Tecnico'] = stats['Total_Atendimentos'] - stats['Passou_Pelo_Tecnico']
    
    # Calcular percentuais
    stats['Percentual_Passou'] = (stats['Passou_Pelo_Tecnico'] / stats['Total_Atendimentos'] * 100).round(2)
    stats['Percentual_Nao_Passou'] = (stats['Nao_Passou_Pelo_Tecnico'] / stats['Total_Atendimentos'] * 100).round(2)
    
    # Ordenar por total de atendimentos
    stats = stats.sort_values('Total_Atendimentos', ascending=False)
    
    return stats


def gerar_graficos_por_medico(stats, pasta_saida='graficos'):
    """
    Gera gráficos para cada médico mostrando:
    - Total de atendimentos
    - Quantidade que passou pelo técnico
    - Quantidade que não passou pelo técnico
    
    Args:
        stats: DataFrame com estatísticas por médico
        pasta_saida: Pasta onde salvar os gráficos
    """
    # Criar pasta de saída se não existir
    os.makedirs(pasta_saida, exist_ok=True)
    
    # Gerar gráfico para cada médico
    for medico in stats.index:
        medico_stats = stats.loc[medico]
        
        # Preparar dados para o gráfico
        categorias = ['Passou pelo\nTécnico', 'Não passou pelo\nTécnico']
        valores = [
            medico_stats['Passou_Pelo_Tecnico'],
            medico_stats['Nao_Passou_Pelo_Tecnico']
        ]
        cores = ['#2ecc71', '#e74c3c']  # Verde para passou, vermelho para não passou
        
        # Criar figura
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Criar gráfico de barras
        bars = ax.bar(categorias, valores, color=cores, alpha=0.8, edgecolor='black', linewidth=1.5)
        
        # Adicionar valores nas barras
        for i, (bar, valor) in enumerate(zip(bars, valores)):
            altura = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., altura + max(valores)*0.01,
                   f'{int(valor)}\n({medico_stats["Percentual_Passou"] if i == 0 else medico_stats["Percentual_Nao_Passou"]}%)',
                   ha='center', va='bottom', fontsize=12, fontweight='bold')
        
        # Configurações do gráfico
        ax.set_ylabel('Quantidade de Atendimentos', fontsize=12, fontweight='bold')
        ax.set_title(f'Atendimentos Médicos - {medico}\n'
                    f'Total: {int(medico_stats["Total_Atendimentos"])} atendimentos',
                    fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.set_ylim(0, max(valores) * 1.2)
        
        # Adicionar linha indicando total
        ax.axhline(y=medico_stats['Total_Atendimentos'], color='blue', linestyle='--', 
                  linewidth=2, alpha=0.5, label=f'Total: {int(medico_stats["Total_Atendimentos"])}')
        ax.legend(loc='upper right')
        
        plt.tight_layout()
        
        # Salvar gráfico
        # Limpar nome do arquivo (remover caracteres inválidos)
        nome_arquivo = re.sub(r'[<>:"/\\|?*]', '_', medico)
        caminho_grafico = os.path.join(pasta_saida, f'{nome_arquivo}.png')
        plt.savefig(caminho_grafico, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"✅ Gráfico salvo: {caminho_grafico}")
    
    # Gerar gráfico consolidado com todos os médicos
    gerar_grafico_consolidado(stats, pasta_saida)


def gerar_grafico_consolidado(stats, pasta_saida='graficos'):
    """
    Gera gráfico consolidado mostrando todos os médicos
    
    Args:
        stats: DataFrame com estatísticas por médico
        pasta_saida: Pasta onde salvar o gráfico
    """
    # Preparar dados
    medicos = stats.index.tolist()
    passou = stats['Passou_Pelo_Tecnico'].tolist()
    nao_passou = stats['Nao_Passou_Pelo_Tecnico'].tolist()
    
    # Criar figura
    fig, ax = plt.subplots(figsize=(14, 8))
    
    x = range(len(medicos))
    width = 0.6
    
    # Criar barras empilhadas
    bars1 = ax.bar(x, passou, width, label='Passou pelo Técnico', color='#2ecc71', alpha=0.8, edgecolor='black')
    bars2 = ax.bar(x, nao_passou, width, bottom=passou, label='Não passou pelo Técnico', 
                   color='#e74c3c', alpha=0.8, edgecolor='black')
    
    # Adicionar valores nas barras
    for i, (p, np, total) in enumerate(zip(passou, nao_passou, stats['Total_Atendimentos'])):
        ax.text(i, total + max(stats['Total_Atendimentos']) * 0.01,
               f'Total: {int(total)}', ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    # Configurações
    ax.set_xlabel('Médico', fontsize=12, fontweight='bold')
    ax.set_ylabel('Quantidade de Atendimentos', fontsize=12, fontweight='bold')
    ax.set_title('Cruzamento de Atendimentos - Todos os Médicos\n'
                f'Total de atendimentos médicos: {int(stats["Total_Atendimentos"].sum())}',
                fontsize=14, fontweight='bold', pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(medicos, rotation=45, ha='right', fontsize=10)
    ax.legend(loc='upper right', fontsize=11)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    
    # Salvar
    caminho = os.path.join(pasta_saida, 'Todos_Medicos.png')
    plt.savefig(caminho, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"✅ Gráfico consolidado salvo: {caminho}")


def gerar_planilha_saida(df_medicos_cruzados, nome_arquivo_saida='cruzamento_atendimentos.xlsx'):
    """
    Gera planilha Excel com pacientes que não passaram pelo técnico
    
    Args:
        df_medicos_cruzados: DataFrame com atendimentos médicos já cruzados
        nome_arquivo_saida: Nome do arquivo de saída
    """
    # Filtrar apenas os que NÃO passaram pelo técnico
    df_nao_passou = df_medicos_cruzados[df_medicos_cruzados['Passou_Pelo_Tecnico'] == False].copy()
    
    # Selecionar colunas relevantes
    colunas_saida = ['Paciente', 'Número Prontuário', 'Dia_Atendimento', 'Profissional', 'Status']
    
    # Verificar quais colunas existem
    colunas_existentes = [col for col in colunas_saida if col in df_nao_passou.columns]
    df_saida = df_nao_passou[colunas_existentes].copy()
    
    # Renomear colunas para melhor apresentação
    df_saida = df_saida.rename(columns={
        'Paciente': 'Paciente',
        'Número Prontuário': 'Prontuário',
        'Dia_Atendimento': 'Dia de Atendimento',
        'Profissional': 'Médico',
        'Status': 'Status do Atendimento'
    })
    
    # Ordenar por médico e dia
    df_saida = df_saida.sort_values(['Médico', 'Dia de Atendimento', 'Paciente'])
    
    # Salvar em Excel
    with pd.ExcelWriter(nome_arquivo_saida, engine='openpyxl') as writer:
        # Aba com pacientes que não passaram pelo técnico
        df_saida.to_excel(writer, sheet_name='Pacientes para Investigação', index=False)
        
        # Aba com estatísticas gerais
        stats = gerar_estatisticas_por_medico(df_medicos_cruzados)
        stats_renomeado = stats.reset_index()
        stats_renomeado = stats_renomeado.rename(columns={
            'Profissional': 'Médico',
            'Total_Atendimentos': 'Total de Atendimentos',
            'Passou_Pelo_Tecnico': 'Passou pelo Técnico',
            'Nao_Passou_Pelo_Tecnico': 'Não Passou pelo Técnico',
            'Percentual_Passou': '% Passou pelo Técnico',
            'Percentual_Nao_Passou': '% Não Passou pelo Técnico'
        })
        stats_renomeado.to_excel(writer, sheet_name='Estatísticas por Médico', index=False)
        
        # Aba com todos os atendimentos médicos (para referência)
        df_todos = df_medicos_cruzados[['Paciente', 'Número Prontuário', 'Dia_Atendimento', 
                                       'Profissional', 'Status', 'Passou_Pelo_Tecnico']].copy()
        df_todos = df_todos.rename(columns={
            'Paciente': 'Paciente',
            'Número Prontuário': 'Prontuário',
            'Dia_Atendimento': 'Dia de Atendimento',
            'Profissional': 'Médico',
            'Status': 'Status do Atendimento',
            'Passou_Pelo_Tecnico': 'Passou pelo Técnico'
        })
        df_todos = df_todos.sort_values(['Médico', 'Dia de Atendimento', 'Paciente'])
        df_todos.to_excel(writer, sheet_name='Todos Atendimentos Médicos', index=False)
    
    print(f"✅ Planilha salva: {nome_arquivo_saida}")
    print(f"   - Total de pacientes para investigação: {len(df_nao_passou)}")
    
    return nome_arquivo_saida


def processar_arquivo(arquivo, pasta_graficos='graficos', nome_planilha_saida='cruzamento_atendimentos.xlsx'):
    """
    Função principal que processa o arquivo completo
    
    Args:
        arquivo: Caminho para o arquivo Excel
        pasta_graficos: Pasta onde salvar os gráficos
        nome_planilha_saida: Nome do arquivo Excel de saída
        
    Returns:
        Tupla com (df_medicos_cruzados, stats)
    """
    print(f"\n{'='*70}")
    print(f"PROCESSANDO ARQUIVO: {arquivo}")
    print(f"{'='*70}\n")
    
    # 1. Carregar dados
    print("📂 Carregando dados...")
    df = carregar_dados(arquivo)
    print(f"   ✅ {len(df)} registros carregados")
    
    # Informar sobre filtro de status
    status_realizados = ['ATENDIMENTO FINALIZADO', 'REALIZANDO PROCEDIMENTO/EXAME']
    print(f"\n⚠️ FILTRO APLICADO: Apenas atendimentos REALIZADOS serão considerados")
    print(f"   Status considerados: {', '.join(status_realizados)}")
    
    # Mostrar distribuição de status antes do filtro
    if 'Status' in df.columns:
        total_medicos = len(df[df['Especialidade'] == 'MÉDICO DA ESTRATÉGIA DE SAÚDE DA FAMÍLIA'])
        medicos_realizados = len(df[
            (df['Especialidade'] == 'MÉDICO DA ESTRATÉGIA DE SAÚDE DA FAMÍLIA') &
            (df['Status'].isin(status_realizados))
        ])
        print(f"   - Total de atendimentos médicos: {total_medicos}")
        print(f"   - Atendimentos médicos REALIZADOS: {medicos_realizados} ({medicos_realizados/total_medicos*100:.1f}%)")
        print(f"   - Atendimentos médicos EXCLUÍDOS (agendados/faltosos/evadidos): {total_medicos - medicos_realizados}")
    
    # 2. Cruzar atendimentos
    print("\n🔍 Cruzando atendimentos...")
    df_medicos_cruzados = cruzar_atendimentos(df)
    print(f"   ✅ {len(df_medicos_cruzados)} atendimentos médicos REALIZADOS encontrados")
    
    total_nao_passou = len(df_medicos_cruzados[df_medicos_cruzados['Passou_Pelo_Tecnico'] == False])
    total_passou = len(df_medicos_cruzados[df_medicos_cruzados['Passou_Pelo_Tecnico'] == True])
    print(f"   - Passou pelo técnico: {total_passou} ({total_passou/len(df_medicos_cruzados)*100:.1f}%)")
    print(f"   - NÃO passou pelo técnico: {total_nao_passou} ({total_nao_passou/len(df_medicos_cruzados)*100:.1f}%)")
    
    # 3. Gerar estatísticas
    print("\n📊 Gerando estatísticas por médico...")
    stats = gerar_estatisticas_por_medico(df_medicos_cruzados)
    print(f"   ✅ Estatísticas geradas para {len(stats)} médico(s)")
    
    # 4. Gerar gráficos
    print(f"\n📈 Gerando gráficos na pasta '{pasta_graficos}'...")
    gerar_graficos_por_medico(stats, pasta_graficos)
    
    # 5. Gerar planilha de saída
    print(f"\n📄 Gerando planilha de saída...")
    gerar_planilha_saida(df_medicos_cruzados, nome_planilha_saida)
    
    print(f"\n{'='*70}")
    print("✅ PROCESSAMENTO CONCLUÍDO!")
    print(f"{'='*70}\n")
    
    return df_medicos_cruzados, stats


if __name__ == "__main__":
    # Exemplo de uso
    import sys
    
    if len(sys.argv) > 1:
        arquivo = sys.argv[1]
    else:
        arquivo = input("Digite o nome do arquivo Excel: ").strip()
    
    if not arquivo:
        print("❌ Nome do arquivo não fornecido!")
        sys.exit(1)
    
    try:
        df_medicos, stats = processar_arquivo(arquivo)
        
        print("\n📋 RESUMO DAS ESTATÍSTICAS:")
        print("=" * 70)
        print(stats.to_string())
        
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        import traceback
        traceback.print_exc()

