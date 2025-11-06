"""
Script para explorar a estrutura das planilhas e identificar colunas
necessárias para o cruzamento de atendimentos (Técnico vs Médico)
"""
import pandas as pd
import sys

# Verificar se o nome do arquivo foi passado como argumento
if len(sys.argv) > 1:
    arquivo = sys.argv[1]
else:
    # Tentar usar um arquivo padrão
    arquivo = input("Digite o nome do arquivo Excel (ou pressione Enter para usar 'Outubro.xlsx'): ").strip()
    if not arquivo:
        arquivo = "Outubro.xlsx"

print(f"\n{'='*70}")
print(f"EXPLORANDO ARQUIVO: {arquivo}")
print(f"{'='*70}\n")

# Função para extrair o número do dia do nome da aba
def extrair_dia_aba(nome_aba):
    """Extrai o número do dia do nome da aba (ex: 'Dia 01' -> 1, 'Dia 24' -> 24)"""
    import re
    # Procura por padrões como "Dia 01", "Dia 1", "Dia01", etc.
    match = re.search(r'[Dd]ia\s*(\d+)', nome_aba)
    if match:
        return int(match.group(1))
    return None

try:
    # Ler todas as abas
    xls = pd.ExcelFile(arquivo)
    
    print("=" * 70)
    print("ABAS ENCONTRADAS NO ARQUIVO:")
    print("=" * 70)
    for i, aba in enumerate(xls.sheet_names, 1):
        print(f"{i}. {aba}")
    
    # Consolidar dados de todas as abas "Dia"
    dados_consolidados = []
    
    print("\n" + "=" * 70)
    print("CONSOLIDANDO DADOS DAS ABAS 'Dia':")
    print("=" * 70)
    
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
                dia_info = f" -> Dia {dia_numero}" if dia_numero else ""
                print(f"✅ {aba}: {len(df)} registros{dia_info}")
            except Exception as e:
                print(f"❌ Erro ao ler {aba}: {e}")
    
    if not dados_consolidados:
        print("⚠️ Nenhuma aba 'Dia' encontrada!")
        sys.exit(1)
    
    # Concatenar todos
    df_consolidado = pd.concat(dados_consolidados, ignore_index=True)
    
    print(f"\n{'='*70}")
    print(f"ESTRUTURA GERAL DO DATASET:")
    print(f"{'='*70}")
    print(f"Total de registros: {len(df_consolidado)}")
    print(f"Total de colunas: {len(df_consolidado.columns)}")
    
    print(f"\n{'='*70}")
    print("TODAS AS COLUNAS ENCONTRADAS:")
    print(f"{'='*70}")
    for i, col in enumerate(df_consolidado.columns, 1):
        print(f"{i}. {col}")
        print(f"   Tipo: {df_consolidado[col].dtype}")
        valores_unicos = df_consolidado[col].nunique()
        print(f"   Valores únicos: {valores_unicos}")
        if valores_unicos <= 10:
            print(f"   Valores: {list(df_consolidado[col].unique()[:10])}")
        print()
    
    print("=" * 70)
    print("VALORES ÚNICOS DE 'Especialidade' (CRÍTICO PARA CRUZAMENTO):")
    print("=" * 70)
    if 'Especialidade' in df_consolidado.columns:
        print(df_consolidado['Especialidade'].value_counts())
        
        # Identificar especialidades de técnico e médico
        especialidades = df_consolidado['Especialidade'].unique()
        print(f"\n{'='*70}")
        print("IDENTIFICANDO ESPECIALIDADES DE TÉCNICO E MÉDICO:")
        print(f"{'='*70}")
        
        tecnicos = []
        medicos = []
        
        for esp in especialidades:
            esp_str = str(esp).upper()
            if 'TECNICO' in esp_str or 'TÉCNICO' in esp_str or 'ENFERMAGEM' in esp_str:
                tecnicos.append(esp)
            elif 'MEDICO' in esp_str or 'MÉDICO' in esp_str or 'MEDICINA' in esp_str:
                medicos.append(esp)
        
        print(f"\n🔍 Possíveis especialidades de TÉCNICO:")
        for tec in tecnicos:
            print(f"   - {tec}")
        
        print(f"\n🔍 Possíveis especialidades de MÉDICO:")
        for med in medicos:
            print(f"   - {med}")
    else:
        print("⚠️ Coluna 'Especialidade' não encontrada!")
    
    print(f"\n{'='*70}")
    print("PROFISSIONAIS ÚNICOS:")
    print(f"{'='*70}")
    if 'Profissional' in df_consolidado.columns:
        print(f"Total: {df_consolidado['Profissional'].nunique()} profissionais")
        print("\nPrimeiros 20 profissionais:")
        print(df_consolidado['Profissional'].value_counts().head(20))
    else:
        print("⚠️ Coluna 'Profissional' não encontrada!")
    
    print(f"\n{'='*70}")
    print("COLUNAS DE IDENTIFICAÇÃO DO PACIENTE:")
    print(f"{'='*70}")
    colunas_paciente = ['Paciente', 'paciente', 'Nome', 'nome', 'Prontuário', 'prontuário', 
                        'Prontuario', 'prontuario', 'CPF', 'cpf', 'ID', 'id']
    
    for col in colunas_paciente:
        if col in df_consolidado.columns:
            print(f"✅ Encontrada: {col}")
            print(f"   Valores únicos: {df_consolidado[col].nunique()}")
    
    print(f"\n{'='*70}")
    print("INFORMAÇÃO DE DIA (EXTRAÍDA DO NOME DA ABA):")
    print(f"{'='*70}")
    if 'Dia_Atendimento' in df_consolidado.columns:
        print(f"✅ Coluna 'Dia_Atendimento' criada a partir do nome da aba")
        print(f"   Valores únicos: {df_consolidado['Dia_Atendimento'].nunique()}")
        print(f"   Dias encontrados:")
        dias_unicos = sorted(df_consolidado['Dia_Atendimento'].unique())
        for dia in dias_unicos[:20]:  # Mostrar primeiros 20 dias
            count = len(df_consolidado[df_consolidado['Dia_Atendimento'] == dia])
            print(f"      - {dia}: {count} registros")
        if len(dias_unicos) > 20:
            print(f"      ... e mais {len(dias_unicos) - 20} dias")
    
    print(f"\n{'='*70}")
    print("EXEMPLOS DE NOMES DE ABAS E DIA EXTRAÍDO:")
    print(f"{'='*70}")
    abas_dia = [aba for aba in xls.sheet_names if aba.startswith("Dia")]
    for aba in abas_dia[:10]:  # Mostrar primeiras 10 abas
        dia_num = extrair_dia_aba(aba)
        print(f"   '{aba}' -> Dia {dia_num if dia_num else 'NÃO ENCONTRADO'}")
    
    print(f"\n{'='*70}")
    print("AMOSTRA DOS DADOS (Primeiras 3 linhas):")
    print(f"{'='*70}")
    print(df_consolidado.head(3).to_string())
    
    print(f"\n{'='*70}")
    print("ESTATÍSTICAS GERAIS:")
    print(f"{'='*70}")
    print(df_consolidado.info())
    
except FileNotFoundError:
    print(f"❌ Arquivo '{arquivo}' não encontrado!")
    print("\nArquivos disponíveis no diretório:")
    import os
    arquivos_xlsx = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    for f in arquivos_xlsx:
        print(f"  - {f}")
except Exception as e:
    print(f"❌ Erro ao processar arquivo: {e}")
    import traceback
    traceback.print_exc()

