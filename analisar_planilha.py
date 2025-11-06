"""
Script de Diagnóstico para Analisar Planilha Excel
Identifica problemas de tipos mistos e dados inconsistentes
"""
import pandas as pd
import sys
from pathlib import Path

def analisar_planilha(caminho_arquivo):
    """
    Analisa uma planilha Excel e identifica problemas de tipos mistos
    """
    print("=" * 80)
    print("🔍 ANÁLISE DE PLANILHA EXCEL")
    print("=" * 80)
    print(f"\n📁 Arquivo: {caminho_arquivo}\n")
    
    try:
        # Ler todas as abas
        xls = pd.ExcelFile(caminho_arquivo)
        print(f"📊 Total de abas: {len(xls.sheet_names)}")
        print(f"📋 Abas encontradas: {', '.join(xls.sheet_names)}\n")
        
        # Consolidar dados (mesmo processo do dashboard)
        dados_consolidados = []
        
        for aba in xls.sheet_names:
            if aba.startswith("Dia"):
                try:
                    df = pd.read_excel(xls, sheet_name=aba)
                    df['Dia'] = aba
                    dados_consolidados.append(df)
                    print(f"✅ Aba '{aba}': {len(df)} registros")
                except Exception as e:
                    print(f"❌ Erro ao ler aba '{aba}': {e}")
        
        if not dados_consolidados:
            print("\n⚠️ Nenhuma aba 'Dia' encontrada!")
            return
        
        # Concatenar todos
        df_consolidado = pd.concat(dados_consolidados, ignore_index=True)
        
        # Remover Unnamed: 0 se existir
        if 'Unnamed: 0' in df_consolidado.columns:
            df_consolidado = df_consolidado.drop(columns=['Unnamed: 0'])
        
        print(f"\n{'=' * 80}")
        print("📊 ESTATÍSTICAS GERAIS")
        print("=" * 80)
        print(f"Total de registros consolidados: {len(df_consolidado)}")
        print(f"Total de colunas: {len(df_consolidado.columns)}")
        # Converter nomes de colunas para string para evitar erro com tipos mistos
        colunas_str = [str(col) for col in df_consolidado.columns.tolist()]
        print(f"Colunas: {', '.join(colunas_str)}\n")
        
        # Verificar se há nomes de colunas numéricos
        print("🔍 ANÁLISE DE NOMES DE COLUNAS:")
        colunas_numericas = []
        for col in df_consolidado.columns:
            if isinstance(col, (int, float)) and not isinstance(col, str):
                colunas_numericas.append(col)
                print(f"   ⚠️ Coluna com nome numérico encontrada: {col} (tipo: {type(col).__name__})")
        if not colunas_numericas:
            print("   ✅ Todos os nomes de colunas são strings")
        print()
        
        # ========== ANÁLISE DE TIPOS ==========
        print("=" * 80)
        print("🔬 ANÁLISE DE TIPOS DE DADOS")
        print("=" * 80)
        
        for coluna in df_consolidado.columns:
            print(f"\n📌 Coluna: '{coluna}'")
            print(f"   Tipo do pandas: {df_consolidado[coluna].dtype}")
            
            # Verificar tipos únicos dos valores
            tipos_unicos = set()
            valores_unicos = []
            
            for valor in df_consolidado[coluna].dropna().head(100):  # Limitar a 100 para performance
                tipo = type(valor).__name__
                tipos_unicos.add(tipo)
                if len(valores_unicos) < 10:  # Mostrar até 10 exemplos
                    valores_unicos.append(valor)
            
            # Converter tipos para string antes de ordenar
            tipos_str = sorted([str(t) for t in tipos_unicos])
            print(f"   Tipos encontrados: {', '.join(tipos_str)}")
            
            # Se há múltiplos tipos, é um problema!
            if len(tipos_unicos) > 1:
                print(f"   ⚠️ ATENÇÃO: Coluna com tipos mistos!")
                print(f"   Exemplos de valores: {valores_unicos[:5]}")
            
            # Estatísticas específicas
            nulos = df_consolidado[coluna].isna().sum()
            print(f"   Valores nulos: {nulos} ({nulos/len(df_consolidado)*100:.1f}%)")
            valores_unicos_total = df_consolidado[coluna].nunique()
            print(f"   Valores únicos: {valores_unicos_total}")
        
        # ========== ANÁLISE ESPECÍFICA DA COLUNA 'Especialidade' ==========
        print("\n" + "=" * 80)
        print("🏥 ANÁLISE DETALHADA: COLUNA 'Especialidade'")
        print("=" * 80)
        
        # Verificar se 'Especialidade' existe (pode estar como string ou int)
        coluna_especialidade = None
        for col in df_consolidado.columns:
            if str(col) == 'Especialidade':
                coluna_especialidade = col
                break
        
        if coluna_especialidade is None:
            print("❌ Coluna 'Especialidade' não encontrada!")
            print(f"   Colunas disponíveis: {[str(c) for c in df_consolidado.columns.tolist()]}")
            return
        
        print(f"   ✅ Coluna encontrada: '{coluna_especialidade}' (tipo do nome: {type(coluna_especialidade).__name__})")
        
        col_esp = df_consolidado[coluna_especialidade]
        
        # Identificar tipos de cada valor
        tipos_por_valor = {}
        valores_problematicos = []
        
        for idx, valor in col_esp.items():
            if pd.notna(valor):
                tipo_valor = type(valor).__name__
                valor_str = str(valor)
                
                if valor_str not in tipos_por_valor:
                    tipos_por_valor[valor_str] = []
                
                if tipo_valor not in tipos_por_valor[valor_str]:
                    tipos_por_valor[valor_str].append(tipo_valor)
                
                # Se o valor existe em múltiplos tipos, é problemático
                if len(tipos_por_valor[valor_str]) > 1:
                    if valor_str not in [v[0] for v in valores_problematicos]:
                        valores_problematicos.append((valor_str, tipos_por_valor[valor_str]))
        
        # Agrupar por tipo
        valores_por_tipo = {}
        for valor in col_esp.dropna():
            tipo = type(valor).__name__
            if tipo not in valores_por_tipo:
                valores_por_tipo[tipo] = []
            if str(valor) not in valores_por_tipo[tipo]:
                valores_por_tipo[tipo].append(str(valor))
        
        print(f"\n📊 Distribuição de tipos na coluna 'Especialidade':")
        for tipo, valores in valores_por_tipo.items():
            print(f"   {tipo}: {len(valores)} valores únicos")
            if len(valores) <= 10:
                print(f"      Exemplos: {', '.join(valores[:10])}")
            else:
                print(f"      Exemplos (primeiros 10): {', '.join(valores[:10])}")
        
        # Valores que aparecem em múltiplos tipos
        if valores_problematicos:
            print(f"\n⚠️ VALORES PROBLEMÁTICOS (aparecem em múltiplos tipos):")
            for valor, tipos in valores_problematicos[:20]:  # Mostrar até 20
                print(f"   '{valor}' → tipos: {tipos}")
        
        # Testar ordenação (causa do erro)
        print(f"\n🧪 TESTE DE ORDENAÇÃO:")
        try:
            valores_unicos = [e for e in col_esp.unique() if pd.notna(e)]
            valores_ordenados = sorted(valores_unicos)
            print(f"   ✅ Ordenação bem-sucedida!")
            print(f"   Total de valores únicos: {len(valores_unicos)}")
        except TypeError as e:
            print(f"   ❌ ERRO AO ORDENAR: {e}")
            print(f"   💡 Este é o problema! A coluna contém tipos mistos.")
            
            # Identificar quais valores causam o problema
            print(f"\n   🔍 Tentando identificar valores problemáticos:")
            valores_int = []
            valores_str = []
            
            for valor in valores_unicos:
                if isinstance(valor, (int, float)):
                    valores_int.append(valor)
                elif isinstance(valor, str):
                    valores_str.append(valor)
                else:
                    print(f"      Tipo desconhecido: {valor} ({type(valor)})")
            
            if valores_int:
                print(f"   📊 Valores numéricos encontrados: {len(valores_int)}")
                print(f"      Exemplos: {valores_int[:10]}")
            if valores_str:
                print(f"   📝 Valores string encontrados: {len(valores_str)}")
                print(f"      Exemplos: {valores_str[:10]}")
        
        # ========== ANÁLISE POR ABA ==========
        print("\n" + "=" * 80)
        print("📋 ANÁLISE POR ABA (para identificar origem do problema)")
        print("=" * 80)
        
        for aba in xls.sheet_names:
            if aba.startswith("Dia"):
                try:
                    df_aba = pd.read_excel(xls, sheet_name=aba)
                    # Verificar se 'Especialidade' existe (pode estar como string ou int)
                    coluna_esp_aba = None
                    for col in df_aba.columns:
                        if str(col) == 'Especialidade':
                            coluna_esp_aba = col
                            break
                    
                    if coluna_esp_aba is not None:
                        col_esp_aba = df_aba[coluna_esp_aba]
                        tipos_aba = set()
                        
                        for valor in col_esp_aba.dropna():
                            tipos_aba.add(type(valor).__name__)
                        
                        print(f"\n📌 Aba: '{aba}'")
                        # Converter tipos para string antes de ordenar
                        tipos_str_aba = sorted([str(t) for t in tipos_aba])
                        print(f"   Tipos encontrados: {', '.join(tipos_str_aba)}")
                        
                        if len(tipos_aba) > 1:
                            print(f"   ⚠️ PROBLEMA: Esta aba tem tipos mistos!")
                            
                            # Mostrar exemplos
                            valores_int = [v for v in col_esp_aba.dropna() if isinstance(v, (int, float))]
                            valores_str = [v for v in col_esp_aba.dropna() if isinstance(v, str)]
                            
                            if valores_int:
                                print(f"      Valores numéricos: {valores_int[:5]}")
                            if valores_str:
                                print(f"      Valores string: {valores_str[:5]}")
                except Exception as e:
                    print(f"   ❌ Erro ao analisar aba '{aba}': {e}")
        
        # ========== RECOMENDAÇÃO ==========
        print("\n" + "=" * 80)
        print("💡 RECOMENDAÇÃO DE CORREÇÃO")
        print("=" * 80)
        print("""
Para corrigir o problema, adicione esta linha na função carregar_dados()
logo após a linha 109 (após o fillna):

    df_consolidado['Especialidade'] = df_consolidado['Especialidade'].astype(str)
    df_consolidado['Especialidade'] = df_consolidado['Especialidade'].replace('nan', 'Não informado')

Isso garantirá que todos os valores sejam strings antes da ordenação.
        """)
        
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        caminho = sys.argv[1]
    else:
        print("Uso: python analisar_planilha.py <caminho_do_arquivo.xlsx>")
        print("\nOu informe o caminho do arquivo:")
        caminho = input("Caminho do arquivo: ").strip().strip('"')
    
    if not Path(caminho).exists():
        print(f"❌ Arquivo não encontrado: {caminho}")
        sys.exit(1)
    
    analisar_planilha(caminho)

