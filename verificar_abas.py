"""
Script para verificar exatamente como os nomes das abas estão no arquivo Excel
e verificar as referências de mês nas abas
"""
import pandas as pd

# Nome do arquivo
arquivo = "Estudo de produtividade Unidade de Saúda da Familia - Sao Cristovao.xlsx"

# Ler todas as abas
xls = pd.ExcelFile(arquivo)

print("=" * 60)
print("TODAS AS ABAS ENCONTRADAS:")
print("=" * 60)
for i, aba in enumerate(xls.sheet_names, 1):
    print(f"{i}. '{aba}' (tipo: {type(aba).__name__})")
    
    # Verificar se contém "09" ou "9"
    if "09" in aba or "9" in aba:
        print(f"   ⚠️ Esta aba contém '09' ou '9'")
        # Tentar ler a aba
        try:
            df_test = pd.read_excel(xls, sheet_name=aba, nrows=2)
            print(f"   ✅ Aba lida com sucesso - {len(df_test)} linhas")
        except Exception as e:
            print(f"   ❌ Erro ao ler aba: {e}")

print("\n" + "=" * 60)
print("ABAS QUE COMEÇAM COM 'Dia':")
print("=" * 60)
abas_dia = [aba for aba in xls.sheet_names if aba.startswith("Dia")]
for aba in sorted(abas_dia):
    print(f"  - '{aba}'")

print("\n" + "=" * 60)
print("ABAS DE MÊS (verificando referência na coluna A, linha 1):")
print("=" * 60)

# Lista de nomes de meses em português
meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
         'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Identificar abas que não são "Dia", "Consolidado" ou outras abas conhecidas
abas_conhecidas = ['Consolidado']
abas_mes = []

for aba in xls.sheet_names:
    # Se não começa com "Dia" e não é "Consolidado", pode ser uma aba de mês
    if not aba.startswith("Dia") and aba not in abas_conhecidas:
        abas_mes.append(aba)

if len(abas_mes) > 0:
    for aba in sorted(abas_mes):
        print(f"\n📅 Aba: '{aba}'")
        try:
            # Ler apenas a primeira linha da coluna A
            df_mes = pd.read_excel(xls, sheet_name=aba, nrows=1, usecols=[0], header=None)
            if len(df_mes) > 0:
                valor_celula_a1 = df_mes.iloc[0, 0]
                print(f"   ✅ Coluna A, Linha 1: '{valor_celula_a1}'")
                print(f"   ✅ Tipo do valor: {type(valor_celula_a1).__name__}")
                
                # Tentar extrair o mês
                mes_extraido = None
                
                # Se for uma data (Timestamp ou datetime)
                if pd.isna(valor_celula_a1):
                    print(f"   ⚠️ Valor é NaN")
                elif isinstance(valor_celula_a1, pd.Timestamp) or hasattr(valor_celula_a1, 'month'):
                    mes_numero = valor_celula_a1.month
                    mes_extraido = meses[mes_numero - 1]
                    print(f"   ✅ Data detectada! Mês extraído: {mes_extraido} (mês {mes_numero})")
                else:
                    # Tentar converter string para data
                    try:
                        data = pd.to_datetime(valor_celula_a1)
                        mes_numero = data.month
                        mes_extraido = meses[mes_numero - 1]
                        print(f"   ✅ Data detectada na string! Mês extraído: {mes_extraido} (mês {mes_numero})")
                    except:
                        # Verificar se é um mês em texto
                        valor_str = str(valor_celula_a1).strip()
                        if valor_str in meses:
                            mes_extraido = valor_str
                            print(f"   ✅ Mês em texto identificado: {mes_extraido}")
                        else:
                            print(f"   ⚠️ Valor não é uma data nem um mês conhecido")
            else:
                print(f"   ⚠️ Aba vazia ou sem dados na primeira linha")
        except Exception as e:
            print(f"   ❌ Erro ao ler aba: {e}")
else:
    print("⚠️ Nenhuma aba de mês encontrada (abas que não começam com 'Dia' e não são 'Consolidado')")

