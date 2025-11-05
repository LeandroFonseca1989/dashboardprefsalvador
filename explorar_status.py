"""
Script para explorar os valores de Status nos dados
"""
import pandas as pd

# Nome do arquivo
arquivo = "Estudo de produtividade Unidade de Saúda da Familia - Sao Cristovao.xlsx"

# Ler todas as abas
xls = pd.ExcelFile(arquivo)

# Consolidar dados de todas as abas "Dia"
dados_consolidados = []

for aba in xls.sheet_names:
    if aba.startswith("Dia"):
        df = pd.read_excel(xls, sheet_name=aba)
        df['Dia'] = aba
        dados_consolidados.append(df)

# Concatenar todos
df_consolidado = pd.concat(dados_consolidados, ignore_index=True)

print("=" * 60)
print("VALORES ÚNICOS DE STATUS:")
print("=" * 60)
print(df_consolidado['Status'].value_counts())

print("\n" + "=" * 60)
print("VALORES ÚNICOS DE ESPECIALIDADE (EQUIPE):")
print("=" * 60)
print(df_consolidado['Especialidade'].value_counts())

print("\n" + "=" * 60)
print("PROFISSIONAIS ÚNICOS:")
print("=" * 60)
print(df_consolidado['Profissional'].unique())

print("\n" + "=" * 60)
print("TOTAL DE REGISTROS:")
print("=" * 60)
print(f"Total: {len(df_consolidado)} registros")

