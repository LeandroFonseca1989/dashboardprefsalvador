"""
Script para explorar a estrutura do arquivo Excel
"""
import pandas as pd

# Nome do arquivo
arquivo = "Estudo de produtividade Unidade de Saúda da Familia - Sao Cristovao.xlsx"

# Ler todas as abas
xls = pd.ExcelFile(arquivo)

print("=" * 60)
print("ABAS ENCONTRADAS NO ARQUIVO:")
print("=" * 60)
for i, aba in enumerate(xls.sheet_names, 1):
    print(f"{i}. {aba}")

print("\n" + "=" * 60)
print("EXAMINANDO A PRIMEIRA ABA:")
print("=" * 60)

# Ler a primeira aba
df_primeira = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
print(f"\nNome da aba: {xls.sheet_names[0]}")
print(f"Formato: {df_primeira.shape[0]} linhas x {df_primeira.shape[1]} colunas")
print(f"\nColunas encontradas:")
for col in df_primeira.columns:
    print(f"  - {col}")
print(f"\nPrimeiras 5 linhas:")
print(df_primeira.head())
print(f"\nTipos de dados:")
print(df_primeira.dtypes)

