#!/usr/bin/env python3
"""
Script de prueba para verificar la estructura del DataFrame SC
"""

import pandas as pd

# Cargar SC.csv con index_col=0
df_sc = pd.read_csv('csv/SC.csv', sep=';', encoding='utf-8', index_col=0)

print("Estructura del DataFrame SC:")
print(f"Forma: {df_sc.shape}")
print(f"Columnas: {list(df_sc.columns)}")
print(f"Primeras 5 columnas: {list(df_sc.columns[:5])}")

# Verificar la primera fila (C1A)
primera_fila = df_sc.iloc[0]
print(f"\nPrimera fila (C1A):")
print(f"Section: {primera_fila.name}")  # El índice debería ser C1A
print(f"B: {primera_fila['B']}")
print(f"H: {primera_fila['H']}")

# Verificar las columnas P(1), P(2), etc.
p_columns = [col for col in df_sc.columns if col.startswith('P(')]
print(f"\nColumnas P encontradas: {p_columns}")

# Verificar las columnas My
my_columns = [col for col in df_sc.columns if 'My' in col]
print(f"Columnas My encontradas: {my_columns}")

# Verificar índices de columnas específicas
print(f"\nÍndices de columnas importantes:")
for i, col in enumerate(df_sc.columns):
    if col in ['Section', 'B', 'H', 'P(1)', 'My (00)', 'Mu (00)', 'P(2)']:
        print(f"  {col}: índice {i}")
    if i > 20:  # Limitar salida
        break
