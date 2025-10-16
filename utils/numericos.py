import pandas as pd
from config import TOLERANCIA_NUMERICA


def convertir_europeo_a_float(valor):
    """Convierte formato numérico europeo (coma decimal) a float"""
    if pd.isna(valor) or valor == '' or valor is None:
        return 0.0
    try:
        str_val = str(valor).replace(',', '.')
        result = float(str_val)
        # Aplicar tolerancia numérica para evitar errores de precisión
        if abs(result) < TOLERANCIA_NUMERICA:
            return 0.0
        return result
    except (ValueError, TypeError):
        return 0.0
