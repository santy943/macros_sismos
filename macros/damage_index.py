import logging
import pandas as pd
from utils.numericos import convertir_europeo_a_float

logger = logging.getLogger(__name__)


def macro_damage_index(df_rt):
    """
    MACRO 3: Damage_Index - Réplica exacta de la macro VBA
    Seismic Damage Assessment basado en Jiang, Chen & Chen (2011)
    """
    print('Ejecutando Damage_Index')

    # Crear DataFrame ID con columnas necesarias
    df_id = df_rt[['Hinge', 'Section', 'Frame', 'Storey', 'EH']].copy()

    # Agregar columnas para parámetros calculados
    df_rt['no'] = 0.0
    df_rt['Beta'] = 0.0
    df_rt['ID'] = 0.0
    df_rt['ND'] = ''

    # PASO 1: Calcular parámetros para cada rótula (líneas 9-29 VBA)
    for idx, rotula in df_rt.iterrows():
        # Obtener todos los parámetros necesarios desde RT
        b = convertir_europeo_a_float(rotula['B'])
        h = convertir_europeo_a_float(rotula['H'])
        fc = convertir_europeo_a_float(rotula["f'c"]) 
        fy = convertir_europeo_a_float(rotula['fy'])
        psx = convertir_europeo_a_float(rotula['ρsx'])
        a = convertir_europeo_a_float(rotula['α'])
        my = convertir_europeo_a_float(rotula['My'])
        lc = convertir_europeo_a_float(rotula['Lc*'])
        ry = convertir_europeo_a_float(rotula['θy'])
        ru = convertir_europeo_a_float(rotula['θu'])
        rc = convertir_europeo_a_float(rotula['θc'])
        rm = convertir_europeo_a_float(rotula['θm'])
        p = convertir_europeo_a_float(rotula['Pm'])
        eh = convertir_europeo_a_float(rotula['EH'])

        # PASO 2: Evaluación de parámetros (líneas 27-29 VBA)
        # no = P / (b * h * fc * 10^3)
        if b > 0 and h > 0 and fc > 0:
            no = p / (b * h * fc * 1000)
        else:
            no = 0.0

        # Beta = (0.023 * Lc/h + 3.352 * no^2.35) * 0.818^(a * psx * fy/fc * 100) + 0.039
        if h > 0 and fc > 0:
            term1 = (0.023 * lc / h + 3.352 * (no ** 2.35)) *   0.818 
            exponent = a * psx * fy / fc * 100
            term2 = term1 ** exponent
            beta = term2 + 0.039
        else:
            beta = 0.039  # Valor mínimo

        # PASO 3: Calcular dM (líneas 31-35 VBA)
        if rm > rc:
            dm = rm
        else:
            dm = rc

        # PASO 4: Evaluación del índice de daño (líneas 37-38 VBA)
        # ID = (1-Beta)*(dM-Rc)/(Ru-Rc) + Beta*EH/(My*(Ru-Ry))
        if ru > rc and my > 0 and (ru - ry) > 0:
            id_value = (1 - beta) * (dm - rc) / (ru - rc) + beta * eh / (my * (ru - ry))
        else:
            id_value = 0.0

        # Asegurar que ID no sea negativo
        if id_value < 0:
            id_value = 0.0

        # PASO 5: Nivel de desempeño (líneas 40-58 VBA)
        if id_value < 0.05:
            nd = "TO"  # Totalmente operativo
        elif id_value < 0.15:
            nd = "IO"  # Ocupación inmediata
        elif id_value < 0.45:
            nd = "LS"  # Seguridad de la vida
        elif id_value < 1.0:
            nd = "CP"  # Prevención de colapso
        else:
            nd = "CL"  # Colapso

        # Actualizar DataFrame RT con resultados (líneas 60-64 VBA)
        df_rt.at[idx, 'no'] = no
        df_rt.at[idx, 'Beta'] = beta
        df_rt.at[idx, 'ID'] = id_value
        df_rt.at[idx, 'ND'] = nd

    # PASO 6: Crear DataFrame ID final (líneas 66-85 VBA)
    for idx, rotula in df_rt.iterrows():
        id_value = rotula['ID']
        ds = rotula['ND']

        # Actualizar DataFrame ID
        df_id.at[idx, 'ID'] = id_value
        df_id.at[idx, 'DS'] = ds

    print(f'Damage_Index completado para {len(df_rt)} rótulas')
    return df_id, df_rt
