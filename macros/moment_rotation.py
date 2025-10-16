import logging
import pandas as pd
from config import DIRECCION
from utils.numericos import convertir_europeo_a_float

logger = logging.getLogger(__name__)


def macro_moment_rotation(archivos, df_rt, direccion=None):
    """
    MACRO 2: Moment_Rotation - Réplica exacta de la macro VBA
    Calcula curvas momento-rotación y energía histerética usando algoritmo delta
    """
    # Usar configuración global si no se proporciona parámetro
    direccion_actual = direccion if direccion is not None else DIRECCION

    print(f'Ejecutando Moment_Rotation (Dirección={direccion_actual})')

    df_cr = archivos['CR']
    df_hk = archivos['HK']

    # PASO 1: Contar número de rótulas (líneas 8-13 VBA)
    no_rotulas = len(df_rt)

    # PASO 2: Crear estructura matricial MR (líneas 14-25 VBA)
    mr_matrix_data = {}
    for i in range(no_rotulas):
        hinge_name = df_rt.iloc[i]['Hinge']
        j = 3 * (i + 1) - 1  # Columna base (3*i-1 en VBA)
        mr_matrix_data[hinge_name] = {
            'column_base': j,
            'moments': [],
            'rotations': [],
            'axials': []
        }

    # PASO 3: Identificar dirección del sismo (líneas 27-36 VBA)
    sismo_valor = 1 if direccion_actual.upper() == 'X' else 2

    if sismo_valor == 1:  # Sismo X
        ck_col = 'Kx'  # Rigidez Kx en HK
    else:  # Sismo Y
        ck_col = 'Ky'  # Rigidez Ky en HK

    # PASO 4: Organizar datos por rótula usando algoritmo delta (líneas 38-65 VBA)
    i = 3  # Índice en CR (base 0, VBA usa base 1)
    j = 7  # Fila en MR (base 0, VBA usa base 1)
    # Nota: La variable 'k' de la macro VBA no es necesaria en esta implementación      

    # Calcular delta como en macro_hinges_list
    dt = 4
    while (dt < len(df_cr) - 2 and
           convertir_europeo_a_float(df_cr.iloc[dt-1]['StepNum']) <
           convertir_europeo_a_float(df_cr.iloc[dt+1]['StepNum'])):
        dt += 2
    dt -= 2

    # Organizar datos matriciales
    while i < len(df_cr) and not pd.isna(df_cr.iloc[i]['Frame']):
        if (i + 1 < len(df_cr) and 
            convertir_europeo_a_float(df_cr.iloc[i]['StepNum']) < 
            convertir_europeo_a_float(df_cr.iloc[i+2]['StepNum']) if i+2 < len(df_cr) else True):

            # Procesar paso de tiempo actual
            for step_idx in range(j-7, len(df_cr), dt):
                if step_idx >= len(df_cr):
                    break

                # Datos para H1 (rótula 1)
                if step_idx < len(df_cr):
                    moment_h1 = convertir_europeo_a_float(df_cr.iloc[step_idx]['M3' if sismo_valor == 1 else 'M2'])
                    rotation_h1 = convertir_europeo_a_float(df_cr.iloc[step_idx]['R3Plastic' if sismo_valor == 1 else 'R2Plastic'])
                    axial_h1 = convertir_europeo_a_float(df_cr.iloc[step_idx]['P'])

                    hinge_h1 = df_cr.iloc[step_idx]['GenHinge']
                    if hinge_h1 in mr_matrix_data:
                        mr_matrix_data[hinge_h1]['moments'].append(moment_h1)
                        mr_matrix_data[hinge_h1]['rotations'].append(rotation_h1)
                        mr_matrix_data[hinge_h1]['axials'].append(axial_h1)

                # Datos para H2 (rótula 2)
                if step_idx + 1 < len(df_cr):
                    moment_h2 = convertir_europeo_a_float(df_cr.iloc[step_idx+1]['M3' if sismo_valor == 1 else 'M2'])
                    rotation_h2 = convertir_europeo_a_float(df_cr.iloc[step_idx+1]['R3Plastic' if sismo_valor == 1 else 'R2Plastic'])
                    axial_h2 = convertir_europeo_a_float(df_cr.iloc[step_idx+1]['P'])

                    hinge_h2 = df_cr.iloc[step_idx+1]['GenHinge']
                    if hinge_h2 in mr_matrix_data:
                        mr_matrix_data[hinge_h2]['moments'].append(moment_h2)
                        mr_matrix_data[hinge_h2]['rotations'].append(rotation_h2)
                        mr_matrix_data[hinge_h2]['axials'].append(axial_h2)

            j = 7
        else:
            j += 1

        i += dt

    # PASO 5: Procesar rotaciones para cada rótula (líneas 67-160 VBA)
    for idx, rotula in df_rt.iterrows():
        hinge_name = rotula['Hinge']

        if hinge_name not in mr_matrix_data or not mr_matrix_data[hinge_name]['moments']:
            continue

        # Obtener rigidez Kmr desde HK (líneas 81-95 VBA)
        hinge_hk = df_hk[df_hk['Hinge Name'] == hinge_name]
        if not hinge_hk.empty:
            kmr = convertir_europeo_a_float(hinge_hk.iloc[0][ck_col])
        else:
            logger.warning(f"No se encontró rótula {hinge_name} en HK")
            continue

        # Calcular rotaciones características (líneas 97-110 VBA)
        my = convertir_europeo_a_float(rotula['My'])
        ry = my / kmr if kmr != 0 else 0
        rp = convertir_europeo_a_float(rotula['Rp'])
        ru = ry + rp

        # Calcular rotación de fisuración Rc (líneas 112-118 VBA)
        fc = convertir_europeo_a_float(rotula["f'c"]) 
        b = convertir_europeo_a_float(rotula['B'])
        h = convertir_europeo_a_float(rotula['H'])

        fr = 0.7 * 1 * (fc ** 0.5)
        mr_ruptura = (1/6) * (fr * 1000) * b * (h ** 2)
        rc = (ry / my) * mr_ruptura if my != 0 else 0

        # PASO 6: Calcular rotación total con factor de escala (líneas 120-160 VBA)
        moments = mr_matrix_data[hinge_name]['moments']
        rotations_plastic = mr_matrix_data[hinge_name]['rotations']

        re = 0  # Rotación elástica acumulada
        rotations_total = []

        for idx, (mi, rp_val) in enumerate(zip(moments, rotations_plastic)):
            if idx == 0:
                # Primer paso: Re = (Mi * 10^6) / Kmr
                re = (mi * 1e6) / kmr if kmr != 0 else 0
            else:
                # Pasos siguientes: calcular incremento
                mj = moments[idx-1]
                dm = (mi - mj) * 1e6

                if mi > 0:  # Momento positivo
                    if my != 0 and mi / my < 1:
                        dre = dm / kmr if kmr != 0 else 0
                    else:
                        dre = 0 if dm > 0 else dm / kmr if kmr != 0 else 0
                else:  # Momento negativo
                    if my != 0 and abs(mi) / my < 1:
                        dre = dm / kmr if kmr != 0 else 0
                    else:
                        dre = 0 if dm < 0 else dm / kmr if kmr != 0 else 0

                re += dre

            # Rotación total = Re + Rp * 10^6
            rt = re + rp_val * 1e6
            rotations_total.append(rt)

        # PASO 7: Escalar rotaciones (líneas 152-158 VBA)
        rotations_total = [rt * 1e-6 for rt in rotations_total]
        mr_matrix_data[hinge_name]['rotations'] = rotations_total

        # PASO 8: Calcular energía histerética y máximos (líneas 162-220 VBA)
        eh = 0
        rmax = 0
        rmin = 0
        pmin = 0

        for j in range(1, len(moments)):
            # Energía histerética incremental
            dm_avg = 0.5 * (moments[j] + moments[j-1])
            dr = rotations_total[j] - rotations_total[j-1]
            deh = dm_avg * dr
            eh += deh

            # Máximos y mínimos
            if rotations_total[j] > rmax:
                rmax = rotations_total[j]
            if rotations_total[j] < rmin:
                rmin = rotations_total[j]
            if mr_matrix_data[hinge_name]['axials'][j] < pmin:
                pmin = mr_matrix_data[hinge_name]['axials'][j]

        # Determinar rotación máxima absoluta
        if abs(rmin) > rmax:
            rmax = abs(rmin)

        # Ajustar energía histerética (líneas 222-240 VBA)
        if eh < 0:
            if eh < -2:
                logger.warning(f"EH negativa para {hinge_name}: {eh}")
            eh = 0

        if rmax < ry:
            eh = 0

        # Actualizar DataFrame RT
        df_rt.at[idx, 'θy'] = ry
        df_rt.at[idx, 'θp'] = rp
        df_rt.at[idx, 'θu'] = ru
        df_rt.at[idx, 'θc'] = rc
        df_rt.at[idx, 'θm'] = rmax
        df_rt.at[idx, 'Pm'] = abs(pmin)
        df_rt.at[idx, 'EH'] = eh

    print(f'Moment_Rotation completado para {len(df_rt)} rótulas')
    return mr_matrix_data, df_rt
