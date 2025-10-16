import logging
import pandas as pd
from config import HP, DIRECCION, FORMATO_SALIDA
from utils.numericos import convertir_europeo_a_float

logger = logging.getLogger(__name__)


def macro_hinges_list(archivos, hp=None, direccion=None):
    """
    MACRO 1: Hinges_List - Réplica exacta de la macro VBA
    Identifica rótulas únicas usando el algoritmo delta y calcula propiedades con interpolación
    """
    # Usar configuración global si no se proporcionan parámetros
    hp_actual = hp if hp is not None else HP
    direccion_actual = direccion if direccion is not None else DIRECCION

    print(f'Ejecutando Hinges_List (Hp={hp_actual}m, Dirección={direccion_actual})')

    df_cr = archivos['CR']
    df_cd = archivos['CD']
    df_sc = archivos['SC']

    # PASO 1: Calcular delta para saltos (líneas 15-21 VBA)
    dt = 4
    while (dt < len(df_cr) - 2 and
           convertir_europeo_a_float(df_cr.iloc[dt-1]['StepNum']) <
           convertir_europeo_a_float(df_cr.iloc[dt+1]['StepNum'])):
        dt += 2
    dt -= 2

    # PASO 2: Identificar rótulas únicas manteniendo orden de aparición
    rotulas_unicas = []
    rotulas_vistas = set()
    nr = 0  # Contador de rótulas

    # Procesar todas las filas de CR para mantener orden de aparición
    for i in range(len(df_cr)):
        if pd.isna(df_cr.iloc[i]['Frame']):
            continue

        rotula_actual = df_cr.iloc[i]['GenHinge']

        # Solo agregar si no la hemos visto antes
        if rotula_actual not in rotulas_vistas:
            rotula = {
                'Hinge': df_cr.iloc[i]['GenHinge'],
                'Section': df_cr.iloc[i]['AssignHinge'],
                'Frame': df_cr.iloc[i]['Frame'],
                'RelDist': convertir_europeo_a_float(df_cr.iloc[i]['RelDist'])
            }
            rotulas_unicas.append(rotula)
            rotulas_vistas.add(rotula_actual)
            nr += 1

    print(f'Identificadas {nr} rótulas únicas')

    # PASO 3: Crear DataFrame RT
    df_rt = pd.DataFrame(rotulas_unicas)

    # PASO 4: Agregar longitudes y coordenadas desde CD (líneas 60-85 VBA)
    for idx, row in df_rt.iterrows():
        frame_num = row['Frame']

        # Buscar frame en CD
        frame_data = df_cd[df_cd['Frame'] == frame_num]
        if not frame_data.empty:
            frame_info = frame_data.iloc[0]
            length = convertir_europeo_a_float(frame_info['Length'])
            centroid_z = convertir_europeo_a_float(frame_info['CentroidZ'])

            df_rt.at[idx, 'L'] = length
            # Fórmula VBA: (CentroidY + Hp/2) / Hp
            df_rt.at[idx, 'Storey'] = (centroid_z + hp_actual/2) / hp_actual

    # PASO 5: Identificar P average (P en t=0) (líneas 87-96 VBA)
    # Buscar P average por nombre de rótula, no por índice secuencial
    for i in range(len(df_cr)):
        if convertir_europeo_a_float(df_cr.iloc[i]['StepNum']) == 0:
            hinge_name = df_cr.iloc[i]['GenHinge']
            p_value = convertir_europeo_a_float(df_cr.iloc[i]['P']) * -1

            # Buscar la rótula en RT y asignar P average
            rt_idx = df_rt[df_rt['Hinge'] == hinge_name].index
            if not rt_idx.empty:
                df_rt.at[rt_idx[0], 'P average'] = p_value

    # PASO 7: Interpolación My, Mu, Cy, Cu (líneas 110-185 VBA)
    for i in range(len(df_rt)):
        section_name = df_rt.iloc[i]['Section']
        axial = abs(df_rt.iloc[i]['P average'])
        hinge_name = df_rt.iloc[i]['Hinge']

        try:
            lc = df_rt.iloc[i]['L']  # Longitud crítica

            # Buscar sección en SC
            section_data = df_sc[df_sc['Section'] == section_name]
            if section_data.empty:
                logger.warning(f"Sección no encontrada: {section_name}")
                continue

            sec = section_data.iloc[0]

            # Propiedades básicas de la sección
            if direccion_actual.upper() == 'X':
                b = convertir_europeo_a_float(sec['H'])
                h = convertir_europeo_a_float(sec['B'])
                psx = convertir_europeo_a_float(sec['?sx'])  # ρsx aparece como ?sx
            else:
                b = convertir_europeo_a_float(sec['B'])
                h = convertir_europeo_a_float(sec['H'])
                psx = convertir_europeo_a_float(sec['?sy'])  # ρsy aparece como ?sy

            fc = convertir_europeo_a_float(sec["f'c"])
            fy = convertir_europeo_a_float(sec['fyw'])
            alpha = convertir_europeo_a_float(sec['?'])   # α aparece como ?

            df_rt.at[i, 'B'] = b
            df_rt.at[i, 'H'] = h
            df_rt.at[i, "f'c"] = fc
            df_rt.at[i, 'fy'] = fy
            df_rt.at[i, 'ρsx'] = psx
            df_rt.at[i, 'α'] = alpha
            df_rt.at[i, 'Lc*'] = lc

            # Interpolación exacta VBA usando nombres de columnas
            # Buscar el rango de interpolación correcto
            p_values = []
            for p_idx in range(1, 4):  # P(1) a P(4)
                p_col = f'P({p_idx})'
                if p_col in sec.index:
                    p_val = convertir_europeo_a_float(sec[p_col])
                    p_values.append((p_idx, p_val))

            # Encontrar el rango de interpolación usando lógica exacta VBA
            # VBA: Do While ThisWorkbook.Sheets("SC").Cells(j, 9 * k + 1).Value < Axial
            pi = pf = myi = myf = mui = muf = cyi = cyf = cui = cuf = 0
            k = 1  # Índice de interpolación (1-based como en VBA)

            # Definir offset t según dirección (exacto como VBA)
            if direccion_actual.upper() == 'X':  # Dirección 90
                t = 4  # Offset para columnas (90)
            else:  # Dirección Y, dirección 00
                t = 0  # Offset para columnas (00)

            # Buscar rango de interpolación donde Pi <= Axial <= Pf
            # VBA: Do While ThisWorkbook.Sheets("SC").Cells(j, 9 * k + 1).Value < Axial
            p_col_k = f'P({k})'
            p_k = convertir_europeo_a_float(sec[p_col_k])

            while p_k < axial:  # Máximo 4 puntos P(1) a P(4)
                p_col_k_plus_1 = f'P({k+1})'
                p_k_plus_1 = convertir_europeo_a_float(sec[p_col_k_plus_1])
                # Rango encontrado
                pi = p_k
                pf = p_k_plus_1

                # Obtener valores usando fórmula exacta VBA: 9 * k + t + offset
                # Columnas SC: My=+2, Mu=+3, Cy=+4, Cu=+5 (relativo a P)
                # Como pandas usa 0-based indexing y usamos index_col=0, ajustamos: col_index = 9 * k + t + offset - 2

                # Para k-ésimo punto (Pi)
                my_col_i_idx = 9 * k + t + 2 - 2   # VBA: 9*k + t + 2, ajustado para pandas con index_col=0
                mu_col_i_idx = 9 * k + t + 3 - 2
                cy_col_i_idx = 9 * k + t + 4 - 2
                cu_col_i_idx = 9 * k + t + 5 - 2

                # Para (k+1)-ésimo punto (Pf)
                my_col_f_idx = 9 * (k + 1) + t + 2 - 2
                mu_col_f_idx = 9 * (k + 1) + t + 3 - 2
                cy_col_f_idx = 9 * (k + 1) + t + 4 - 2
                cu_col_f_idx = 9 * (k + 1) + t + 5 - 2

                # Obtener valores usando índices de columna de la fila sec (que es una Series)
                # sec es section_data.iloc[0], usamos .iloc para acceso posicional en la Series
                try:
                    myi = convertir_europeo_a_float(sec.iloc[my_col_i_idx])
                    myf = convertir_europeo_a_float(sec.iloc[my_col_f_idx])
                    mui = convertir_europeo_a_float(sec.iloc[mu_col_i_idx])
                    muf = convertir_europeo_a_float(sec.iloc[mu_col_f_idx])
                    cyi = convertir_europeo_a_float(sec.iloc[cy_col_i_idx])
                    cyf = convertir_europeo_a_float(sec.iloc[cy_col_f_idx])
                    cui = convertir_europeo_a_float(sec.iloc[cu_col_i_idx])
                    cuf = convertir_europeo_a_float(sec.iloc[cu_col_f_idx])
                except IndexError as e:
                    logger.warning(f"Error accediendo índices para sección {rotula['Section']}: {e}")
                    logger.debug(f"Índices calculados: my_i={my_col_i_idx}, my_f={my_col_f_idx}")
                    logger.debug(f"Longitud de sec: {len(sec)}")
                    continue
                k += 1
                p_col_k = f'P({k})'
                p_k = convertir_europeo_a_float(sec[p_col_k])

            # Interpolación lineal exacta VBA (líneas 176-194)

            m = (myf - myi) / (pf - pi)
            bo = myi - m * pi
            my = m * axial + bo

            # Mu (líneas 181-183)
            m = (muf - mui) / (pf - pi)
            bo = mui - m * pi
            mu = m * axial + bo

            # Cy (líneas 185-187)
            m = (cyf - cyi) / (pf - pi)
            bo = cyi - m * pi
            cy = m * axial + bo

            # Cu (líneas 189-191)
            m = (cuf - cui) / (pf - pi)
            bo = cui - m * pi
            cu = m * axial + bo

            # Lp y Rp (líneas 193-194)
            lp = 0.5 * h
            rp = (cu - cy) * lp

            # Asignar valores calculados
            df_rt.at[i, 'My'] = my
            df_rt.at[i, 'Mu'] = mu
            df_rt.at[i, 'Cy'] = cy
            df_rt.at[i, 'Cu'] = cu
            df_rt.at[i, 'Rp'] = rp

        except (KeyError, ValueError):
            continue

    print(f'RT generado con {len(df_rt)} rótulas')
    return df_rt
