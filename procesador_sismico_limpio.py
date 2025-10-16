"""
Procesador Sísmico - Conversión exacta de macros VBA a Python
Sigue la lógica exacta de las macros: Hinges_List, Moment_Rotation, Damage_Index
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime
import logging
from config import (
    HP, DIRECCION, LOG_LEVEL, ARCHIVOS_ENTRADA, FORMATO_SALIDA, 
    TOLERANCIA_NUMERICA, PRECISION_DECIMAL, validar_configuracion, mostrar_configuracion
)
from macros.hinges_list import macro_hinges_list
from macros.moment_rotation import macro_moment_rotation
from macros.damage_index import macro_damage_index

# Configurar logging usando configuración global
log_level = getattr(logging, LOG_LEVEL.upper(), logging.INFO)
logging.basicConfig(level=log_level, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

    

def cargar_archivos():
    """Carga los 4 archivos de entrada usando configuración global"""
    archivos = {}
    
    logger.info("📂 Cargando archivos de entrada desde configuración...")
    
    # Cargar archivos usando rutas de configuración
    for nombre, ruta in ARCHIVOS_ENTRADA.items():
        try:
            if nombre == 'SC':
                # SC.csv tiene una columna vacía al inicio, la usamos como índice
                archivos[nombre] = pd.read_csv(ruta, sep=FORMATO_SALIDA['separador'], 
                                             encoding=FORMATO_SALIDA['encoding'], 
                                             index_col=0)
            else:
                archivos[nombre] = pd.read_csv(ruta, sep=FORMATO_SALIDA['separador'], encoding=FORMATO_SALIDA['encoding'])
            logger.debug(f"✅ {nombre}.csv cargado: {len(archivos[nombre])} filas")
        except FileNotFoundError:
            logger.error(f"❌ Archivo no encontrado: {ruta}")
            raise
        except Exception as e:
            logger.error(f"❌ Error cargando {nombre}: {str(e)}")
            raise
    
    return archivos

def crear_mr_matricial(mr_matrix_data, rotulas_ordenadas, archivo_salida):
    """Crea archivo MR.csv en formato matricial usando configuración global"""
    
    # Crear encabezados
    header_row1 = []
    header_row2 = []
    header_row3 = []
    
    for rotula in rotulas_ordenadas:
        header_row1.extend([rotula, '', ''])
        header_row2.extend(['M', 'Rot', 'P'])
        header_row3.extend(['kN-m', 'Rad', 'kN'])
    
    # Determinar número máximo de filas
    max_rows = 0
    for rotula in rotulas_ordenadas:
        if rotula in mr_matrix_data:
            max_rows = max(max_rows, len(mr_matrix_data[rotula]['moments']))
    
    # Crear filas de datos
    data_rows = []
    for i in range(max_rows):
        row = []
        for rotula in rotulas_ordenadas:
            if rotula in mr_matrix_data and i < len(mr_matrix_data[rotula]['moments']):
                moment = mr_matrix_data[rotula]['moments'][i]
                rotation = mr_matrix_data[rotula]['rotations'][i]
                axial = mr_matrix_data[rotula]['axials'][i]
                
                # Formatear números usando configuración global
                moment_str = f"{moment:.{PRECISION_DECIMAL}f}".replace('.', FORMATO_SALIDA['decimal'])
                rotation_str = f"{rotation:.{PRECISION_DECIMAL}f}".replace('.', FORMATO_SALIDA['decimal'])
                axial_str = f"{axial:.2f}".replace('.', FORMATO_SALIDA['decimal'])
                
                row.extend([moment_str, rotation_str, axial_str])
            else:
                row.extend(['', '', ''])
        data_rows.append(row)
    
    # Escribir archivo CSV usando configuración global
    with open(archivo_salida, 'w', encoding=FORMATO_SALIDA['encoding']) as f:
        f.write(FORMATO_SALIDA['separador'].join(header_row1) + '\n')
        f.write(FORMATO_SALIDA['separador'].join(header_row2) + '\n')
        f.write(FORMATO_SALIDA['separador'].join(header_row3) + '\n')
        
        for row in data_rows:
            f.write(FORMATO_SALIDA['separador'].join(row) + '\n')

def guardar_resultados(df_rt, mr_matrix_data, df_id, carpeta_salida):
    """Guarda los archivos de salida RT, MR, ID usando configuración global"""
    
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
    
    # Guardar RT.csv usando configuración global
    archivo_rt = os.path.join(carpeta_salida, 'RT.csv')
    df_rt.to_csv(archivo_rt, 
                 sep=FORMATO_SALIDA['separador'], 
                 index=False, 
                 encoding=FORMATO_SALIDA['encoding'],
                 float_format=f'%.{PRECISION_DECIMAL}f')
    logger.info(f"RT.csv guardado: {archivo_rt}")
    
    # Guardar MR.csv
    archivo_mr = os.path.join(carpeta_salida, 'MR.csv')
    rotulas_ordenadas = df_rt['Hinge'].tolist()
    crear_mr_matricial(mr_matrix_data, rotulas_ordenadas, archivo_mr)
    logger.info(f"MR.csv guardado: {archivo_mr}")
    
    # Guardar ID.csv usando configuración global
    archivo_id = os.path.join(carpeta_salida, 'ID.csv')
    df_id.to_csv(archivo_id, 
                 sep=FORMATO_SALIDA['separador'], 
                 index=False, 
                 encoding=FORMATO_SALIDA['encoding'],
                 float_format=f'%.{PRECISION_DECIMAL}f')
    logger.info(f"ID.csv guardado: {archivo_id}")

def procesar_analisis_sismico(hp=None, direccion=None):
    """Función principal que ejecuta las 3 macros en secuencia"""
    # Usar configuración global si no se proporcionan parámetros
    hp_actual = hp if hp is not None else HP
    direccion_actual = direccion if direccion is not None else DIRECCION
    
    # Mostrar configuración actual
    mostrar_configuracion()
    
    
    # Crear carpeta de salida con timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    carpeta_salida = f'resultados_{timestamp}'
    
    try:
        # Cargar archivos de entrada
        archivos = cargar_archivos()
        
        # MACRO 1: Hinges_List
        df_rt = macro_hinges_list(archivos, hp_actual, direccion_actual)
        
        # MACRO 2: Moment_Rotation
        mr_matrix_data, df_rt = macro_moment_rotation(archivos, df_rt, direccion_actual)
        
        # MACRO 3: Damage_Index
        df_id, df_rt = macro_damage_index(df_rt)
        
        # Guardar resultados
        guardar_resultados(df_rt, mr_matrix_data, df_id, carpeta_salida)
        
        print(f'Análisis completado - Resultados en: {carpeta_salida}')
        print(f'Rótulas procesadas: {len(df_rt)}')
        
        # Distribución de niveles de desempeño
        if len(df_id) > 0:
            distribucion = df_id['DS'].value_counts()
            for nivel, cantidad in distribucion.items():
                niveles = {'TO': 'Totalmente operativo', 'IO': 'Ocupación inmediata', 'LS': 'Seguridad de la vida', 'CP': 'Prevención de colapso', 'CL': 'Colapso'}
                print(f'{nivel} ({niveles.get(nivel, nivel)}): {cantidad} rótulas')
        
        return {
            'RT': df_rt,
            'MR': mr_matrix_data,
            'ID': df_id,
            'carpeta_salida': carpeta_salida
        }
        
    except Exception as e:
        print(f"Error en análisis: {str(e)}")
        raise

if __name__ == "__main__":
    # Ejecutar análisis usando configuración global
    resultados = procesar_analisis_sismico()
