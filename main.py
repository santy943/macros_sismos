#!/usr/bin/env python3
"""
Procesador Sísmico - Análisis de Daño Estructural
Conversión completa de macros VBA a Python

Basado en:
- Jiang, H.J., Chen, L.Z. & Chen, Q. (2011) para evaluación de daño
- Macros VBA originales: Hinges_List, Moment_Rotation, Damage_Index

Uso:
    python main.py --hp 3.0 --direccion X
"""

import argparse
import logging
import os
from datetime import datetime
from helpers.processor_helper import cargar_archivos
from macros.hinges_list import macro_hinges_list
from macros.moment_rotation import macro_moment_rotation
from macros.damage_index import macro_damage_index
from config import (
    HP, DIRECCION, LOG_LEVEL, FORMATO_SALIDA, PRECISION_DECIMAL,
    mostrar_configuracion
)

def configurar_logging():
    """Configura el sistema de logging usando configuración global"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = f"analisis_sismico_{timestamp}.log"
    
    # Usar nivel de logging de configuración global
    log_level = getattr(logging, LOG_LEVEL.upper(), logging.INFO)
    
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding=FORMATO_SALIDA['encoding']),
            logging.StreamHandler()
        ]
    )
    return log_file

def crear_directorio_resultados():
    """Crea directorio para resultados con timestamp"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    directorio = f"resultados_{timestamp}"
    os.makedirs(directorio, exist_ok=True)
    return directorio

def guardar_csv_formato_europeo(df, archivo, columnas_numericas=None):
    """Guarda DataFrame usando configuración global de formato"""
    df_export = df.copy()
    
    if columnas_numericas:
        for col in columnas_numericas:
            if col in df_export.columns:
                df_export[col] = df_export[col].astype(str).str.replace('.', FORMATO_SALIDA['decimal'])
    
    df_export.to_csv(archivo, 
                     sep=FORMATO_SALIDA['separador'], 
                     index=False, 
                     encoding=FORMATO_SALIDA['encoding'],
                     float_format=f'%.{PRECISION_DECIMAL}f')

def main():
    """Función principal del procesador sísmico"""
    parser = argparse.ArgumentParser(description='Procesador Sísmico - Análisis de Daño Estructural')
    parser.add_argument('--hp', type=float, help=f'Altura de piso en metros (default desde config.py: {HP})')
    parser.add_argument('--direccion', choices=['X', 'Y'], help=f'Dirección del sismo (default desde config.py: {DIRECCION})')
    
    args = parser.parse_args()
    
    # Usar configuración global si no se proporcionan argumentos
    hp_actual = args.hp if args.hp is not None else HP
    direccion_actual = args.direccion if args.direccion is not None else DIRECCION
    
    # Configurar logging
    log_file = configurar_logging()
    logger = logging.getLogger(__name__)
    
    logger.info("="*60)
    logger.info("🏗️  PROCESADOR SÍSMICO - ANÁLISIS DE DAÑO ESTRUCTURAL")
    logger.info("="*60)
    logger.info(f"📋 Parámetros: Hp={hp_actual}m, Dirección={direccion_actual}")
    logger.info(f"📝 Log guardado en: {log_file}")
    
    # Mostrar configuración
    mostrar_configuracion()
    
    try:
        # Crear directorio de resultados
        dir_resultados = crear_directorio_resultados()
        logger.info(f"📁 Directorio de resultados: {dir_resultados}")
        
        # PASO 1: Cargar archivos de entrada
        logger.info("\n🔄 PASO 1: Cargando archivos de entrada...")
        archivos = cargar_archivos()
        
        # PASO 2: Ejecutar MACRO 1 - Hinges_List
        logger.info("\n🔄 PASO 2: Ejecutando Hinges_List...")
        df_rt = macro_hinges_list(archivos, hp=hp_actual, direccion=direccion_actual)
        
        # Guardar RT.csv
        archivo_rt = os.path.join(dir_resultados, 'RT.csv')
        columnas_numericas_rt = ['RelDist', 'L', 'Storey', 'B', 'H', "f'c", 'fy', 'ρsx', 'α', 
                                'P average', 'My', 'Mu', 'Cy', 'Cu', 'Lc*', 'Rp']
        guardar_csv_formato_europeo(df_rt, archivo_rt, columnas_numericas_rt)
        logger.info(f"💾 RT.csv guardado: {archivo_rt}")
        
        # PASO 3: Ejecutar MACRO 2 - Moment_Rotation
        logger.info("\n🔄 PASO 3: Ejecutando Moment_Rotation...")
        mr_matrix_data, df_rt_updated = macro_moment_rotation(archivos, df_rt, direccion=direccion_actual)
        
        # Guardar MR.csv (formato matricial)
        archivo_mr = os.path.join(dir_resultados, 'MR.csv')
        crear_mr_matricial(mr_matrix_data, df_rt['Hinge'].tolist(), archivo_mr)
        logger.info(f"💾 MR.csv guardado: {archivo_mr}")
        
        # Actualizar RT.csv con rotaciones calculadas
        columnas_numericas_rt_updated = columnas_numericas_rt + ['θy', 'θp', 'θu', 'θc', 'θm', 'Pm', 'EH']
        guardar_csv_formato_europeo(df_rt_updated, archivo_rt, columnas_numericas_rt_updated)
        
        # PASO 4: Ejecutar MACRO 3 - Damage_Index
        logger.info("\n🔄 PASO 4: Ejecutando Damage_Index...")
        df_id, df_rt_final = macro_damage_index(df_rt_updated)
        
        # Guardar ID.csv
        archivo_id = os.path.join(dir_resultados, 'ID.csv')
        columnas_numericas_id = ['EH', 'ID']
        guardar_csv_formato_europeo(df_id, archivo_id, columnas_numericas_id)
        logger.info(f"💾 ID.csv guardado: {archivo_id}")
        
        # Guardar RT.csv final con todos los cálculos
        columnas_numericas_rt_final = columnas_numericas_rt_updated + ['no', 'Beta', 'ID']
        guardar_csv_formato_europeo(df_rt_final, archivo_rt, columnas_numericas_rt_final)
        
        # PASO 5: Resumen de resultados
        logger.info("\n📊 RESUMEN DE RESULTADOS:")
        logger.info(f"   🔗 Rótulas procesadas: {len(df_rt_final)}")
        
        # Distribución de niveles de desempeño
        distribucion = df_id['DS'].value_counts()
        logger.info("   📈 Distribución de niveles de desempeño:")
        niveles = {
            'TO': 'Totalmente operativo',
            'IO': 'Ocupación inmediata',
            'LS': 'Seguridad de la vida', 
            'CP': 'Prevención de colapso',
            'CL': 'Colapso'
        }
        for nivel, cantidad in distribucion.items():
            logger.info(f"      {nivel} ({niveles.get(nivel, nivel)}): {cantidad} rótulas")
        
        # Estadísticas de daño
        id_stats = df_id['ID'].astype(str).str.replace(',', '.').astype(float)
        logger.info(f"   📊 Índice de daño promedio: {id_stats.mean():.4f}")
        logger.info(f"   📊 Índice de daño máximo: {id_stats.max():.4f}")
        
        logger.info("\n✅ ANÁLISIS COMPLETADO EXITOSAMENTE")
        logger.info(f"📁 Resultados disponibles en: {dir_resultados}/")
        logger.info("="*60)
        
    except Exception as e:
        logger.error(f"❌ Error durante el análisis: {str(e)}")
        logger.exception("Detalles del error:")
        return 1
    
    return 0

def crear_mr_matricial(mr_matrix_data, rotulas_ordenadas, archivo_salida):
    """Crea archivo MR.csv en formato matricial"""
    
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
                
                # Formatear números con coma decimal
                moment_str = f"{moment:.6f}".replace('.', ',')
                rotation_str = f"{rotation:.6f}".replace('.', ',')
                axial_str = f"{axial:.2f}".replace('.', ',')
                
                row.extend([moment_str, rotation_str, axial_str])
            else:
                row.extend(['', '', ''])
        data_rows.append(row)
    
    # Escribir archivo CSV
    with open(archivo_salida, 'w', encoding='utf-8') as f:
        f.write(';'.join(header_row1) + '\n')
        f.write(';'.join(header_row2) + '\n')
        f.write(';'.join(header_row3) + '\n')
        
        for row in data_rows:
            f.write(';'.join(row) + '\n')

if __name__ == "__main__":
    exit(main())
