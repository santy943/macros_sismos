"""
Configuración Global del Procesador Sísmico
============================================

CAMBIA ESTOS VALORES AQUÍ PARA CONFIGURAR EL ANÁLISIS:
"""

# ========================================
# CONFIGURACIÓN PRINCIPAL - CAMBIAR AQUÍ
# ========================================

# Altura de piso en metros
HP = 3.0

# Dirección del sismo ('X' o 'Y')
DIRECCION = 'Y'

# ========================================
# CONFIGURACIÓN AVANZADA (OPCIONAL)
# ========================================

# Configuración de logging
LOG_LEVEL = 'INFO'  # DEBUG, INFO, WARNING, ERROR

# Configuración de archivos de entrada
ARCHIVOS_ENTRADA = {
    'CD': 'csv/CD.csv',
    'CR': 'csv/CR.csv', 
    'HK': 'csv/HK.csv',
    'SC': 'csv/SC.csv'
}

# Configuración de formato de salida
FORMATO_SALIDA = {
    'separador': ';',
    'decimal': ',',
    'encoding': 'utf-8'
}

# Configuración de cálculos
TOLERANCIA_NUMERICA = 1e-10
PRECISION_DECIMAL = 6

def validar_configuracion():
    """Valida que la configuración sea correcta"""
    errores = []
    
    if HP <= 0:
        errores.append(f"HP debe ser mayor que 0, valor actual: {HP}")
    
    if DIRECCION not in ['X', 'Y']:
        errores.append(f"DIRECCION debe ser 'X' o 'Y', valor actual: {DIRECCION}")
    
    if errores:
        raise ValueError("Errores de configuración:\n" + "\n".join(f"- {error}" for error in errores))
    
    return True

def mostrar_configuracion():
    """Muestra la configuración actual"""
    print("=" * 50)
    print("CONFIGURACIÓN ACTUAL:")
    print("=" * 50)
    print(f"Altura de piso (HP): {HP} m")
    print(f"Dirección del sismo: {DIRECCION}")
    print("=" * 50)

# Validar configuración al importar
validar_configuracion()
