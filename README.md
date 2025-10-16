# Procesador Sísmico - Análisis de Daño Estructural

Conversión completa de macros VBA a Python para análisis de daño sísmico estructural.

## Descripción

Este proyecto implementa las tres macros principales de Excel VBA:

- **Hinges_List**: Identificación y caracterización de rótulas plásticas
- **Moment_Rotation**: Análisis de curvas momento-rotación y energía histerética  
- **Damage_Index**: Evaluación de índices de daño según Jiang, Chen & Chen (2011)

## Archivos principales

- `main.py`: Ejecutor principal del análisis completo
- `procesador_sismico_limpio.py`: Funciones principales (réplicas exactas de macros VBA)
- `csv/`: Archivos de entrada (CD.csv, CR.csv, SC.csv, HK.csv)
- `requirements.txt`: Dependencias de Python

## Instalación

1. Crear entorno virtual:

```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
```

2. Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Uso

### Ejecución completa

```bash
python main.py --hp 3.0 --direccion X
```

### Parámetros

- `--hp`: Altura de piso en metros (default: 3.0)
- `--direccion`: Dirección del sismo X o Y (default: X)

### Ejemplo

```bash
python main.py --hp 3.5 --direccion Y
```

## Archivos de entrada requeridos

Ubicar en carpeta `csv/`:

- **CD.csv**: Conectividad de elementos (Frame, JointI, JointJ, Length, CentroidY, etc.)
- **CR.csv**: Estados de rótulas (Frame, GenHinge, StepNum, P, M2, M3, R2Plastic, R3Plastic, etc.)
- **SC.csv**: Propiedades de secciones (Section, B, H, f'c, fyw, ρsx, α, P(1-4), My, Mu, φy, φu)
- **HK.csv**: Rigideces momento-rotación (Hinge Name, Kx, Ky)

## Archivos de salida

El programa genera en `resultados_YYYYMMDD_HHMMSS/`:

- **RT.csv**: Resultados de rótulas con propiedades y rotaciones calculadas
- **MR.csv**: Curvas momento-rotación en formato matricial
- **ID.csv**: Índices de daño y niveles de desempeño

## Metodología

### Evaluación de daño (Jiang, Chen & Chen, 2011)

```text
ID = (1-β)*(dM-Rc)/(Ru-Rc) + β*EH/(My*(Ru-Ry))
```

Donde:

- **β**: Factor de degradación de rigidez
- **dM**: Rotación máxima de demanda = max(Rm, Rc)
- **EH**: Energía histerética disipada
- **Ru, Ry, Rc**: Rotaciones última, de fluencia y de fisuración

### Niveles de desempeño

- **TO** (ID < 0.05): Totalmente operativo
- **IO** (ID < 0.15): Ocupación inmediata
- **LS** (ID < 0.45): Seguridad de la vida
- **CP** (ID < 1.0): Prevención de colapso
- **CL** (ID ≥ 1.0): Colapso

## Estructura del proyecto

```text
procesador_csv/
├── main.py                     # Ejecutor principal
├── procesador_sismico_limpio.py # Funciones principales
├── csv/                        # Archivos de entrada
│   ├── CD.csv
│   ├── CR.csv  
│   ├── SC.csv
│   └── HK.csv
├── resultados_*/               # Resultados (generado automáticamente)
│   ├── RT.csv
│   ├── MR.csv
│   └── ID.csv
├── requirements.txt
└── README.md
```

## Logs

El programa genera logs detallados en `analisis_sismico_YYYYMMDD_HHMMSS.log` con información del proceso y estadísticas de resultados.
# macros_sismos
