Sub Hinges_List()

'=== PYTHON EQUIVALENTE: macro_hinges_list(archivos, hp, direccion) ===
'Esta macro se replica en procesador_sismico_limpio.py líneas 38-277

Salida = 0

'PASO 1: Limpiar datos previos (Python: inicializar DataFrames)
'Python: rotulas_unicas = []
ThisWorkbook.Sheets("RT").Range("B8:S1048576").Value = Empty
ThisWorkbook.Sheets("RT").Range("U8:U1048576").Value = Empty

'PASO 2: Calcular delta para saltos en datos (algoritmo VBA crítico)
'Python: dt = 4; while (dt < len(df_cr)-2 and convertir_europeo_a_float(df_cr.iloc[dt-1]['StepNum']) < ...)
'Delta para saltos
dt = 4
Do While ThisWorkbook.Sheets("CR").Cells(dt, 5).Value < ThisWorkbook.Sheets("CR").Cells(dt + 2, 5).Value
dt = dt + 2
Loop
dt = dt - 2

'PASO 3: Identificar rótulas únicas usando algoritmo delta (líneas 23-53 VBA)
'Python: while i < len(df_cr) and not pd.isna(df_cr.iloc[i]['Frame']):
'NR: Cuantificador del número de rotulas
NR = 0
d = 8
i = 4
Do While ThisWorkbook.Sheets("CR").Cells(i, 1).Value <> Empty
'PASO 3.1: Verificar si la rótula ya existe
'Python: ya_existe = any(r['Hinge'] == rotula_actual for r in rotulas_unicas)
Bandera = 1
j = 8

Do While Bandera = 1
If j > NR + 8 Then
Bandera = 0
Else
If ThisWorkbook.Sheets("CR").Cells(i, 7).Value = ThisWorkbook.Sheets("RT").Cells(j, 2).Value Then
Bandera = 0
Else
j = j + 1
End If
End If
Loop

'PASO 3.2: Si es nueva, agregar ambas rótulas del par (H1 y H2)
'Python: rotula1 = {'Hinge': df_cr.iloc[i]['GenHinge'], 'Section': df_cr.iloc[i]['AssignHinge'], ...}
If j > NR + 8 Then
Section = ThisWorkbook.Sheets("CR").Cells(i, 6).Value
RName = ThisWorkbook.Sheets("CR").Cells(i, 7).Value
Frame = ThisWorkbook.Sheets("CR").Cells(i, 1).Value
RelDist = ThisWorkbook.Sheets("CR").Cells(i, 8).Value
ThisWorkbook.Sheets("RT").Cells(d, 2).Value = RName
ThisWorkbook.Sheets("RT").Cells(d, 3).Value = Section
ThisWorkbook.Sheets("RT").Cells(d, 4).Value = Frame
ThisWorkbook.Sheets("RT").Cells(d, 5).Value = RelDist
d = d + 1
NR = NR + 1
'Segunda rótula del par (i+1)
Section = ThisWorkbook.Sheets("CR").Cells(i + 1, 6).Value
RName = ThisWorkbook.Sheets("CR").Cells(i + 1, 7).Value
Frame = ThisWorkbook.Sheets("CR").Cells(i + 1, 1).Value
RelDist = ThisWorkbook.Sheets("CR").Cells(i + 1, 8).Value
ThisWorkbook.Sheets("RT").Cells(d, 2).Value = RName
ThisWorkbook.Sheets("RT").Cells(d, 3).Value = Section
ThisWorkbook.Sheets("RT").Cells(d, 4).Value = Frame
ThisWorkbook.Sheets("RT").Cells(d, 5).Value = RelDist
d = d + 1
NR = NR + 1

Else
End If
i = i + dt 'Saltar usando delta

Loop

'PASO 4: Agregar longitudes y coordenadas desde CD (líneas 60-85 VBA)
'Python: for idx, row in df_rt.iterrows(): frame_data = df_cd[df_cd['Frame'] == frame_num]

'PASO 4.1: Solicitar altura de piso (Python: parámetro hp)
'Python: hp=3.0 (parámetro de función)
Hp = (InputBox("Ingresar Altura de piso Hp [m]: ", "Dato de entrada", 3))

'PASO 4.2: Buscar frame en CD y asignar longitud y piso
'Python: length = convertir_europeo_a_float(frame_info['Length']); storey = (centroid_y + hp/2) / hp
i = 4
j = 8
'Entrada = 0

Do While ThisWorkbook.Sheets("RT").Cells(j, 2).Value <> Empty
    If ThisWorkbook.Sheets("RT").Cells(j, 4).Value = ThisWorkbook.Sheets("CD").Cells(i, 1).Value Then
    ThisWorkbook.Sheets("RT").Cells(j, 6).Value = ThisWorkbook.Sheets("CD").Cells(i, 5).Value 'Indica Longitud del frame

    'Código para ubicar frame en altura (Floor)
    'Cálculo de piso basado en coordenada Y del centroide
    'If Entrada = 0 Then
    'y = -Hp / 2
    'Entrada = 1
    'Else
    'y = Hp / 2
    'Entrada = 0
    'End If

    'ThisWorkbook.Sheets("RT").Cells(j, 7).Value = (ThisWorkbook.Sheets("CD").Cells(i, 8).Value + y) / 3 'Ubica el frame Storey
    ThisWorkbook.Sheets("RT").Cells(j, 7).Value = (ThisWorkbook.Sheets("CD").Cells(i, 8).Value + Hp / 2) / Hp 'Ubica el frame Storey

    i = 4
    j = j + 1
    Else
    i = i + 1
End If
Loop

'PASO 5: Identificar axial promedio P(t=0) (líneas 86-97 VBA)
'Python: p_values = df_cr[(df_cr['GenHinge'] == rotula_nombre) & (df_cr['StepNum'] == 0)]['P']
i = 4
j = 8
Do While ThisWorkbook.Sheets("CR").Cells(i, 7).Value <> Empty
    'Buscar valores donde StepNum = 0 (tiempo inicial)
    If ThisWorkbook.Sheets("CR").Cells(i, 5).Value = 0 Then
    ThisWorkbook.Sheets("RT").Cells(j, 14).Value = ThisWorkbook.Sheets("CR").Cells(i, 10).Value * -1
    j = j + 1
    Else
    End If
i = i + 1
Loop

'PASO 6: Determinar parámetros de dirección (líneas 98-108 VBA)
'Python: if direccion.upper() == 'X': t, f, g, y = 4, 7, 4, 3 else: t, f, g, y = 0, 8, 3, 4
'Preguntar si es Sismo X o Sismo Y
Action = MsgBox("¿Sismo Dirección X?", vbYesNo + vbExclamation, Rcaso & " Dirección del Evento")
Select Case Action

    Case vbYes:
    'Offsets para dirección X (columnas en SC para My(90), Mu(90), etc.)
    t = 4
    f = 7
    g = 4
    y = 3
    ThisWorkbook.Sheets("RT").Cells(3, 12).Value = 1 'Valor 1 indica Sismo X
    
    Case vbNo:
    'Offsets para dirección Y (columnas en SC para My(00), Mu(00), etc.)
    t = 0
    f = 8
    g = 3
    y = 4
    ThisWorkbook.Sheets("RT").Cells(3, 12).Value = 2 'Valor 2 indica Sismo Y
End Select

'PASO 7: Interpolación My, Mu, Cy, Cu (líneas 110-185 VBA)
'Python: for i in range(len(df_rt)): section_name = df_rt.iloc[i]['Section']; axial = abs(df_rt.iloc[i]['P average'])
For i = 1 To NR
Section = ThisWorkbook.Sheets("RT").Cells(i + 7, 3).Value
Axial = ThisWorkbook.Sheets("RT").Cells(i + 7, 14).Value
Lc = ThisWorkbook.Sheets("RT").Cells(i + 7, 6).Value 'Adoptado como la longitud del elemento

'PASO 7.1: Buscar sección en SC
'Python: section_data = df_sc[df_sc['Section'] == section_name]
Bandera = 1
j = 8 'Fila inicio en "SC"
Do While Bandera = 1

    If ThisWorkbook.Sheets("SC").Cells(j, 2).Value = Empty Then
    Mensaje = "No se encontró Sección: " & Section & "; Acción No Completada"
    MsgBox Mensaje
    Bandera = 0
    j = j + 1
    Exit Sub
    Else
    End If

'PASO 7.2: Obtener propiedades básicas de la sección
'Python: b = convertir_europeo_a_float(sec['B']); h = convertir_europeo_a_float(sec['H']); etc.
If ThisWorkbook.Sheets("SC").Cells(j, 2).Value = Section Then

b = ThisWorkbook.Sheets("SC").Cells(j, g).Value
h = ThisWorkbook.Sheets("SC").Cells(j, y).Value
fc = ThisWorkbook.Sheets("SC").Cells(j, 5).Value
fy = ThisWorkbook.Sheets("SC").Cells(j, 6).Value
psx = ThisWorkbook.Sheets("SC").Cells(j, f).Value
alpha = ThisWorkbook.Sheets("SC").Cells(j, 9).Value

'PASO 7.3: Buscar rango de interpolación correcto
'Python: for p_idx in range(1, 5): p_col = f'P({p_idx})'; p_val = convertir_europeo_a_float(sec[p_col])
Bandera = 0
k = 1

'Encontrar puntos de interpolación donde Pi <= Axial <= Pf
Do While ThisWorkbook.Sheets("SC").Cells(j, 9 * k + 1).Value < Axial
Pi = ThisWorkbook.Sheets("SC").Cells(j, 9 * k + 1).Value
Pf = ThisWorkbook.Sheets("SC").Cells(j, 9 * (k + 1) + 1).Value
'Obtener valores My, Mu, Cy, Cu para interpolación (usando offsets t)
Myi = ThisWorkbook.Sheets("SC").Cells(j, 9 * k + t + 2).Value
Myf = ThisWorkbook.Sheets("SC").Cells(j, 9 * (k + 1) + t + 2).Value
Mui = ThisWorkbook.Sheets("SC").Cells(j, 9 * k + t + 3).Value
Muf = ThisWorkbook.Sheets("SC").Cells(j, 9 * (k + 1) + t + 3).Value
Cyi = ThisWorkbook.Sheets("SC").Cells(j, 9 * k + t + 4).Value
Cyf = ThisWorkbook.Sheets("SC").Cells(j, 9 * (k + 1) + t + 4).Value
Cui = ThisWorkbook.Sheets("SC").Cells(j, 9 * k + t + 5).Value
Cuf = ThisWorkbook.Sheets("SC").Cells(j, 9 * (k + 1) + t + 5).Value

k = k + 1
Loop

'PASO 7.4: Interpolación lineal exacta VBA
'Python: my = myi + (myf - myi) * (axial - pi) / (pf - pi)
'Interpolación My
m = (Myf - Myi) / (Pf - Pi)
bo = Myi - m * Pi
My = m * Axial + bo

'Interpolación Mu
m = (Muf - Mui) / (Pf - Pi)
bo = Mui - m * Pi
Mu = m * Axial + bo

'Interpolación Cy
m = (Cyf - Cyi) / (Pf - Pi)
bo = Cyi - m * Pi
Cy = m * Axial + bo

'Interpolación Cu
m = (Cuf - Cui) / (Pf - Pi)
bo = Cui - m * Pi
Cu = m * Axial + bo

'PASO 7.5: Calcular longitud plástica y rotación plástica
'Python: lp = 0.5 * h; rp = (cu - cy) * lp
Lp = 0.5 * h
Rp = (Cu - Cy) * Lp

'PASO 7.6: Asignar valores calculados a hoja RT
'Python: df_rt.at[i, 'My'] = my; df_rt.at[i, 'Mu'] = mu; etc.
ThisWorkbook.Sheets("RT").Cells(i + 7, 8).Value = b
ThisWorkbook.Sheets("RT").Cells(i + 7, 9).Value = h

ThisWorkbook.Sheets("RT").Cells(i + 7, 10).Value = fc
ThisWorkbook.Sheets("RT").Cells(i + 7, 11).Value = fy

ThisWorkbook.Sheets("RT").Cells(i + 7, 12).Value = psx
ThisWorkbook.Sheets("RT").Cells(i + 7, 13).Value = alpha

ThisWorkbook.Sheets("RT").Cells(i + 7, 15).Value = My
ThisWorkbook.Sheets("RT").Cells(i + 7, 16).Value = Mu
ThisWorkbook.Sheets("RT").Cells(i + 7, 17).Value = Cy
ThisWorkbook.Sheets("RT").Cells(i + 7, 18).Value = Cu

ThisWorkbook.Sheets("RT").Cells(i + 7, 19).Value = Lc

ThisWorkbook.Sheets("RT").Cells(i + 7, 21).Value = Rp

Else

End If
j = j + 1

Loop

Next i

Mensaje = "Se encontraron " & NR & " Rotulas"
MsgBox Mensaje

End Sub