Sub Moment_Rotation()

'=== PYTHON EQUIVALENTE: macro_moment_rotation(archivos, df_rt, direccion) ===
'Esta macro se replica en procesador_sismico_limpio.py líneas 279-510

'PASO 1: Limpiar información existente (Python: inicializar estructuras)
'Python: mr_matrix_data = {}; df_rt columnas θy, θp, θu, θc, θm, Pm, EH
ThisWorkbook.Sheets("MR").Range("A5:XFD1048576").Value = Empty
ThisWorkbook.Sheets("RT").Range("T8:T1048576").Value = Empty
ThisWorkbook.Sheets("RT").Range("V8:Z1048576").Value = Empty

'Cuenta el número de rótulas
i = 8
NoRotulas = 0
Do While ThisWorkbook.Sheets("RT").Cells(i, 2).Value <> Empty
NoRotulas = NoRotulas + 1
i = i + 1
Loop

'Lista en filas los nombres de rótulas
For i = 1 To NoRotulas
j = 3 * i - 1
k = i + 7
ThisWorkbook.Sheets("MR").Cells(5, j).Value = ThisWorkbook.Sheets("RT").Cells(k, 2).Value
ThisWorkbook.Sheets("MR").Cells(6, j) = "M"
ThisWorkbook.Sheets("MR").Cells(6, j + 1) = "Rot"
ThisWorkbook.Sheets("MR").Cells(6, j + 2) = "P"
ThisWorkbook.Sheets("MR").Cells(7, j) = "kN-m"
ThisWorkbook.Sheets("MR").Cells(7, j + 1) = "Rad"
ThisWorkbook.Sheets("MR").Cells(7, j + 2) = "kN"
Next i

'Identifica si es Sismo X o Sismo Y
Sismo = ThisWorkbook.Sheets("RT").Cells(3, 12).Value

If Sismo = 1 Then 'Identifica columna de rotaciones en función de Sismo X o Sismo Y
Cm = 15 'Columna para Momentos en CR
Crp = 21 'Columna para Rotaciones en CR
Ck = 3 'Columna Rigidez Kx en HK
Else
Cm = 14 'Columna para Momentos en CR
Crp = 20 'Columna para Rotaciones en CR
Ck = 4 'Columna Rigidez Ky en HK
End If

'Organiza en columnas la información por rotula: Momento, Rotación, Axial

i = 4
j = 8
k = 2

Do While ThisWorkbook.Sheets("CR").Cells(i, 1).Value <> Empty

If ThisWorkbook.Sheets("CR").Cells(i, 5).Value < ThisWorkbook.Sheets("CR").Cells(i + 2, 5).Value Then
ThisWorkbook.Sheets("MR").Cells(j, k).Value = ThisWorkbook.Sheets("CR").Cells(i, Cm).Value 'Momento Rotula Extremo Inicial H1
ThisWorkbook.Sheets("MR").Cells(j, k + 1).Value = ThisWorkbook.Sheets("CR").Cells(i, Crp).Value 'Rotación Plástica Rotula Extremo Inicial H1
ThisWorkbook.Sheets("MR").Cells(j, k + 2).Value = ThisWorkbook.Sheets("CR").Cells(i, 10).Value  'Axial Rotula Extemo Inicial H1

ThisWorkbook.Sheets("MR").Cells(j, k + 3).Value = ThisWorkbook.Sheets("CR").Cells(i + 1, Cm).Value 'Momento Rotula Extemo Final H2
ThisWorkbook.Sheets("MR").Cells(j, k + 4).Value = ThisWorkbook.Sheets("CR").Cells(i + 1, Crp).Value 'Rotación Plástica Rotula Extremo Inicial H2
ThisWorkbook.Sheets("MR").Cells(j, k + 5).Value = ThisWorkbook.Sheets("CR").Cells(i + 1, 10).Value 'Axial Rotula Extremo H2

j = j + 1

Else
ThisWorkbook.Sheets("MR").Cells(j, k).Value = ThisWorkbook.Sheets("CR").Cells(i, Cm).Value 'Momento Rotula Extremo Inicial H1
ThisWorkbook.Sheets("MR").Cells(j, k + 1).Value = ThisWorkbook.Sheets("CR").Cells(i, Crp).Value 'Rotación Plástica Rotula Extremo Inicial H1
ThisWorkbook.Sheets("MR").Cells(j, k + 2).Value = ThisWorkbook.Sheets("CR").Cells(i, 10).Value  'Axial Rotula Extemo Inicial H1

ThisWorkbook.Sheets("MR").Cells(j, k + 3).Value = ThisWorkbook.Sheets("CR").Cells(i + 1, Cm).Value 'Momento Rotula Extemo Final H2
ThisWorkbook.Sheets("MR").Cells(j, k + 4).Value = ThisWorkbook.Sheets("CR").Cells(i + 1, Crp).Value 'Rotación Plástica Rotula Extremo Inicial H2
ThisWorkbook.Sheets("MR").Cells(j, k + 5).Value = ThisWorkbook.Sheets("CR").Cells(i + 1, 10).Value 'Axial Rotula Extremo H2

j = 8
k = k + 6

End If

i = i + 2
Loop

'Inicio código de rotaciones..................

i = 8
x = 1 'Contabiliza las entradas en código de rotaciones

Do While ThisWorkbook.Sheets("RT").Cells(i, 2).Value <> Empty

Hinge = ThisWorkbook.Sheets("RT").Cells(i, 2).Value
n = 3 * x 'Columna rotaciones
m = n - 1 'Columna de momentos

Bandera = 0 'Bandera para encontrar Rigidez Momento Rotación
j = 8
Do While Bandera = 0

If ThisWorkbook.Sheets("HK").Cells(j, 2).Value = Hinge Then
Bandera = 1
Kmr = ThisWorkbook.Sheets("HK").Cells(j, Ck).Value
Else
Bandera = 0
End If

If ThisWorkbook.Sheets("HK").Cells(j, 2).Value = Empty Then
Mensaje = "No se encontró Rotula : " & Hinge & " En Hoja HK"
MsgBox Mensaje

Mensaje = "Momentos Rotaciones No Completado"
MsgBox Mensaje

Exit Sub
End If

j = j + 1
Loop

'Evaluación de Rotación de fluencia Ry y Rotación última Ru

My = ThisWorkbook.Sheets("RT").Cells(i, 15).Value
Ry = My / Kmr

ThisWorkbook.Sheets("RT").Cells(i, 20).Value = Ry 'Rotación de fluencia

Rp = ThisWorkbook.Sheets("RT").Cells(i, 21).Value 'Rotación Plástica

Ru = Ry + Rp 'Rotación última

ThisWorkbook.Sheets("RT").Cells(i, 22).Value = Ru

'Evaluación de rotación de fisuración

fc = ThisWorkbook.Sheets("RT").Cells(i, 10).Value
b = ThisWorkbook.Sheets("RT").Cells(i, 8).Value
h = ThisWorkbook.Sheets("RT").Cells(i, 9).Value

fr = 0.7 * 1 * (fc) ^ 0.5

Mr = 1 / 6 * (fr * 10 ^ 3) * b * h ^ 2
Rc = Ry / My * Mr

ThisWorkbook.Sheets("RT").Cells(i, 23).Value = Rc 'Rotación de fisuración

'Código para encontrar rotación total

k = 8
Re = 0
dRe = 0

Do While ThisWorkbook.Sheets("MR").Cells(k, m).Value <> Empty

Mi = ThisWorkbook.Sheets("MR").Cells(k, m).Value
Rp = ThisWorkbook.Sheets("MR").Cells(k, n).Value

If k < 9 Then

Re = (Mi * 10 ^ 6) / Kmr
'Se multiplica por 10^6 porque los resultados "Re" son muy pequeños y VBA los hace igual a cero
'Posteriormente se escalan los resultados

Else

Mj = ThisWorkbook.Sheets("MR").Cells(k - 1, m).Value
dM = (Mi - Mj) * 10 ^ 6

If Mi > 0 Then 'Para Momento Positivo
    If Mi / My < 1 Then
    dRe = dM / Kmr
    Else
        If dM > 0 Then
        dRe = 0
        Else
        dRe = dM / Kmr
        End If
    End If

Else 'Para Momento Negativo
    If -1 * Mi / My < 1 Then
    dRe = dM / Kmr
    Else
        If dM < 0 Then
        dRe = 0
        Else
        dRe = dM / Kmr
        End If
    End If

End If

End If

Re = Re + dRe
Rt = Re + Rp * 10 ^ 6
ThisWorkbook.Sheets("MR").Cells(k, n).Value = Rt

k = k + 1
Loop

'Código para escalar los valores de rotaciones
k = 8
Do While ThisWorkbook.Sheets("MR").Cells(k, m).Value <> Empty
ThisWorkbook.Sheets("MR").Cells(k, n).Value = ThisWorkbook.Sheets("MR").Cells(k, n).Value * 10 ^ -6
k = k + 1
Loop
'Fin código para escalar

i = i + 1
x = x + 1
Loop 'Loop Rotaciones.....................

'Evaluación de energía EH, Máximos y Mínimos

For i = 1 To NoRotulas
EH = 0
dEh = 0
Rmax = 0
Rmin = 0
Pmin = 0
n = 3 * i 'Columna rotaciones
m = n - 1 'Columna de momentos
j = 9 'Inicia una fila despúes
Do While ThisWorkbook.Sheets("MR").Cells(j, m).Value <> Empty
dM = 0.5 * (ThisWorkbook.Sheets("MR").Cells(j, m).Value + ThisWorkbook.Sheets("MR").Cells(j - 1, m).Value)
DR = ThisWorkbook.Sheets("MR").Cells(j, n).Value - ThisWorkbook.Sheets("MR").Cells(j - 1, n).Value
dEh = dM * DR
EH = EH + dEh

'Encuentra valores máximos y mínimos
If ThisWorkbook.Sheets("MR").Cells(j, n).Value > Rmax Then
Rmax = ThisWorkbook.Sheets("MR").Cells(j, n).Value
End If

If ThisWorkbook.Sheets("MR").Cells(j, n).Value < Rmin Then
Rmin = ThisWorkbook.Sheets("MR").Cells(j, n).Value
End If

If ThisWorkbook.Sheets("MR").Cells(j, n + 1).Value < Pmin Then
Pmin = ThisWorkbook.Sheets("MR").Cells(j, n + 1).Value
End If

j = j + 1
Loop

If Rmin * -1 > Rmax Then
Rmax = Rmin * -1
Else
Rmax = Rmax
End If

ThisWorkbook.Sheets("RT").Cells(i + 7, 24).Value = Rmax
ThisWorkbook.Sheets("RT").Cells(i + 7, 25).Value = -1 * Pmin 'Compresión

If EH < 0 Then
If EH < -2 Then
Rcaso = ThisWorkbook.Sheets("MR").Cells(i + 7, 2).Value
Action = MsgBox("EH = " & EH & " ¿Ajustar EH=0?", vbYesNo + vbExclamation, Rcaso & " Negative Absorved Energy")
Select Case Action

    Case vbYes:
    EH = 0
    
    Case vbNo:
    EH = EH
End Select
Else
EH = 0
End If
Else
EH = EH
End If

Ry = ThisWorkbook.Sheets("RT").Cells(i + 7, 20).Value
If Rmax < Ry Then
EH = 0
Else
EH = EH
End If

ThisWorkbook.Sheets("RT").Cells(i + 7, 26).Value = EH

Next i

'Código diagramas de histéresis en Hoja "HY"
j = 23
For i = 1 To 250 'Limpia información existente

ThisWorkbook.Sheets("HY").Cells(j, 11).Value = Empty
ThisWorkbook.Sheets("HY").Cells(j + 1, 11).Value = Empty

ThisWorkbook.Sheets("HY").Cells(j, 27).Value = Empty
ThisWorkbook.Sheets("HY").Cells(j + 1, 27).Value = Empty

j = j + 19
Next i

j = 23
n = NoRotulas / 2
For i = 1 To n 'Asigna valores
H1 = ThisWorkbook.Sheets("RT").Cells(2 * i + 6, 2).Value
EH1 = ThisWorkbook.Sheets("RT").Cells(2 * i + 6, 26).Value
ThisWorkbook.Sheets("HY").Cells(j, 11).Value = H1
ThisWorkbook.Sheets("HY").Cells(j + 1, 11).Value = EH1

H2 = ThisWorkbook.Sheets("RT").Cells(2 * i + 7, 2).Value
EH2 = ThisWorkbook.Sheets("RT").Cells(2 * i + 7, 26).Value
ThisWorkbook.Sheets("HY").Cells(j, 27).Value = H2
ThisWorkbook.Sheets("HY").Cells(j + 1, 27).Value = EH2

j = j + 19
Next i

'Fin código

Menssage = "Action Moment - Rotation Completed"
MsgBox Menssage

End Sub