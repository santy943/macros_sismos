Sub Damage_Index()

'Seismic Damage Assessment and Perforance Levels of Reinforced Concrete members
'H.J. Jiang, L. Z. Chen and Q. Chen, 2011

'=== PYTHON EQUIVALENTE: macro_damage_index(df_rt) ===
'Esta macro se replica en procesador_sismico_limpio.py líneas 512-586

'PASO 1: Limpiar información existente (Python: inicializar DataFrames)
'Python: df_rt['no'] = 0.0, df_rt['Beta'] = 0.0, df_rt['ID'] = 0.0, df_rt['ND'] = ''
ThisWorkbook.Sheets("RT").Range("AA8:AD1048576").Value = Empty
ThisWorkbook.Sheets("ID").Range("B8:H1048576").Value = Empty

'PASO 2: Bucle principal para calcular índice de daño por rótula
'Python: for idx, rotula in df_rt.iterrows():
'Declara variables
i = 8
Do While ThisWorkbook.Sheets("RT").Cells(i, 2).Value <> Empty
'PASO 2.1: Leer parámetros desde hoja RT (Python: obtener valores desde DataFrame)
'Python: b = convertir_europeo_a_float(rotula['B'])
b = ThisWorkbook.Sheets("RT").Cells(i, 8).Value
h = ThisWorkbook.Sheets("RT").Cells(i, 9).Value
fc = ThisWorkbook.Sheets("RT").Cells(i, 10).Value
fy = ThisWorkbook.Sheets("RT").Cells(i, 11).Value
psx = ThisWorkbook.Sheets("RT").Cells(i, 12).Value
a = ThisWorkbook.Sheets("RT").Cells(i, 13).Value
My = ThisWorkbook.Sheets("RT").Cells(i, 15).Value
Lc = ThisWorkbook.Sheets("RT").Cells(i, 19).Value
Ry = ThisWorkbook.Sheets("RT").Cells(i, 20).Value
Ru = ThisWorkbook.Sheets("RT").Cells(i, 22).Value
Rc = ThisWorkbook.Sheets("RT").Cells(i, 23).Value
Rm = ThisWorkbook.Sheets("RT").Cells(i, 24).Value
P = ThisWorkbook.Sheets("RT").Cells(i, 25).Value

EH = ThisWorkbook.Sheets("RT").Cells(i, 26).Value

'PASO 2.2: Evaluación de parámetros según Jiang, Chen & Chen (2011)
'Python: no = p / (b * h * fc * 1000) if b > 0 and h > 0 and fc > 0 else 0
no = P / (b * h * fc * 10 ^ 3)
'Python: beta = (0.023*lc/h + 3.352*no^2.35) * 0.818^(a*psx*fy/fc*100) + 0.039
Beta = (0.023 * Lc / h + 3.352 * no ^ 2.35) * 0.818 ^ (a * psx * fy / fc * 100) + 0.039

'PASO 2.3: Calcular rotación máxima de demanda
'Python: dm = max(rm, rc)
If Rm > Rc Then
dM = Rm
Else
dM = Rc
End If

'PASO 2.4: Evaluación del índice de daño (fórmula principal)
'Python: id_value = (1-beta)*(dm-rc)/(ru-rc) + beta*eh/(my*(ru-ry))
ID = (1 - Beta) * (dM - Rc) / (Ru - Rc) + Beta * EH / (My * (Ru - Ry))

'PASO 2.5: Clasificación de nivel de desempeño según índice de daño
'Python: if id_value < 0.05: nd = "TO" elif id_value < 0.15: nd = "IO" ...
If ID < 0.05 Then
ND = "TO" 'Totalmente operativo
Else

    If ID < 0.15 Then
    ND = "IO" 'Ocupación inmediata
    Else

        If ID < 0.45 Then
        ND = "LS" 'Seguridad de la vida
        Else

            If ID < 1 Then
            ND = "CP" 'Prevención de colapso
            Else
            ND = "CL" 'Colapso
            End If
        End If
    End If
End If

'PASO 2.6: Guardar resultados calculados en hoja RT
'Python: df_rt.at[idx, 'no'] = no; df_rt.at[idx, 'Beta'] = beta; etc.
ThisWorkbook.Sheets("RT").Cells(i, 27).Value = no
ThisWorkbook.Sheets("RT").Cells(i, 28).Value = Beta
ThisWorkbook.Sheets("RT").Cells(i, 29).Value = ID
ThisWorkbook.Sheets("RT").Cells(i, 30).Value = ND

'Salta a siguiente fila en RT
i = i + 1
Loop

'PASO 3: Crear resumen en hoja ID (Python: crear DataFrame df_id)
'Python: df_id = df_rt[['Hinge', 'Section', 'Frame', 'Storey', 'EH']].copy()
i = 8

Do While ThisWorkbook.Sheets("RT").Cells(i, 2).Value <> 0

'PASO 3.1: Leer datos calculados desde RT
'Python: hinge = rotula['Hinge']; section = rotula['Section']; etc.
Hinge = ThisWorkbook.Sheets("RT").Cells(i, 2).Value
Section = ThisWorkbook.Sheets("RT").Cells(i, 3).Value
Frame = ThisWorkbook.Sheets("RT").Cells(i, 4).Value
Storey = ThisWorkbook.Sheets("RT").Cells(i, 7).Value
EH = ThisWorkbook.Sheets("RT").Cells(i, 26).Value
ID = ThisWorkbook.Sheets("RT").Cells(i, 29).Value
DS = ThisWorkbook.Sheets("RT").Cells(i, 30).Value

'PASO 3.2: Copiar a hoja ID para resumen
'Python: df_id.at[idx, 'ID'] = id_value; df_id.at[idx, 'DS'] = ds
ThisWorkbook.Sheets("ID").Cells(i, 2).Value = Hinge
ThisWorkbook.Sheets("ID").Cells(i, 3).Value = Section
ThisWorkbook.Sheets("ID").Cells(i, 4).Value = Frame
ThisWorkbook.Sheets("ID").Cells(i, 5).Value = Storey
ThisWorkbook.Sheets("ID").Cells(i, 6).Value = EH
ThisWorkbook.Sheets("ID").Cells(i, 7).Value = ID
ThisWorkbook.Sheets("ID").Cells(i, 8).Value = DS

i = i + 1

Loop

End Sub