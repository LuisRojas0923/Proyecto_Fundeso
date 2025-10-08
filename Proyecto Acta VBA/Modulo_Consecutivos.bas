Attribute VB_Name = "Modulo_Consecutivos"

' === MÓDULO PARA GESTIÓN DE CONSECUTIVOS ===

' Obtener consecutivo de actividad - LEE DESDE LA HOJA EXPORTADA
Public Function ObtenerConsecutivoActividad(frm As Object, area As String, capitulo As String) As Long
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim maxConsecutivo As Long
    Dim ws As Worksheet
    
    maxConsecutivo = 0
    
    ' PRIMERO: Buscar en la hoja exportada (fuente de verdad)
    Set ws = ObtenerHojaDestino()
    If Not ws Is Nothing Then
        ' Buscar desde la fila 2 (después de encabezados)
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ' Verificar si coincide el área y capítulo
            If ws.Cells(i, 2).Value = area And ws.Cells(i, 4).Value = capitulo Then
                If IsNumeric(ws.Cells(i, 5).Value) Then ' Columna 5: CONSECUTIVO ACTIVIDAD
                    If CDbl(ws.Cells(i, 5).Value) > maxConsecutivo Then
                        maxConsecutivo = CDbl(ws.Cells(i, 5).Value)
                    End If
                End If
            End If
        Next i
    End If
    
    ' SEGUNDO: Buscar en el ListBox de Trabajo (datos pendientes)
    With frm.Listbox_Trabajo
        For i = 0 To .ListCount - 1
            If .List(i, 1) = area And .List(i, 3) = capitulo Then ' Área y Capítulo coinciden
                If IsNumeric(.List(i, 4)) Then ' Columna 5: CONSECUTIVO ACTIVIDAD
                    If CDbl(.List(i, 4)) > maxConsecutivo Then
                        maxConsecutivo = CDbl(.List(i, 4))
                    End If
                End If
            End If
        Next i
    End With
    
    ObtenerConsecutivoActividad = maxConsecutivo + 1
    Debug.Print "Consecutivo Actividad para Área '" & area & "' y Capítulo '" & capitulo & "': " & ObtenerConsecutivoActividad
    Exit Function
ErrHandler:
    Debug.Print "Error en ObtenerConsecutivoActividad: " & Err.Description
    ObtenerConsecutivoActividad = 1
End Function

' Validar consecutivos duplicados
Public Function ValidarConsecutivosDuplicados(frm As Object, area As String, capitulo As String, consecutivoCapitulo As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim ws As Worksheet
    
    ' PRIMERO: Verificar en la hoja exportada
    Set ws = ObtenerHojaDestino()
    If Not ws Is Nothing Then
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If ws.Cells(i, 2).Value = area And ws.Cells(i, 4).Value = capitulo Then
                If IsNumeric(ws.Cells(i, 3).Value) Then
                    If CDbl(ws.Cells(i, 3).Value) = consecutivoCapitulo Then
                        Debug.Print "ADVERTENCIA: Consecutivo de capítulo duplicado encontrado en hoja exportada"
                        ValidarConsecutivosDuplicados = False
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    ' SEGUNDO: Verificar en el ListBox de Trabajo
    With frm.Listbox_Trabajo
        For i = 0 To .ListCount - 1
            If .List(i, 1) = area And .List(i, 3) = capitulo Then
                If IsNumeric(.List(i, 2)) Then
                    If CDbl(.List(i, 2)) = consecutivoCapitulo Then
                        Debug.Print "ADVERTENCIA: Consecutivo de capítulo duplicado encontrado en ListBox de Trabajo"
                        ValidarConsecutivosDuplicados = False
                        Exit Function
                    End If
                End If
            End If
        Next i
    End With
    
    ValidarConsecutivosDuplicados = True
    Exit Function
ErrHandler:
    Debug.Print "Error en ValidarConsecutivosDuplicados: " & Err.Description
    ValidarConsecutivosDuplicados = True ' En caso de error, permitir continuar
End Function

' Actualizar consecutivos en el sistema
Public Sub ActualizarConsecutivos(area As String, consecutivoCapitulo As Long, consecutivoActividad As Long)
    ' Esta función registra los consecutivos asignados para auditoría
    Debug.Print "=== CONSECUTIVOS ASIGNADOS ==="
    Debug.Print "Área: " & area
    Debug.Print "Consecutivo Capítulo: " & consecutivoCapitulo
    Debug.Print "Consecutivo Actividad: " & consecutivoActividad
    Debug.Print "============================="
End Sub

' Mostrar resumen de consecutivos por área y capítulo
Public Sub MostrarResumenConsecutivos()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim i As Long
    Dim resumen As String
    Dim areasUnicas As String
    Dim capitulosUnicos As String
    
    resumen = "=== RESUMEN DE CONSECUTIVOS EXPORTADOS ===" & vbCrLf & vbCrLf
    
    Set ws = ObtenerHojaDestino()
    If ws Is Nothing Then
        resumen = resumen & "No hay datos exportados aún." & vbCrLf
    Else
        ' Contar desde la fila 2 (después de encabezados)
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If ws.Cells(i, 1).Value <> "" And ws.Cells(i, 2).Value <> "" Then
                Dim area As String, capitulo As String
                Dim consecCapitulo As String, consecActividad As String
                
                area = ws.Cells(i, 2).Value ' Columna 2: AREA
                capitulo = ws.Cells(i, 4).Value ' Columna 4: DESCRIPCION CAPITULO
                consecCapitulo = ws.Cells(i, 3).Value ' Columna 3: CONSECUTIVO CAPITULO
                consecActividad = ws.Cells(i, 5).Value ' Columna 5: CONSECUTIVO ACTIVIDAD
                
                resumen = resumen & "Área: " & area & " | Capítulo: " & capitulo & vbCrLf
                resumen = resumen & "  Consecutivo Capítulo: " & consecCapitulo & " | Consecutivo Actividad: " & consecActividad & vbCrLf & vbCrLf
            End If
        Next i
    End If
    
    Debug.Print resumen
    Exit Sub
ErrHandler:
    Debug.Print "Error en MostrarResumenConsecutivos: " & Err.Description
End Sub

' Función auxiliar para obtener hoja destino (necesaria para consecutivos)
Private Function ObtenerHojaDestino(Optional crearSiNoExiste As Boolean = False) As Worksheet
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    ' Buscar la hoja "Acta-Presupuesto"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Acta-Presupuesto" Then
            Set ObtenerHojaDestino = ws
            Exit Function
        End If
    Next ws
    
    ' Si no existe y se solicita crear
    If crearSiNoExiste Then
        If MsgBox("La hoja 'Acta-Presupuesto' no existe. ¿Desea crearla?", vbQuestion + vbYesNo, "Hoja No Encontrada") = vbYes Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = "Acta-Presupuesto"
            
            ' Crear encabezados
            ws.Cells(1, 1).Value = "CONSECUTIVO AREA"
            ws.Cells(1, 2).Value = "AREA"
            ws.Cells(1, 3).Value = "CONSECUTIVO CAPITULO"
            ws.Cells(1, 4).Value = "DESCRIPCION CAPITULO"
            ws.Cells(1, 5).Value = "CONSECUTIVO ACTIVIDAD"
            ws.Cells(1, 6).Value = "CODIGO ACTIVIDAD"
            ws.Cells(1, 7).Value = "ACTIVIDAD"
            ws.Cells(1, 8).Value = "UND"
            ws.Cells(1, 9).Value = "CANTIDAD"
            ws.Cells(1, 10).Value = "VR. UNITARIO"
            ws.Cells(1, 11).Value = "VR. PARCIAL"
            
            Set ObtenerHojaDestino = ws
        Else
            Set ObtenerHojaDestino = Nothing
        End If
    Else
        ' Solo consultar, no crear
        Set ObtenerHojaDestino = Nothing
    End If
    
    Exit Function
ErrHandler:
    Debug.Print "Error en ObtenerHojaDestino: " & Err.Description
    Set ObtenerHojaDestino = Nothing
End Function
