Attribute VB_Name = "Modulo_ListBoxes"

' === MÓDULO PARA GESTIÓN DE LISTBOXES ===

' Configurar ListBox de Trabajo (Página 2) con 11 columnas
Public Sub ConfigurarListBoxTrabajo_Principal(frm As Object)
    On Error GoTo ErrHandler
    With frm.Listbox_Trabajo
        .Clear
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .ColumnCount = 11
        ' Anchos de columna optimizados. "Item" (col 1) reducido y espacio reasignado a Descripción (col 7).
        .ColumnWidths = "40;125;0;125;0;60;255;40;50;70;70"
        ' Ancho total del control para evitar barra horizontal.
        .Width = 850
    End With
    Debug.Print "ListBox de Trabajo configurado con 11 columnas"
    Exit Sub
ErrHandler:
    Debug.Print "Error en ConfigurarListBoxTrabajo: " & Err.Description
End Sub

' Configurar ListBox de Exportados (Página 3) con 11 columnas
Public Sub ConfigurarListBoxExportados(frm As Object)
    On Error GoTo ErrHandler
    With frm.Listbox_Exportados
        .Clear
        .ColumnCount = 11
        ' Mismos anchos que el ListBox de Trabajo para consistencia visual.
        .ColumnWidths = "40;125;0;125;0;60;255;40;50;70;70"
        ' Ancho total del control para evitar barra horizontal.
        .Width = 850
    End With
    Debug.Print "ListBox de Exportados configurado con 11 columnas"
    Exit Sub
ErrHandler:
    Debug.Print "Error en ConfigurarListBoxExportados: " & Err.Description
End Sub

' Ajustar el ancho total del ListBox de Registros
Public Sub AjustarAnchoListBox(frm As Object)
    ' Cambia el valor según el diseño deseado
    frm.Listbox_Registros.Width = 850 ' Ancho fijo en puntos, aumentado para mejor visibilidad
    Debug.Print "AjustarAnchoListBox: Ancho actual del ListBox_Registros = " & frm.Listbox_Registros.Width
End Sub

' Limpiar todos los ListBoxes
Public Sub LimpiarTodosLosListBoxes(frm As Object)
    On Error GoTo ErrHandler
    
    frm.Listbox_Registros.Clear
    frm.Listbox_Trabajo.Clear
    frm.Listbox_Exportados.Clear
    
    ' Reconfigurar ListBoxes
    Call ConfigurarListBoxTrabajo_Principal(frm)
    Call ConfigurarListBoxExportados(frm)
    
    Debug.Print "Todos los ListBoxes han sido limpiados"
    Exit Sub
ErrHandler:
    Debug.Print "Error en LimpiarTodosLosListBoxes: " & Err.Description
End Sub

' Filtrar y cargar ListBox de Registros
Public Sub FiltrarYCargarListBox(frm As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet, tbl As ListObject, i As Long, filaDestino As Long
    Dim datos()
    Dim filtroPalabra As String
    Dim nombreTabla As String
    Dim tablaEncontrada As Boolean
    
    nombreTabla = "ListaPrecios_PreciosClientes"
    tablaEncontrada = False
    
    ' Buscar la tabla en todas las hojas
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = nombreTabla Then
                tablaEncontrada = True
                Exit For
            End If
        Next tbl
        If tablaEncontrada Then Exit For
    Next ws
    
    If Not tablaEncontrada Then
        MsgBox "La tabla '" & nombreTabla & "' no existe en el libro.", vbCritical, "Error de configuración"
        Exit Sub
    End If
    
    ' Obtener filtro de palabra clave
    filtroPalabra = ""
    If Not frm.Controls("Palabra_Clave") Is Nothing Then 
        filtroPalabra = UCase(Trim(frm.Palabra_Clave.Value))
    End If
    
    ' Contar filas que cumplen el filtro
    filaDestino = 0
    With tbl.DataBodyRange
        For i = 1 To .Rows.Count
            If filtroPalabra = "" Or InStr(UCase(CStr(.Cells(i, 2).Value)), filtroPalabra) > 0 Then
                filaDestino = filaDestino + 1
            End If
        Next i
    End With
    
    If filaDestino = 0 Then
        frm.Listbox_Registros.Clear
        Exit Sub
    End If
    
    ' Preparar array de datos - ahora con 4 columnas
    ReDim datos(1 To filaDestino, 1 To 4)
    filaDestino = 1
    
    ' Cargar datos filtrados
    With tbl.DataBodyRange
        For i = 1 To .Rows.Count
            If filtroPalabra = "" Or InStr(UCase(CStr(.Cells(i, 2).Value)), filtroPalabra) > 0 Then
                ' Cargar las 4 columnas directamente
                datos(filaDestino, 1) = .Cells(i, 1).Value ' Código
                datos(filaDestino, 2) = .Cells(i, 2).Value ' Descripción
                datos(filaDestino, 3) = .Cells(i, 3).Value ' Unidad
                
                ' Formatear precio con $ si es numérico
                If IsNumeric(.Cells(i, 4).Value) Then
                    datos(filaDestino, 4) = Format(.Cells(i, 4).Value, "$#,##0.00")
                Else
                    datos(filaDestino, 4) = .Cells(i, 4).Value
                End If
                
                filaDestino = filaDestino + 1
            End If
        Next i
    End With
    
    ' Configurar y cargar ListBox - ahora con 4 columnas
    With frm.Listbox_Registros
        .Clear
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .ColumnCount = 4
        .ColumnWidths = "60;500;40;60" ' Ancho ajustado para la descripción
        .List = datos
        
        Debug.Print "ListBox_Registros configurado:"
        Debug.Print "  ColumnCount: " & .ColumnCount
        Debug.Print "  ListCount: " & .ListCount
    End With
    
    Debug.Print "ListBox cargado con " & (filaDestino - 1) & " registros"
    Exit Sub
ErrHandler:
    Debug.Print "Error en FiltrarYCargarListBox: " & Err.Description
End Sub
