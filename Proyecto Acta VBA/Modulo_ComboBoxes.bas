Attribute VB_Name = "Modulo_ComboBoxes"

' === MÓDULO PARA GESTIÓN DE COMBOBOXES 8:04 PM 20/08/2025 ===

' Cargar ComboBox de Áreas
Public Sub CargarComboBoxAreas(frm As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet, tbl As ListObject
    Dim nombreTabla As String
    Dim tablaEncontrada As Boolean
    Dim i As Long
    
    nombreTabla = "Cons_Presupuesto"
    tablaEncontrada = False
    
    ' Buscar la tabla Cons_Presupuesto en todas las hojas
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
        Debug.Print "ADVERTENCIA: La tabla '" & nombreTabla & "' no existe en el libro."
        Exit Sub
    End If
    
    ' Limpiar el ComboBox
    frm.cmb_Area.Clear
    
    ' Cargar datos concatenando columnas 1 y 2, filtrando valores vacíos de columna 1
    With tbl.DataBodyRange
        For i = 1 To .Rows.Count
            Dim valorCol1 As String, valorCol2 As String
            valorCol1 = Trim(CStr(.Cells(i, 1).Value))
            valorCol2 = Trim(CStr(.Cells(i, 2).Value))
            
            ' Solo agregar si la columna 1 no está vacía
            If valorCol1 <> "" And valorCol1 <> "0" Then
                Dim valorConcatenado As String
                valorConcatenado = valorCol1 & " - " & valorCol2
                frm.cmb_Area.AddItem valorConcatenado
            End If
        Next i
    End With
    
    Debug.Print "ComboBox de Áreas cargado con " & frm.cmb_Area.ListCount & " elementos"
    Exit Sub
ErrHandler:
    Debug.Print "Error en CargarComboBoxAreas: " & Err.Description
End Sub

' Cargar ComboBox de Capítulos
Public Sub CargarComboBoxCapitulos(frm As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet, tbl As ListObject
    Dim nombreTabla As String
    Dim tablaEncontrada As Boolean
    Dim i As Long
    
    nombreTabla = "Capitulos"
    tablaEncontrada = False
    
    ' Buscar la tabla Capitulos en todas las hojas
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
        Debug.Print "ADVERTENCIA: La tabla '" & nombreTabla & "' no existe en el libro."
        Exit Sub
    End If
    
    ' Limpiar el ComboBox
    frm.cmb_Capitulos.Clear
    
    ' Cargar datos concatenando columnas 1 (Consecutivo) y 2 (Nombre)
    With tbl.DataBodyRange
        For i = 1 To .Rows.Count
            Dim valorCol1 As String, valorCol2 As String
            valorCol1 = Trim(CStr(.Cells(i, 1).Value))
            valorCol2 = Trim(CStr(.Cells(i, 2).Value))
            
            ' Solo agregar si el consecutivo no está vacío
            If valorCol1 <> "" And valorCol1 <> "0" Then
                Dim valorConcatenado As String
                valorConcatenado = valorCol1 & " - " & valorCol2
                frm.cmb_Capitulos.AddItem valorConcatenado
            End If
        Next i
    End With
    
    Debug.Print "ComboBox de Capítulos cargado con " & frm.cmb_Capitulos.ListCount & " elementos"
    Exit Sub
ErrHandler:
    Debug.Print "Error en CargarComboBoxCapitulos: " & Err.Description
End Sub

' Extraer consecutivo y capítulo del valor seleccionado (ej: "1 - ESTRUCTURA")
Public Sub ObtenerValorComboBoxCapitulo(capituloCompleto As String, ByRef consecutivo As String, ByRef capitulo As String)
    If InStr(capituloCompleto, " - ") > 0 Then
        consecutivo = Trim(Left(capituloCompleto, InStr(capituloCompleto, " - ") - 1))
        capitulo = Trim(Mid(capituloCompleto, InStr(capituloCompleto, " - ") + 2))
    Else
        ' Si no hay separador, asumimos que es solo el nombre y asignamos un consecutivo por defecto.
        consecutivo = "1" ' O manejar el error como se prefiera
        capitulo = capituloCompleto
    End If
    
    Debug.Print "Capítulo completo seleccionado: '" & capituloCompleto & "'"
    Debug.Print "Consecutivo extraído: '" & consecutivo & "'"
    Debug.Print "Capítulo extraído: '" & capitulo & "'"
End Sub

' Extraer consecutivo y área del valor seleccionado (ej: "1 - UBA")
Public Sub ObtenerValorComboBoxArea(areaCompleta As String, ByRef consecutivo As String, ByRef area As String)
    If InStr(areaCompleta, " - ") > 0 Then
        consecutivo = Trim(Left(areaCompleta, InStr(areaCompleta, " - ") - 1))
        area = Trim(Mid(areaCompleta, InStr(areaCompleta, " - ") + 2))
    Else
        consecutivo = "1"
        area = areaCompleta
    End If
    
    Debug.Print "Área completa seleccionada: '" & areaCompleta & "'"
    Debug.Print "Consecutivo extraído: '" & consecutivo & "'"
    Debug.Print "Área extraída: '" & area & "'"
End Sub
