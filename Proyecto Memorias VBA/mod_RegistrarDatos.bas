' === Módulo: mod_RegistrarDatos ===
' Lógica para registrar los valores de F_Desde y F_Hasta en las columnas siguientes del ListBox

Option Explicit

Public FDesdeValor As Variant
Public FHastaValor As Variant

' Guardar el valor de F_Desde cuando se activa el control
Public Sub GuardarFDesde(valor As Variant)
    FDesdeValor = valor
    Debug.Print "FDesdeValor guardado: " & FDesdeValor
End Sub

' Guardar el valor de F_Hasta cuando se activa el control
Public Sub GuardarFHasta(valor As Variant)
    FHastaValor = valor
    Debug.Print "FHastaValor guardado: " & FHastaValor
End Sub

' Registrar los valores en las columnas siguientes del ListBox
' === AJUSTE: Mostrar solo la fila seleccionada en el ListBox y agregar las fechas al lado ===
' === CORRECCIÓN: Manejo seguro para asignar .List en el ListBox ===
' === MANEJO DE ERRORES DETALLADO EN RegistrarFechasEnListBox ===
' === DEPURACIÓN: Mostrar contenido del array antes de asignar al ListBox ===
' === CORRECCIÓN: Solo usar el array para mostrar fechas, no modificar directamente el ListBox ===
' Elimina cualquier intento de asignar valores directamente a frm.Listbox_Registros.List(i, x)
' Solo se construye el array y se asigna a .List
Public Sub RegistrarFechasEnListBox(frm As Object)
    On Error GoTo ErrorHandlerGeneral
    Dim i As Long, datos() As Variant
    Dim desde As Variant, hasta As Variant, obs As String
    desde = frm.F_Desde.Value
    hasta = frm.F_Hasta.Value
    obs = frm.Observaciones.Value
    Dim nFilas As Long, nCols As Long
    nFilas = frm.Listbox_Registros.ListCount
    nCols = 7
    If nFilas = 0 Then Exit Sub
    datos = frm.Listbox_Registros.List
    Dim algunoSeleccionado As Boolean: algunoSeleccionado = False
    For i = 0 To nFilas - 1
        If frm.Listbox_Registros.Selected(i) Then
            datos(i, 4) = desde ' Fecha Desde en columna 5
            datos(i, 5) = hasta ' Fecha Hasta en columna 6
            datos(i, 6) = obs   ' Observación en columna 7
            algunoSeleccionado = True
        End If
    Next i
    If algunoSeleccionado Then
        With frm.Listbox_Registros
            .Clear
            .ColumnCount = nCols
            .ColumnWidths = "60 pt;300 pt;60 pt;60 pt;80 pt;80 pt;120 pt"
            .List = datos
            Debug.Print "[mod_RegistrarDatos.RegistrarFechasEnListBox] ColumnWidths: " & .ColumnWidths
            Debug.Print "[mod_RegistrarDatos.RegistrarFechasEnListBox] ListBox.Width: " & .Width
            Debug.Print "[mod_RegistrarDatos.RegistrarFechasEnListBox] Ancho calculado (twips): " & .Width
        End With
    Else
        MsgBox "Debes seleccionar al menos una fila antes de registrar las fechas.", vbExclamation
    End If
    Exit Sub
ErrorHandlerGeneral:
    MsgBox "Error general en RegistrarFechasEnListBox: " & Err.Description, vbCritical
    Debug.Print "Error general en RegistrarFechasEnListBox: " & Err.Description
End Sub


