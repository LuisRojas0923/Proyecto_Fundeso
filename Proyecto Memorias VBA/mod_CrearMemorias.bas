' Attribute VB_Name = "mod_ExportarListboxABaseDatos"
' === MANEJO DE ERRORES EXPLÍCITO EN EXPORTACIÓN ===
Public Sub ExportarListboxACrearHojas(frm As Object)
    On Error GoTo ManejoErrorGeneral
    Dim i As Long
    Dim nombreHoja As String
    Dim wsNueva As Worksheet
    Dim wsPlantilla As Worksheet
    Dim existeHoja As Boolean
    Dim filaExportada As Long
    Dim colNombreHoja As Long
    Dim wsControl As Worksheet

    Call LogInfo("Iniciando exportacion de datos del ListBox a hojas nuevas usando plantilla (solo renombrar)")

    Call LogDebug("ListCount del ListBox: " & frm.Listbox_Registros.ListCount)
    If frm.Listbox_Registros.ListCount = 0 Then
        MsgBox "No hay registros en la vista para exportar.", vbExclamation
        Call LogWarn("Exportacion cancelada: ListBox esta vacio")
        Exit Sub
    End If
    
    ' Configurar hoja de control (posición 2)
    Set wsControl = ThisWorkbook.Worksheets(6)
    Call LogDebug("Hoja de control tomada: '" & wsControl.Name & "'")
    wsControl.Visible = xlSheetVisible
    Call LogDebug("Hoja de control establecida como visible")

    filaExportada = 0
    colNombreHoja = 2 ' Columna 3 (�ndice 2)

    ' Verificar si la plantilla existe
    Set wsPlantilla = Nothing
    On Error Resume Next
    Set wsPlantilla = ThisWorkbook.Sheets("Plantilla")
    On Error GoTo ManejoErrorGeneral
    
    If wsPlantilla Is Nothing Then
        MsgBox "Error: No se encontro la hoja 'Plantilla'. La macro no puede continuar.", vbCritical
        Call LogError("Plantilla de creacion no encontrada. La macro no puede continuar.")
        Exit Sub
    End If
    
    Call LogInfo("Plantilla de Creacion encontrada (visible: " & (wsPlantilla.Visible = xlSheetVisible) & ").")

    For i = 0 To frm.Listbox_Registros.ListCount - 1
        Call LogDebug("Revisando registro: " & i & ", seleccionado: " & frm.Listbox_Registros.Selected(i))
        If frm.Listbox_Registros.Selected(i) Then
            ' Validar que haya fecha de inicio y fin en el ListBox
            If Trim(frm.Listbox_Registros.List(i, 3)) = "" Or Trim(frm.Listbox_Registros.List(i, 4)) = "" Then
                Call LogWarn("Registro " & i & " omitido: falta fecha de inicio o fin en el ListBox. Nombre hoja: '" & frm.Listbox_Registros.List(i, 2) & "'")
                MsgBox "El registro seleccionado no tiene fecha de inicio o fin. No se creará la hoja para este registro.", vbExclamation
                GoTo SiguienteRegistro
            End If
            On Error GoTo ManejoErrorHoja
            ' === Aquí se obtiene el nombre de la hoja desde la primera columna del ListBox (Concatenado) ===
            nombreHoja = CStr(frm.Listbox_Registros.List(i, 0))
            Call LogDebug("=== DIAGNÓSTICO DE DATOS DEL LISTBOX ===")
            Call LogDebug("Registro " & i & " - Datos completos:")
            Call LogDebug("  Columna 0 (Nombre hoja): '" & frm.Listbox_Registros.List(i, 0) & "'")
            Call LogDebug("  Columna 1 (Col1): '" & frm.Listbox_Registros.List(i, 1) & "'")
            Call LogDebug("  Columna 2 (Actividad): '" & frm.Listbox_Registros.List(i, 2) & "'")
            Call LogDebug("  Columna 3 (Unidad): '" & frm.Listbox_Registros.List(i, 3) & "'")
            Call LogDebug("  Columna 4 (Fecha Desde): '" & frm.Listbox_Registros.List(i, 4) & "'")
            Call LogDebug("  Columna 5 (Fecha Hasta): '" & frm.Listbox_Registros.List(i, 5) & "'")
            Call LogDebug("  Columna 6 (Observación): '" & frm.Listbox_Registros.List(i, 6) & "'")
            Call LogDebug("  Columna 7 (Area): '" & frm.Listbox_Registros.List(i, 7) & "'")
            Call LogDebug("Nombre de hoja a crear: '" & nombreHoja & "'")
            
            ' VALIDACIÓN: Evitar usar nombres de hojas del sistema
            If nombreHoja = "Nom_Tablas" Or nombreHoja = "Plantilla" Then
                Call LogError("*** NOMBRE DE HOJA INVÁLIDO: '" & nombreHoja & "' ***")
                Call LogError("Este nombre está reservado para hojas del sistema")
                MsgBox "Error: No se puede crear una memoria con el nombre '" & nombreHoja & "'" & vbCrLf & _
                       "Este nombre está reservado para hojas del sistema.", vbCritical, "Nombre de Hoja Inválido"
                GoTo SiguienteRegistro
            End If
            
            existeHoja = False
            For Each wsNueva In ThisWorkbook.Worksheets
                Call LogDebug("Comparando con hoja existente: '" & wsNueva.Name & "'")
                If wsNueva.Name = nombreHoja Then
                    existeHoja = True
                    Call LogWarn("La hoja '" & nombreHoja & "' ya existe.")
                    Exit For
                End If
            Next wsNueva
            If existeHoja Then
                MsgBox "La hoja '" & nombreHoja & "' ya existe. No se creará de nuevo.", vbExclamation
            Else
                Call LogInfo("Copiando plantilla para crear hoja nueva: '" & nombreHoja & "'")
                wsPlantilla.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set wsNueva = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ' === Aquí se valida y asigna el nombre de la hoja ===
                Dim nombreHojaValido As String
                nombreHojaValido = Left(nombreHoja, 31)
                nombreHojaValido = Replace(nombreHojaValido, "/", "-")
                nombreHojaValido = Replace(nombreHojaValido, "\\", "-")
                nombreHojaValido = Replace(nombreHojaValido, "*", "-")
                nombreHojaValido = Replace(nombreHojaValido, "[", "(")
                nombreHojaValido = Replace(nombreHojaValido, "]", ")")
                nombreHojaValido = Replace(nombreHojaValido, ":", "-")
                nombreHojaValido = Replace(nombreHojaValido, "?", "-")
                nombreHojaValido = Replace(nombreHojaValido, "'", "-")
                nombreHojaValido = Replace(nombreHojaValido, "\", "-")
                On Error GoTo ManejoErrorNombre
                wsNueva.Name = nombreHojaValido ' <--- ASIGNACIÓN DEL NOMBRE DE LA HOJA
                
                ' IMPORTANTE: Hacer visible la nueva hoja (hereda visibilidad de plantilla)
                wsNueva.Visible = xlSheetVisible
                Call LogDebug("Hoja '" & nombreHojaValido & "' establecida como visible")
                
                ' Activar la nueva hoja para que sea visible al usuario
                wsNueva.Activate
                Call LogDebug("Hoja '" & nombreHojaValido & "' activada")
                
                On Error GoTo ManejoErrorHoja
                Call LogOK("Hoja renombrada correctamente a '" & nombreHojaValido & "'")
                ' Asignar valores a la hoja creada desde la ListBox (OPTIMIZADO):
                On Error GoTo ManejoErrorAsignacion
                
                ' Preparar todos los datos en memoria para escritura en bloque
                Dim celdasDatos(1 To 6, 1 To 2) As Variant
                celdasDatos(1, 1) = "B4": celdasDatos(1, 2) = frm.Listbox_Registros.List(i, 7) ' Area
                celdasDatos(2, 1) = "B7": celdasDatos(2, 2) = nombreHojaValido ' Nombre concatenado
                celdasDatos(3, 1) = "D7": celdasDatos(3, 2) = frm.Listbox_Registros.List(i, 2) ' Actividad
                celdasDatos(4, 1) = "T7": celdasDatos(4, 2) = frm.Listbox_Registros.List(i, 3) ' Unidad
                celdasDatos(5, 1) = "O4": celdasDatos(5, 2) = frm.Listbox_Registros.List(i, 4) ' Fecha inicio
                celdasDatos(6, 1) = "S4": celdasDatos(6, 2) = frm.Listbox_Registros.List(i, 5) ' Fecha fin
                
                ' Escribir todos los datos de una vez (más eficiente)
                Dim j As Long
                For j = 1 To 6
                    wsNueva.Range(celdasDatos(j, 1)).Value = celdasDatos(j, 2)
                Next j
                
                Call LogDebug("Datos escritos en bloque para hoja: " & nombreHojaValido)
                On Error GoTo ManejoErrorHoja
                ' Cambiar el nombre de la primera tabla de la hoja
                Dim tbl As ListObject
                If wsNueva.ListObjects.Count > 0 Then
                    Set tbl = wsNueva.ListObjects(1)
                    Dim nuevoNombreTabla As String
                    nuevoNombreTabla = "TBL_" & nombreHojaValido
                    On Error Resume Next
                    tbl.Name = nuevoNombreTabla
                    If Err.Number <> 0 Then
                        Call LogError("Error al renombrar la tabla: " & Err.Description & " | Hoja: '" & nombreHojaValido & "' | Registro: " & i)
                        Err.Clear
                    Else
                        Call LogOK("Tabla renombrada a: " & nuevoNombreTabla & " | Hoja: '" & nombreHojaValido & "' | Registro: " & i)
                    End If
                    On Error GoTo 0
                    ' Registrar el nombre de la tabla en la hoja de control
                    On Error GoTo ManejoErrorRegistroTabla
                    Dim filaControl As Long
                    
                    Call LogDebug("=== REGISTRANDO TABLA EN HOJA DE CONTROL ===")
                    Call LogDebug("Usando hoja de control ya configurada: '" & wsControl.Name & "'")
                    
                    ' Buscar la primera celda vacía desde A2 hacia abajo
                    Call LogDebug("Buscando primera celda vacía en columna A...")
                    filaControl = 2
                    Do While filaControl <= 1000 And wsControl.Cells(filaControl, 1).Value <> ""
                        filaControl = filaControl + 1
                    Loop
                    Call LogDebug("Primera celda vacía encontrada en fila: " & filaControl)
                    
                    ' Escribir el nombre de la tabla en la hoja de control
                    Call LogDebug("Escribiendo nombre de tabla en " & wsControl.Name & "!A" & filaControl)
                    wsControl.Cells(filaControl, 1).Value = nuevoNombreTabla
                    Call LogOK("*** NOMBRE DE TABLA REGISTRADO EN HOJA DE CONTROL ***")
                    Call LogOK("Tabla '" & nuevoNombreTabla & "' → " & wsControl.Name & "!A" & filaControl & " | Registro: " & i)
                    
ContinuarSinRegistro:
                    On Error GoTo ManejoErrorHoja
                Else
                    Call LogWarn("No se encontro tabla para renombrar en la hoja '" & nombreHojaValido & "' | Registro: " & i)
                End If
            End If
            frm.Listbox_Registros.Selected(i) = False ' Deseleccionar después de exportar
            filaExportada = filaExportada + 1
        End If
SiguienteRegistro:
        On Error GoTo ManejoErrorGeneral
    Next i

    Call LogInfo("Total de hojas creadas/exportadas: " & filaExportada)
    If filaExportada = 0 Then
        MsgBox "No seleccionaste registros para exportar.", vbExclamation
        Call LogInfo("Exportacion cancelada: ningun elemento estaba seleccionado.")
    Else
        MsgBox "Las hojas se crearon correctamente desde la plantilla.", vbInformation
        Call LogOK("Exportacion completada con " & filaExportada & " hojas creadas.")
    End If
    
    ' Ocultar la hoja de control al finalizar
    If Not wsControl Is Nothing Then
        wsControl.Visible = xlSheetHidden
        Call LogDebug("Hoja de control ocultada al finalizar el proceso")
    End If
    Exit Sub

ManejoErrorNombre:
    MsgBox "Error al asignar el nombre de la hoja: " & Err.Description, vbCritical
    Call LogErrorVBA("ManejoErrorNombre", Err.Description)
    Resume Next
ManejoErrorAsignacion:
    MsgBox "Error al asignar valores a la hoja: " & Err.Description & vbCrLf & _
           "Registro: " & i & vbCrLf & _
           "Nombre hoja: " & nombreHojaValido, vbCritical
    Call LogErrorVBA("ManejoErrorAsignacion", Err.Description & " (Registro: " & i & ")")
    Resume Next
ManejoErrorHoja:
    MsgBox "Error en la creación de la hoja: " & Err.Description & vbCrLf & _
           "Registro: " & i & vbCrLf & _
           "Nombre hoja: " & nombreHoja, vbCritical
    Call LogErrorVBA("ManejoErrorHoja", Err.Description & " (Registro: " & i & ")")
    Resume Next
ManejoErrorGeneral:
    MsgBox "Error general al exportar: " & Err.Description, vbCritical
    Call LogErrorVBA("ManejoErrorGeneral", Err.Description)
    
    ' Ocultar la hoja de control en caso de error
    If Not wsControl Is Nothing Then
        On Error Resume Next
        wsControl.Visible = xlSheetHidden
        Call LogDebug("Hoja de control ocultada después de error general")
        On Error GoTo 0
    End If
ManejoErrorRegistroTabla:
    Call LogError("No se pudo registrar el nombre de la tabla '" & nuevoNombreTabla & "' en hoja de control. Error: " & Err.Description & " | Registro: " & i)
    MsgBox "No se pudo registrar el nombre de la tabla '" & nuevoNombreTabla & "' en la hoja de control.", vbExclamation
    Resume Next
End Sub



