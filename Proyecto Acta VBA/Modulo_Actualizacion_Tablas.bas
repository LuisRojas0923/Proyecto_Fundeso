Attribute VB_Name = "Modulo_Actualizacion_Tablas"
Option Explicit

' Macro para encontrar y actualizar la tabla especifica 'EXPORTE_PRESUPUESTO_1'.
' Muestra el progreso en la barra de estado de Excel y guarda el libro al finalizar.

Public Sub ActualizarTablasConQuery()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim totalQueries As Long
    Dim refreshedCount As Long
    Dim originalStatusBar As Variant
    Dim blnGuardadoExitoso As Boolean
    
    ' Guardar el estado original de la barra de estado
    originalStatusBar = Application.StatusBar
    
    ' Desactivar actualizacion de pantalla para mejor rendimiento
    Application.ScreenUpdating = False
    
    ' --- PASO 1: Contar cuantas de las tablas especificas existen ---
    totalQueries = 0
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            ' Verificar si la tabla es la que queremos actualizar
            If tbl.SourceType = xlSrcQuery And tbl.Name = "EXPORTE_PRESUPUESTO_1" Then
                totalQueries = totalQueries + 1
            End If
        Next tbl
    Next ws
    
    If totalQueries = 0 Then
        MsgBox "No se encontro la tabla 'EXPORTE_PRESUPUESTO_1' en este libro.", vbInformation, "Proceso Terminado"
        Exit Sub
    End If
    
    ' --- PASO 2: Actualizar cada tabla y mostrar el progreso ---
    refreshedCount = 0
    RegistrarInfo "ActualizarTablasConQuery", "Iniciando actualizacion de queries especificas"
    
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            ' Actualizar solo la tabla especifica
            If tbl.SourceType = xlSrcQuery And tbl.Name = "EXPORTE_PRESUPUESTO_1" Then
                refreshedCount = refreshedCount + 1
                
                ' Actualizar la barra de estado
                Application.StatusBar = "Actualizando " & refreshedCount & " de " & totalQueries & ": '" & tbl.Name & "' en la hoja '" & ws.Name & "'..."
                RegistrarInfo "ActualizarTablasConQuery", "Actualizando " & refreshedCount & "/" & totalQueries & " - Tabla: " & tbl.Name & " en hoja: " & ws.Name
                
                ' Actualizar la consulta. BackgroundQuery:=False asegura que VBA espere
                ' a que la actualizacion termine antes de continuar.
                tbl.QueryTable.Refresh BackgroundQuery:=False
            End If
        Next tbl
    Next ws
    
    ' --- PASO 3: Guardar el libro ---
    Application.StatusBar = "Guardando libro..."
    RegistrarInfo "ActualizarTablasConQuery", "Iniciando guardado del libro"
    
    On Error Resume Next
    ThisWorkbook.Save
    If Err.Number = 0 Then
        blnGuardadoExitoso = True
        RegistrarInfo "ActualizarTablasConQuery", "Libro guardado exitosamente"
        Application.StatusBar = "Libro guardado exitosamente."
    Else
        blnGuardadoExitoso = False
        RegistrarError "ActualizarTablasConQuery", "Error al guardar: " & Err.Description
        Application.StatusBar = "Error al guardar el libro."
    End If
    On Error GoTo ErrHandler
    
    ' --- PASO 4: Finalizar y limpiar ---
    Application.StatusBar = "Proceso completado."
    
    ' Mensaje final con informacion del guardado
    Dim strMensajeFinal As String
    strMensajeFinal = refreshedCount & " de " & totalQueries & " tabla(s) han sido actualizada(s)."
    If blnGuardadoExitoso Then
        strMensajeFinal = strMensajeFinal & vbCrLf & "El libro ha sido guardado exitosamente."
    Else
        strMensajeFinal = strMensajeFinal & vbCrLf & "ADVERTENCIA: No se pudo guardar el libro."
    End If
    
    MsgBox strMensajeFinal, vbInformation, "Proceso Finalizado"
    RegistrarInfo "ActualizarTablasConQuery", "Proceso completado. Tablas actualizadas: " & refreshedCount & "/" & totalQueries & ". Guardado exitoso: " & blnGuardadoExitoso
    
    ' Restaurar configuraciones
    Application.ScreenUpdating = True
    Application.StatusBar = originalStatusBar
    
    Exit Sub
    
ErrHandler:
    RegistrarError "ActualizarTablasConQuery", "Error durante la actualizacion: " & Err.Description
    MsgBox "Ocurrio un error durante la actualizacion: " & vbCrLf & Err.Description, vbCritical, "Error"
    ' Asegurarse de restaurar configuraciones incluso si hay un error
    Application.ScreenUpdating = True
    Application.StatusBar = originalStatusBar
End Sub
