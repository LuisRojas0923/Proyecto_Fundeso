' Attribute VB_Name = "ACTUALIZAR"
Sub ActualizarTablasPowerQuery()
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject
    Dim totalTablas As Integer
    Dim tablasActualizadas As Integer
    Dim tablasSinConexion As Integer
    
    
    ' Optimizaciones de rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Buscando tablas de Power Query..."
    
    totalTablas = 0
    tablasActualizadas = 0
    tablasSinConexion = 0
    
    ' Recorrer todas las hojas del libro
    For Each ws In ThisWorkbook.Worksheets
        ' Buscar y actualizar todas las QueryTables (Power Query en versiones antiguas)
        For Each qt In ws.QueryTables
            totalTablas = totalTablas + 1
            Application.StatusBar = "Actualizando tabla de Power Query en " & ws.Name & "..."
            qt.Refresh BackgroundQuery:=False
            tablasActualizadas = tablasActualizadas + 1
        Next qt
        
        ' Buscar y actualizar todas las ListObjects con conexi�n de Power Query
        For Each lo In ws.ListObjects
            On Error Resume Next ' Manejar error si QueryTable no existe
            If lo.SourceType = xlSrcQuery Then
                totalTablas = totalTablas + 1
                Application.StatusBar = "Actualizando tabla con conexi�n en " & ws.Name & "..."
                lo.QueryTable.Refresh BackgroundQuery:=False
                tablasActualizadas = tablasActualizadas + 1
            Else
                tablasSinConexion = tablasSinConexion + 1 ' Contar las tablas sin conexi�n
            End If
            On Error GoTo 0 ' Restaurar manejo de errores
        Next lo
    Next ws
    
    Set ws = ThisWorkbook.Sheets("REGISTRO")
    
    ws.Range("M2").Value = "Actualizaci�n completada"
    ws.Range("M1").Value = Now
    ' Restaurar configuraciones de aplicaci�n
    Application.StatusBar = "Actualizaci�n completada: " & tablasActualizadas & " de " & totalTablas & " tablas con conexi�n actualizadas. "
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Actualizaci�n completada.", vbInformation
    Application.StatusBar = False
End Sub

