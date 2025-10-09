Attribute VB_Name = "Inicializar_Hoja_Config"
Option Explicit

' =======================================================
' MODULO DE INICIALIZACION DE HOJA DE CONFIGURACION
' =======================================================
' Este modulo contiene la macro para crear la hoja Config_Sistema
' con los usuarios iniciales y configurarla correctamente.
'
' IMPORTANTE: Ejecuta este modulo UNA SOLA VEZ al inicio
' =======================================================

Private Const NOMBRE_HOJA_CONFIG As String = "Config_Sistema"
Private Const PASSWORD_PROTECCION As String = "SistemaSeguridadVBA2024"

' Procedimiento principal para crear e inicializar la hoja de configuracion
Public Sub CrearHojaConfiguracion()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim wsExistente As Worksheet
    Dim respuesta As VbMsgBoxResult
    
    ' Verificar si la hoja ya existe
    On Error Resume Next
    Set wsExistente = ThisWorkbook.Worksheets(NOMBRE_HOJA_CONFIG)
    On Error GoTo ErrHandler
    
    If Not wsExistente Is Nothing Then
        respuesta = MsgBox("La hoja '" & NOMBRE_HOJA_CONFIG & "' ya existe." & vbCrLf & vbCrLf & _
                          "¿Desea reemplazarla?" & vbCrLf & _
                          "ADVERTENCIA: Se perderan todos los usuarios existentes.", _
                          vbQuestion + vbYesNo + vbDefaultButton2, "Hoja Existente")
        
        If respuesta = vbNo Then
            MsgBox "Operacion cancelada.", vbInformation
            Exit Sub
        Else
            ' Hacer visible y desproteger antes de eliminar
            wsExistente.Visible = xlSheetVisible
            wsExistente.Unprotect Password:=PASSWORD_PROTECCION
            Application.DisplayAlerts = False
            wsExistente.Delete
            Application.DisplayAlerts = True
        End If
    End If
    
    Call RegistrarLog("CrearHojaConfiguracion", "Iniciando creacion de hoja de configuracion", LOG_INFO)
    
    ' Crear nueva hoja
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = NOMBRE_HOJA_CONFIG
    
    ' Configurar encabezados
    With ws
        ' Fila de encabezados
        .Cells(1, 1).Value = "Usuario"
        .Cells(1, 2).Value = "Contrasena"
        .Cells(1, 3).Value = "Estado"
        
        ' Formato de encabezados
        With .Range("A1:C1")
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(68, 114, 196) ' Azul
            .Font.Color = RGB(255, 255, 255) ' Blanco
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Datos iniciales
        .Cells(2, 1).Value = "admin"
        .Cells(2, 2).Value = "1234"
        .Cells(2, 3).Value = "Activo"
        
        .Cells(3, 1).Value = "usuario1"
        .Cells(3, 2).Value = "pass1"
        .Cells(3, 3).Value = "Activo"
        
        .Cells(4, 1).Value = "usuario2"
        .Cells(4, 2).Value = "pass2"
        .Cells(4, 3).Value = "Activo"
        
        ' Formato de datos
        With .Range("A2:C4")
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Ajustar ancho de columnas
        .Columns("A:C").AutoFit
        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 12
        
        ' Agregar notas de instrucciones
        .Cells(6, 1).Value = "INSTRUCCIONES:"
        .Cells(6, 1).Font.Bold = True
        .Cells(7, 1).Value = "1. Columna A: Nombre de usuario (sin espacios)"
        .Cells(8, 1).Value = "2. Columna B: Contrasena del usuario"
        .Cells(9, 1).Value = "3. Columna C: Estado (Activo o Inactivo)"
        .Cells(10, 1).Value = "4. Solo usuarios con estado 'Activo' podran iniciar sesion"
        .Cells(11, 1).Value = "5. El usuario 'admin' tiene permisos especiales"
        .Range("A7:A11").Font.Italic = True
        .Range("A7:A11").Font.Size = 9
    End With
    
    ' Proteger la hoja
    ws.Protect Password:=PASSWORD_PROTECCION, _
              DrawingObjects:=True, _
              Contents:=True, _
              Scenarios:=True
    
    ' Ocultar la hoja (muy oculta - no visible desde interfaz normal)
    ws.Visible = xlSheetVeryHidden
    
    Call RegistrarLog("CrearHojaConfiguracion", "Hoja de configuracion creada exitosamente", LOG_INFO)
    
    MsgBox "Hoja de configuracion creada exitosamente!" & vbCrLf & vbCrLf & _
           "Usuarios creados:" & vbCrLf & _
           "- admin / 1234" & vbCrLf & _
           "- usuario1 / pass1" & vbCrLf & _
           "- usuario2 / pass2" & vbCrLf & vbCrLf & _
           "La hoja ha sido ocultada y protegida." & vbCrLf & _
           "Para acceder a ella, inicia sesion como admin y usa el boton de Configuracion.", _
           vbInformation, "Hoja Creada"
    
    Exit Sub

ErrHandler:
    Call RegistrarError("CrearHojaConfiguracion", Err.Description)
    MsgBox "Error al crear la hoja de configuracion: " & vbCrLf & Err.Description, vbCritical, "Error"
End Sub

' Procedimiento auxiliar para hacer visible la hoja Config_Sistema temporalmente
' (util para debugging o recuperacion)
Public Sub MostrarHojaConfigManual()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim password As String
    
    password = InputBox("Ingrese la contrasena de administrador:", "Autenticacion Requerida")
    
    If password <> "1234" Then
        MsgBox "Contrasena incorrecta", vbCritical
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets(NOMBRE_HOJA_CONFIG)
    ws.Unprotect Password:=PASSWORD_PROTECCION
    ws.Visible = xlSheetVisible
    ws.Activate
    
    MsgBox "Hoja visible. Recuerda volver a ocultarla cuando termines.", vbInformation
    
    Exit Sub

ErrHandler:
    Call RegistrarError("MostrarHojaConfigManual", Err.Description)
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Procedimiento auxiliar para ocultar la hoja Config_Sistema
Public Sub OcultarHojaConfigManual()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(NOMBRE_HOJA_CONFIG)
    ws.Protect Password:=PASSWORD_PROTECCION
    ws.Visible = xlSheetVeryHidden
    
    MsgBox "Hoja ocultada y protegida correctamente.", vbInformation
    
    Exit Sub

ErrHandler:
    Call RegistrarError("OcultarHojaConfigManual", Err.Description)
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

