Attribute VB_Name = "Modulo_Seguridad"
Option Explicit

' === MODULO DE SEGURIDAD Y AUTENTICACION ===
' Este modulo centraliza las funciones relacionadas con la seguridad y autenticacion de usuarios.
' Incluye proteccion de libro, validacion de credenciales y gestion de acceso a configuracion.

' Constantes de configuracion
Private Const NOMBRE_HOJA_CONFIG As String = "Config_Sistema"
Private Const PASSWORD_PROTECCION As String = "SistemaSeguridadVBA2024"

' Variable publica para almacenar el usuario actual logueado
Public UsuarioActual As String

' ===================================
' FUNCION PRINCIPAL DE AUTENTICACION
' ===================================
' Muestra el formulario de inicio de sesion y devuelve True si el login fue exitoso.
' Argumentos:
'   - cerrarSiFalla: Si es True y falla la autenticacion, cierra el libro (uso en Workbook_Open)
'                    Si es False, solo retorna False (uso en operaciones criticas)
Public Function AutenticarUsuario(Optional ByVal cerrarSiFalla As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    
    ' Por defecto, la autenticacion falla hasta que se demuestre lo contrario
    AutenticarUsuario = False
    
    Call RegistrarLog("AutenticarUsuario", "Iniciando proceso de autenticacion...", LOG_INFO)
    
    ' Mostrar el formulario de inicio de sesion de forma modal
    InicioSesion.Show vbModal
    
    ' Verificar si el login fue exitoso
    If InicioSesion.LoginExitoso Then
        Call RegistrarLog("AutenticarUsuario", "Autenticacion exitosa para usuario: " & UsuarioActual, LOG_INFO)
        AutenticarUsuario = True
    Else
        Call RegistrarLog("AutenticarUsuario", "Autenticacion fallida o cancelada", LOG_WARNING)
        
        ' Si se solicito cerrar el libro en caso de fallo
        If cerrarSiFalla Then
            Call RegistrarLog("AutenticarUsuario", "Cerrando libro por fallo de autenticacion", LOG_INFO)
            Call CerrarLibroSinGuardar
        End If
    End If
    
    ' Descargar el formulario de la memoria para limpiar su estado
    Unload InicioSesion
    
    Exit Function

ErrHandler:
    Call RegistrarError("AutenticarUsuario", Err.Description)
    ' Asegurarse de descargar el formulario en caso de error
    On Error Resume Next
    Unload InicioSesion
    On Error GoTo 0
    AutenticarUsuario = False
End Function

' ===================================
' VALIDACION DE CREDENCIALES
' ===================================
' Valida usuario y contrasena contra la hoja Config_Sistema
' Argumentos:
'   - strUsuario: Nombre de usuario a validar
'   - strPassword: Contrasena a validar
' Devuelve: True si las credenciales son correctas y el usuario esta activo
Public Function ValidarCredencialesDesdeHoja(ByVal strUsuario As String, ByVal strPassword As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim usuarioHoja As String
    Dim passwordHoja As String
    Dim estadoHoja As String
    
    ValidarCredencialesDesdeHoja = False
    
    ' Obtener referencia a la hoja de configuracion
    Set ws = ObtenerHojaConfig()
    If ws Is Nothing Then Exit Function
    
    ' Encontrar la ultima fila con datos
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Recorrer la tabla de usuarios (desde fila 2, asumiendo encabezados en fila 1)
    For i = 2 To ultimaFila
        usuarioHoja = Trim(ws.Cells(i, 1).Value)
        passwordHoja = Trim(ws.Cells(i, 2).Value)
        estadoHoja = Trim(ws.Cells(i, 3).Value)
        
        ' Verificar si coinciden las credenciales y el usuario esta activo
        If usuarioHoja = strUsuario And passwordHoja = strPassword Then
            If UCase(estadoHoja) = "ACTIVO" Then
                ValidarCredencialesDesdeHoja = True
                UsuarioActual = strUsuario
                Call RegistrarLog("ValidarCredencialesDesdeHoja", "Credenciales validas para: " & strUsuario, LOG_INFO)
            Else
                Call RegistrarAdvertencia("ValidarCredencialesDesdeHoja", "Usuario inactivo: " & strUsuario)
            End If
            Exit Function
        End If
    Next i
    
    Call RegistrarAdvertencia("ValidarCredencialesDesdeHoja", "Credenciales invalidas para: " & strUsuario)
    
    Exit Function

ErrHandler:
    Call RegistrarError("ValidarCredencialesDesdeHoja", Err.Description)
    ValidarCredencialesDesdeHoja = False
End Function

' ===================================
' CARGA DE USUARIOS PARA COMBOBOX
' ===================================
' Lee los usuarios activos de la hoja Config_Sistema y los devuelve como array
' Devuelve: Array con los nombres de usuarios activos, o array vacio si hay error
Public Function CargarUsuariosEnComboBox() As Variant
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim usuarioHoja As String
    Dim estadoHoja As String
    Dim listaUsuarios() As String
    Dim contador As Long
    
    ' Obtener referencia a la hoja de configuracion
    Set ws = ObtenerHojaConfig()
    If ws Is Nothing Then
        CargarUsuariosEnComboBox = Array()
        Exit Function
    End If
    
    ' Encontrar la ultima fila con datos
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Dimensionar array (tamanio maximo posible)
    ReDim listaUsuarios(1 To ultimaFila - 1)
    contador = 0
    
    ' Recorrer la tabla de usuarios (desde fila 2)
    For i = 2 To ultimaFila
        usuarioHoja = Trim(ws.Cells(i, 1).Value)
        estadoHoja = Trim(ws.Cells(i, 3).Value)
        
        ' Solo agregar usuarios activos
        If UCase(estadoHoja) = "ACTIVO" And usuarioHoja <> "" Then
            contador = contador + 1
            listaUsuarios(contador) = usuarioHoja
        End If
    Next i
    
    ' Redimensionar array al tamanio exacto
    If contador > 0 Then
        ReDim Preserve listaUsuarios(1 To contador)
        CargarUsuariosEnComboBox = listaUsuarios
        Call RegistrarLog("CargarUsuariosEnComboBox", "Cargados " & contador & " usuarios activos", LOG_INFO)
    Else
        CargarUsuariosEnComboBox = Array()
        Call RegistrarAdvertencia("CargarUsuariosEnComboBox", "No se encontraron usuarios activos")
    End If
    
    Exit Function

ErrHandler:
    Call RegistrarError("CargarUsuariosEnComboBox", Err.Description)
    CargarUsuariosEnComboBox = Array()
End Function

' ===================================
' PROTECCION DEL LIBRO
' ===================================
' Protege todas las hojas del libro y la estructura
Public Sub ProtegerLibroCompleto()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim contadorProtegidas As Long
    
    Application.ScreenUpdating = False
    contadorProtegidas = 0
    
    ' Proteger cada hoja del libro (excepto Config_Sistema que ya esta protegida)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> NOMBRE_HOJA_CONFIG Then
            If Not ws.ProtectContents Then
                ws.Protect Password:=PASSWORD_PROTECCION, _
                          DrawingObjects:=True, _
                          Contents:=True, _
                          Scenarios:=True
                contadorProtegidas = contadorProtegidas + 1
            End If
        End If
    Next ws
    
    ' Proteger estructura del libro
    ThisWorkbook.Protect Password:=PASSWORD_PROTECCION, Structure:=True, Windows:=False
    
    Application.ScreenUpdating = True
    
    Call RegistrarLog("ProtegerLibroCompleto", "Protegidas " & contadorProtegidas & " hojas y estructura del libro", LOG_INFO)
    
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Call RegistrarError("ProtegerLibroCompleto", Err.Description)
End Sub

' ===================================
' DESPROTECCION DEL LIBRO
' ===================================
' Desprotege todas las hojas del libro y la estructura
Public Sub DesprotegerLibroCompleto()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim contadorDesprotegidas As Long
    
    Application.ScreenUpdating = False
    contadorDesprotegidas = 0
    
    ' Desproteger estructura del libro
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PASSWORD_PROTECCION
    On Error GoTo ErrHandler
    
    ' Desproteger cada hoja del libro (excepto Config_Sistema)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> NOMBRE_HOJA_CONFIG Then
            If ws.ProtectContents Then
                ws.Unprotect Password:=PASSWORD_PROTECCION
                contadorDesprotegidas = contadorDesprotegidas + 1
            End If
        End If
    Next ws
    
    Application.ScreenUpdating = True
    
    Call RegistrarLog("DesprotegerLibroCompleto", "Desprotegidas " & contadorDesprotegidas & " hojas y estructura del libro", LOG_INFO)
    
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Call RegistrarError("DesprotegerLibroCompleto", Err.Description)
End Sub

' ===================================
' CIERRE DE LIBRO SIN GUARDAR
' ===================================
' Cierra el libro actual sin guardar cambios
Public Sub CerrarLibroSinGuardar()
    On Error Resume Next
    
    Call RegistrarLog("CerrarLibroSinGuardar", "Cerrando libro sin guardar cambios", LOG_WARNING)
    
    ' Desactivar eventos para evitar bucles
    Application.EnableEvents = False
    
    ' Cerrar el libro sin guardar
    ThisWorkbook.Saved = True ' Marca el libro como guardado para evitar dialogo
    ThisWorkbook.Close SaveChanges:=False
    
    Application.EnableEvents = True
End Sub

' ===================================
' GESTION DE HOJA DE CONFIGURACION
' ===================================
' Abre y muestra la hoja Config_Sistema para edicion (solo admin)
Public Sub AbrirHojaConfiguracion()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    ' Verificar que el usuario actual es admin
    If UCase(UsuarioActual) <> "ADMIN" Then
        MsgBox "Solo el administrador puede acceder a la configuracion de usuarios.", vbExclamation, "Acceso Denegado"
        Call RegistrarAdvertencia("AbrirHojaConfiguracion", "Intento de acceso denegado para usuario: " & UsuarioActual)
        Exit Sub
    End If
    
    ' Obtener referencia a la hoja
    Set ws = ObtenerHojaConfig()
    If ws Is Nothing Then Exit Sub
    
    ' Desproteger la hoja
    ws.Unprotect Password:=PASSWORD_PROTECCION
    
    ' Hacer visible la hoja
    ws.Visible = xlSheetVisible
    
    ' Activar la hoja
    ws.Activate
    
    Call RegistrarLog("AbrirHojaConfiguracion", "Hoja de configuracion abierta por: " & UsuarioActual, LOG_INFO)
    
    MsgBox "Hoja de configuracion abierta." & vbCrLf & vbCrLf & _
           "Estructura:" & vbCrLf & _
           "Columna A: Usuario" & vbCrLf & _
           "Columna B: Contrasena" & vbCrLf & _
           "Columna C: Estado (Activo/Inactivo)" & vbCrLf & vbCrLf & _
           "IMPORTANTE: Cierra esta hoja cuando termines de editarla.", _
           vbInformation, "Configuracion de Usuarios"
    
    Exit Sub

ErrHandler:
    Call RegistrarError("AbrirHojaConfiguracion", Err.Description)
    MsgBox "Error al abrir la hoja de configuracion: " & Err.Description, vbCritical
End Sub

' Cierra y oculta la hoja Config_Sistema
Public Sub CerrarHojaConfiguracion()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    ' Obtener referencia a la hoja
    Set ws = ObtenerHojaConfig()
    If ws Is Nothing Then Exit Sub
    
    ' Proteger la hoja
    ws.Protect Password:=PASSWORD_PROTECCION, _
              DrawingObjects:=True, _
              Contents:=True, _
              Scenarios:=True
    
    ' Ocultar la hoja (muy oculta)
    ws.Visible = xlSheetVeryHidden
    
    Call RegistrarLog("CerrarHojaConfiguracion", "Hoja de configuracion cerrada y protegida", LOG_INFO)
    
    MsgBox "Hoja de configuracion protegida y ocultada correctamente.", vbInformation, "Configuracion Guardada"
    
    Exit Sub

ErrHandler:
    Call RegistrarError("CerrarHojaConfiguracion", Err.Description)
    MsgBox "Error al cerrar la hoja de configuracion: " & Err.Description, vbCritical
End Sub

' ===================================
' FUNCIONES AUXILIARES PRIVADAS
' ===================================
' Obtiene referencia a la hoja Config_Sistema
' Devuelve: Objeto Worksheet o Nothing si no existe
Private Function ObtenerHojaConfig() As Worksheet
    On Error Resume Next
    Set ObtenerHojaConfig = ThisWorkbook.Worksheets(NOMBRE_HOJA_CONFIG)
    
    If ObtenerHojaConfig Is Nothing Then
        Call RegistrarError("ObtenerHojaConfig", "No se encontro la hoja: " & NOMBRE_HOJA_CONFIG)
        MsgBox "Error: No se encontro la hoja de configuracion del sistema." & vbCrLf & _
               "Contacte al administrador.", vbCritical, "Error de Sistema"
    End If
    
    On Error GoTo 0
End Function
