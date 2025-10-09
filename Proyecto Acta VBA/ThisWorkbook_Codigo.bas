' =======================================================
' CODIGO PARA PEGAR EN EL MODULO "ThisWorkbook"
' =======================================================
' INSTRUCCIONES:
' 1. En Excel, abre el Editor de VBA (Alt + F11)
' 2. En el explorador de proyecto, busca "ThisWorkbook"
' 3. Haz doble clic en "ThisWorkbook" para abrir su editor
' 4. Copia y pega el codigo siguiente (desde Private Sub hasta End Sub)
' =======================================================

Option Explicit

' Evento que se ejecuta automaticamente al abrir el libro
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    
    Call RegistrarLog("Workbook_Open", "Libro abierto - Iniciando proceso de autenticacion", LOG_INFO)
    
    ' PASO 1: Proteger todo el libro inmediatamente
    Call ProtegerLibroCompleto
    
    ' PASO 2: Intentar autenticar usuario con opcion de cerrar si falla
    If Not AutenticarUsuario(cerrarSiFalla:=True) Then
        ' Si falla la autenticacion, se cerrara el libro automaticamente
        ' (esta linea solo se ejecuta si hay algun error en CerrarLibroSinGuardar)
        Call RegistrarError("Workbook_Open", "Fallo en autenticacion - El libro deberia haberse cerrado")
        Exit Sub
    End If
    
    ' PASO 3: Si la autenticacion fue exitosa, desproteger el libro
    Call DesprotegerLibroCompleto
    
    ' PASO 4: Mensaje de bienvenida
    MsgBox "Bienvenido " & UsuarioActual & "!" & vbCrLf & _
           "El libro esta ahora disponible para trabajar.", _
           vbInformation, "Acceso Concedido"
    
    Call RegistrarLog("Workbook_Open", "Acceso concedido a usuario: " & UsuarioActual, LOG_INFO)
    
    Exit Sub

ErrHandler:
    Call RegistrarError("Workbook_Open", Err.Description)
    MsgBox "Error al abrir el libro: " & Err.Description, vbCritical, "Error Critico"
    ' En caso de error critico, cerrar el libro
    Call CerrarLibroSinGuardar
End Sub

' Evento que se ejecuta antes de cerrar el libro
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    Call RegistrarLog("Workbook_BeforeClose", "Cerrando libro - Usuario: " & UsuarioActual, LOG_INFO)
    
    ' Limpiar variable de usuario actual
    UsuarioActual = ""
End Sub

' Evento que se ejecuta antes de guardar el libro
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    
    Call RegistrarLog("Workbook_BeforeSave", "Guardando libro - Usuario: " & UsuarioActual, LOG_INFO)
End Sub

