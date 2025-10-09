VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InicioSesion 
   Caption         =   "Inicio de Sesion"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   OleObjectBlob   =   "InicioSesion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "InicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' === FORMULARIO DE INICIO DE SESION ===
' Variable privada para almacenar el estado del login
Private m_LoginExitoso As Boolean

' Propiedad publica de solo lectura para acceder al estado del login desde fuera del formulario
Public Property Get LoginExitoso() As Boolean
    LoginExitoso = m_LoginExitoso
End Property

' Evento de inicializacion del formulario
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    Dim usuarios As Variant
    Dim i As Integer
    
    ' Inicializar la bandera de login
    m_LoginExitoso = False
    
    ' Centrar el formulario de inicio de sesion
    Call CentrarFormularioSimple(Me)
    
    ' Cargar usuarios desde la hoja Config_Sistema
    usuarios = CargarUsuariosEnComboBox()
    
    With Me.cmbUsuario
        .Clear
        ' Verificar si hay usuarios para cargar
        If IsArray(usuarios) Then
            For i = LBound(usuarios) To UBound(usuarios)
                .AddItem usuarios(i)
            Next i
        Else
            ' Si no hay usuarios, agregar uno por defecto (solo para emergencias)
            .AddItem "admin"
            Call RegistrarAdvertencia("UserForm_Initialize", "No se pudieron cargar usuarios, usando defaults")
        End If
    End With
    
    ' Configurar controles
    Me.txtPassword.PasswordChar = "*"
    
    Exit Sub

ErrHandler:
    Call RegistrarError("UserForm_Initialize", Err.Description)
    ' En caso de error, cargar usuarios por defecto
    With Me.cmbUsuario
        .Clear
        .AddItem "admin"
    End With
End Sub

' Evento de click en el boton Login
Private Sub cmdLogin_Click()
    On Error GoTo ErrHandler
    
    Dim strUsuario As String
    Dim strPassword As String
    
    ' Validar que se haya ingresado usuario y contrasena
    strUsuario = Trim(Me.cmbUsuario.Value)
    strPassword = Me.txtPassword.Value
    
    If strUsuario = "" Then
        MsgBox "Por favor seleccione un usuario", vbExclamation, "Usuario Requerido"
        Me.cmbUsuario.SetFocus
        Exit Sub
    End If
    
    If strPassword = "" Then
        MsgBox "Por favor ingrese la contrasena", vbExclamation, "Contrasena Requerida"
        Me.txtPassword.SetFocus
        Exit Sub
    End If
    
    ' Validar credenciales contra la hoja Config_Sistema
    If ValidarCredencialesDesdeHoja(strUsuario, strPassword) Then
        m_LoginExitoso = True
        Call RegistrarLog("cmdLogin_Click", "Login exitoso para: " & strUsuario, LOG_INFO)
        Me.Hide
    Else
        m_LoginExitoso = False
        Call RegistrarAdvertencia("cmdLogin_Click", "Intento de login fallido para: " & strUsuario)
        MsgBox "Usuario o contrasena incorrectos", vbCritical, "Acceso Denegado"
        Me.txtPassword.Value = ""
        Me.txtPassword.SetFocus
    End If
    
    Exit Sub

ErrHandler:
    Call RegistrarError("cmdLogin_Click", Err.Description)
    MsgBox "Error durante el proceso de login: " & Err.Description, vbCritical, "Error"
    m_LoginExitoso = False
End Sub

' Evento de click en el boton Cancelar
Private Sub cmdCancelar_Click()
    m_LoginExitoso = False
    Call RegistrarLog("cmdCancelar_Click", "Login cancelado por el usuario", LOG_INFO)
    Me.Hide
End Sub

' Evento de click en el checkbox para mostrar/ocultar contrasena
Private Sub chkMostrar_Click()
    If chkMostrar.Value = True Then
        txtPassword.PasswordChar = ""    ' Mostrar
    Else
        txtPassword.PasswordChar = "*"   ' Ocultar
    End If
End Sub

' Evento de click en el boton de Configuracion
' NOTA: Este boton debe agregarse manualmente al formulario en el editor de VBA
' Nombre del boton: cmdConfiguracion
' Caption: "Configuracion"
Private Sub cmdConfiguracion_Click()
    On Error GoTo ErrHandler
    
    ' Cerrar el formulario de login temporalmente
    Me.Hide
    
    ' Verificar si el usuario ya inicio sesion
    If m_LoginExitoso Then
        ' Si ya hay un usuario logueado, abrir configuracion
        Call AbrirHojaConfiguracion
    Else
        ' Si no hay usuario logueado, pedir autenticacion primero
        MsgBox "Debe iniciar sesion primero para acceder a la configuracion.", vbExclamation, "Autenticacion Requerida"
        ' Mostrar nuevamente el formulario
        Me.Show
    End If
    
    Exit Sub

ErrHandler:
    Call RegistrarError("cmdConfiguracion_Click", Err.Description)
    MsgBox "Error al abrir configuracion: " & Err.Description, vbCritical, "Error"
End Sub
