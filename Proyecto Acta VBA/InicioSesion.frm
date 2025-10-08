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

' Variable privada para almacenar el estado del login.
Private m_LoginExitoso As Boolean

' Propiedad publica de solo lectura para acceder al estado del login desde fuera del formulario.
Public Property Get LoginExitoso() As Boolean
    LoginExitoso = m_LoginExitoso
End Property

Private Sub chkMostrar_Click()
    If chkMostrar.Value = True Then
        txtPassword.PasswordChar = ""    ' Mostrar
    Else
        txtPassword.PasswordChar = "*"   ' Ocultar
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Inicializar la bandera de login
    m_LoginExitoso = False
    
    ' Centrar el formulario de inicio de sesion
    Call CentrarFormulario(Me)
    
    With Me.cmbUsuario
        .Clear
        .AddItem "admin"
        .AddItem "usuario1"
        .AddItem "usuario2"
    End With
End Sub

Private Sub cmdLogin_Click()
    If Me.cmbUsuario.Value = "admin" And Me.txtPassword.Value = "1234" Then
        m_LoginExitoso = True
        Me.Hide
    Else
        m_LoginExitoso = False
        MsgBox "Usuario o contrasena incorrectos", vbCritical
        Me.txtPassword.Value = ""
        Me.txtPassword.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    m_LoginExitoso = False
    Me.Hide
End Sub
