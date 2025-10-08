Attribute VB_Name = "Modulo_Seguridad"
Option Explicit

' Este módulo centraliza las funciones relacionadas con la seguridad y autenticación de usuarios.

' Muestra el formulario de inicio de sesión y devuelve True si el login fue exitoso.
' Devuelve False si el usuario cancela o las credenciales son incorrectas.
Public Function AutenticarUsuario() As Boolean
    On Error GoTo ErrHandler
    
    ' Por defecto, la autenticación falla hasta que se demuestre lo contrario.
    AutenticarUsuario = False
    
    Debug.Print "Iniciando proceso de autenticacion..."
    
    ' Mostrar el formulario de inicio de sesión de forma modal.
    ' El código se detiene aquí hasta que el formulario se cierre.
    InicioSesion.Show vbModal
    
    ' Después de que el formulario se cierra, verificamos si el login fue exitoso.
    ' Se asume que el formulario "InicioSesion" tiene una variable pública
    ' llamada "LoginExitoso" que se establece en True si las credenciales son correctas.
    If InicioSesion.LoginExitoso Then
        Debug.Print "Autenticacion exitosa."
        AutenticarUsuario = True
    Else
        Debug.Print "Autenticacion fallida o cancelada por el usuario."
    End If
    
    ' Descargar el formulario de la memoria para limpiar su estado.
    Unload InicioSesion
    
    Exit Function

ErrHandler:
    Debug.Print "Error en AutenticarUsuario: " & Err.Description
    ' Asegurarse de descargar el formulario en caso de error.
    On Error Resume Next
    Unload InicioSesion
    On Error GoTo 0
    AutenticarUsuario = False
End Function
