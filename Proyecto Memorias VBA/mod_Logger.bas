' Attribute VB_Name = "mod_Logger"
' === MODULO DE LOGGING CENTRALIZADO ===
' Este modulo centraliza todos los mensajes de depuracion
' Para desactivar todos los logs, simplemente comenta la linea correspondiente en cada funcion

' Variable global para controlar si el logging esta activo
Public Const LOGGING_ACTIVO As Boolean = True

' Funcion principal de logging
Public Sub LogDebug(mensaje As String)
    If LOGGING_ACTIVO Then
        Debug.Print mensaje
    End If
End Sub

' Funcion de logging con prefijo de informacion
Public Sub LogInfo(mensaje As String)
    If LOGGING_ACTIVO Then
        Debug.Print "[INFO] " & mensaje
    End If
End Sub

' Funcion de logging con prefijo de advertencia
Public Sub LogWarn(mensaje As String)
    If LOGGING_ACTIVO Then
        Debug.Print "[WARN] " & mensaje
    End If
End Sub

' Funcion de logging con prefijo de error
Public Sub LogError(mensaje As String)
    If LOGGING_ACTIVO Then
        Debug.Print "[ERROR] " & mensaje
    End If
End Sub

' Funcion de logging con prefijo de exito
Public Sub LogOK(mensaje As String)
    If LOGGING_ACTIVO Then
        Debug.Print "[OK] " & mensaje
    End If
End Sub

' Funcion de logging para errores de VBA con descripcion completa
Public Sub LogErrorVBA(procedimiento As String, descripcionError As String, Optional numeroLinea As Long = 0)
    If LOGGING_ACTIVO Then
        Dim mensaje As String
        mensaje = "[ERROR] " & procedimiento & ": " & descripcionError
        If numeroLinea > 0 Then
            mensaje = mensaje & " (Linea: " & numeroLinea & ")"
        End If
        Debug.Print mensaje
    End If
End Sub

' Funcion para logging de inicio de procedimientos
Public Sub LogIniciar(procedimiento As String)
    If LOGGING_ACTIVO Then
        Debug.Print "[INICIO] " & procedimiento
    End If
End Sub

' Funcion para logging de finalizacion de procedimientos
Public Sub LogFinalizar(procedimiento As String)
    If LOGGING_ACTIVO Then
        Debug.Print "[FIN] " & procedimiento
    End If
End Sub
