'Attribute VB_Name = "Modulo_Logs"
Option Explicit

' === MODULO DE GESTION DE LOGS ===
' Este modulo centraliza el registro de eventos y errores de la aplicacion.
' Implementa un sistema de logs robusto con diferentes niveles y timestamps.

' Enum para niveles de log
Public Enum LogLevel
    LOG_ERROR = 1
    LOG_WARNING = 2
    LOG_INFO = 3
    LOG_DEBUG = 4
End Enum

' Configuracion del sistema de logs
Public Const LOGS_ACTIVOS As Boolean = True
Public Const NIVEL_LOG_MAXIMO As Integer = 3 ' LOG_INFO = 3, solo mostrar INFO, WARNING y ERROR

' Proposito: Registra un mensaje de log con timestamp y nivel
' Argumentos:
'   - strProcedimiento: Nombre del procedimiento que genera el log
'   - strMensaje: El mensaje a registrar
'   - intNivel: Nivel del log (opcional, por defecto INFO)
Public Sub RegistrarLog(ByVal strProcedimiento As String, ByVal strMensaje As String, _
                       Optional ByVal intNivel As LogLevel = LOG_INFO)
    Dim strTimestamp As String
    Dim strNivelTexto As String
    Dim strLogCompleto As String
    
    ' Verificar si los logs estan activos y el nivel es valido
    If Not LOGS_ACTIVOS Or intNivel > NIVEL_LOG_MAXIMO Then
        Exit Sub
    End If
    
    ' Determinar texto del nivel
    Select Case intNivel
        Case LOG_ERROR: strNivelTexto = "ERROR"
        Case LOG_WARNING: strNivelTexto = "WARNING"
        Case LOG_INFO: strNivelTexto = "INFO"
        Case LOG_DEBUG: strNivelTexto = "DEBUG"
        Case Else: strNivelTexto = "INFO"
    End Select
    
    ' Crear timestamp
    strTimestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Crear mensaje completo
    strLogCompleto = "[" & strTimestamp & "] [" & strNivelTexto & "] [" & strProcedimiento & "] - " & strMensaje
    
    ' Mostrar en ventana inmediato
    Debug.Print strLogCompleto
End Sub

' Proposito: Funcion de compatibilidad con version anterior (sin parametros)
' Argumentos:
'   - strMensaje: El mensaje a registrar
Public Sub RegistrarLogSimple(ByVal strMensaje As String)
    Call RegistrarLog("ProcedimientoNoEspecificado", strMensaje, LOG_INFO)
End Sub

' Proposito: Funcion especifica para registrar errores
' Argumentos:
'   - strProcedimiento: Nombre del procedimiento
'   - strMensaje: Mensaje de error
Public Sub RegistrarError(ByVal strProcedimiento As String, ByVal strMensaje As String)
    Call RegistrarLog(strProcedimiento, "ERROR: " & strMensaje, LOG_ERROR)
End Sub

' Proposito: Funcion especifica para registrar advertencias
' Argumentos:
'   - strProcedimiento: Nombre del procedimiento
'   - strMensaje: Mensaje de advertencia
Public Sub RegistrarAdvertencia(ByVal strProcedimiento As String, ByVal strMensaje As String)
    Call RegistrarLog(strProcedimiento, "ADVERTENCIA: " & strMensaje, LOG_WARNING)
End Sub

' Proposito: Funcion especifica para registrar informacion
' Argumentos:
'   - strProcedimiento: Nombre del procedimiento
'   - strMensaje: Mensaje informativo
Public Sub RegistrarInfo(ByVal strProcedimiento As String, ByVal strMensaje As String)
    Call RegistrarLog(strProcedimiento, strMensaje, LOG_INFO)
End Sub