Attribute VB_Name = "Utilidades"

' === MÓDULO DE UTILIDADES PARA FORMULARIOS ===

' Declaraciones de API de Windows para centrado avanzado
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetWindowRect Lib "user32" _
    (ByVal hWnd As LongPtr, lpRect As rect) As Long

Private Declare PtrSafe Function MonitorFromWindow Lib "user32" _
    (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" _
    (ByVal hMonitor As LongPtr, ByRef lpmi As MONITORINFO) As Long

Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As rect
    rcWork As rect
    dwFlags As Long
End Type

Const MONITOR_DEFAULTTONEAREST = &H2

' Función reutilizable para centrar cualquier UserForm en la pantalla donde está Excel
Public Sub CentrarFormularioEnExcel(ByRef formulario As Object)
    On Error GoTo ErrHandler
    
    Dim hWnd As LongPtr
    Dim rect As rect
    Dim hMonitor As LongPtr
    Dim mi As MONITORINFO
    Dim LeftPos As Single, TopPos As Single
    Dim AnchoMonitor As Long, AltoMonitor As Long

    Debug.Print "[INFO] Buscando ventana de Excel..."
    hWnd = FindWindow("XLMAIN", Application.Caption)
    If hWnd = 0 Then
        MsgBox "No se pudo encontrar la ventana de Excel", vbExclamation
        Debug.Print "[ERROR] No se encontró la ventana de Excel."
        Exit Sub
    End If
    Debug.Print "[OK] Ventana Excel encontrada. Handle: " & hWnd

    GetWindowRect hWnd, rect

    ' Obtener el monitor donde está la ventana de Excel
    hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
    mi.cbSize = Len(mi)
    If GetMonitorInfo(hMonitor, mi) = 0 Then
        MsgBox "No se pudo obtener la información del monitor", vbExclamation
        Exit Sub
    End If

    ' Dimensiones del monitor (usar rcMonitor para centrar en todo el monitor)
    AnchoMonitor = mi.rcMonitor.Right - mi.rcMonitor.Left
    AltoMonitor = mi.rcMonitor.Bottom - mi.rcMonitor.Top

    With formulario
        .StartUpPosition = 0
        LeftPos = mi.rcMonitor.Left + (AnchoMonitor - .Width) / 2
        TopPos = mi.rcMonitor.Top + (AltoMonitor - .Height) / 2

        .Left = LeftPos
        .Top = TopPos

        Debug.Print "[OK] Formulario centrado en el monitor de Excel."
        Debug.Print "[INFO] Posición: Left=" & .Left & ", Top=" & .Top
        Debug.Print "[INFO] Dimensiones monitor: " & AnchoMonitor & "x" & AltoMonitor
        Debug.Print "[INFO] Dimensiones formulario: " & .Width & "x" & .Height
    End With
    
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] Error en CentrarFormularioEnExcel: " & Err.Description
    ' Fallback: centrado básico
    With formulario
        .StartUpPosition = 0
        .Left = (Application.Width - .Width) / 2 + Application.Left
        .Top = (Application.Height - .Height) / 2 + Application.Top
    End With
End Sub

' Función para validar que Excel esté activo
Public Function ExcelEstaActivo() As Boolean
    On Error GoTo ErrHandler
    
    Dim hWnd As LongPtr
    hWnd = FindWindow("XLMAIN", Application.Caption)
    ExcelEstaActivo = (hWnd <> 0)
    
    Exit Function
ErrHandler:
    ExcelEstaActivo = False
End Function

' Función alternativa simple para centrar formularios
Public Sub CentrarFormularioSimple(ByRef formulario As Object)
    On Error GoTo ErrHandler
    
    Dim screenWidth As Long, screenHeight As Long
    Dim formWidth As Long, formHeight As Long
    Dim leftPos As Long, topPos As Long
    
    ' Obtener dimensiones de la pantalla
    screenWidth = Application.Width
    screenHeight = Application.Height
    
    ' Obtener dimensiones del formulario
    formWidth = formulario.Width
    formHeight = formulario.Height
    
    ' Calcular posición centrada
    leftPos = (screenWidth - formWidth) / 2
    topPos = (screenHeight - formHeight) / 2
    
    ' Aplicar posición
    With formulario
        .StartUpPosition = 0
        .Left = leftPos
        .Top = topPos
    End With
    
    Debug.Print "[SIMPLE] Formulario centrado: Left=" & leftPos & ", Top=" & topPos
    Debug.Print "[SIMPLE] Pantalla: " & screenWidth & "x" & screenHeight
    Debug.Print "[SIMPLE] Formulario: " & formWidth & "x" & formHeight
    
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] Error en CentrarFormularioSimple: " & Err.Description
    ' Fallback básico
    With formulario
        .StartUpPosition = 1 ' Centrar en propietario
    End With
End Sub

' Función para obtener información del monitor actual
Public Sub ObtenerInfoMonitor()
    On Error GoTo ErrHandler
    
    Dim hWnd As LongPtr
    Dim rect As rect
    Dim hMonitor As LongPtr
    Dim mi As MONITORINFO
    
    hWnd = FindWindow("XLMAIN", Application.Caption)
    If hWnd = 0 Then
        Debug.Print "[ERROR] No se encontró la ventana de Excel."
        Exit Sub
    End If
    
    GetWindowRect hWnd, rect
    hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
    mi.cbSize = Len(mi)
    
    If GetMonitorInfo(hMonitor, mi) <> 0 Then
        Debug.Print "[INFO] Monitor actual:"
        Debug.Print "  - Ancho: " & (mi.rcWork.Right - mi.rcWork.Left)
        Debug.Print "  - Alto: " & (mi.rcWork.Bottom - mi.rcWork.Top)
        Debug.Print "  - Posición: " & mi.rcWork.Left & "," & mi.rcWork.Top
    End If
    
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] Error obteniendo información del monitor: " & Err.Description
End Sub

' === MÓDULO DE UTILIDADES DE RED ===

' Declaracion de API para verificar el estado de la conexion a Internet
Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Boolean

' Funcion para verificar si hay una conexion a Internet activa.
' Devuelve: Verdadero si hay conexion, Falso si no.
Public Function HayConexionInternet() As Boolean
    On Error GoTo ErrHandler
    HayConexionInternet = InternetGetConnectedState(0&, 0&)
    If HayConexionInternet Then
        Debug.Print "[RED] Conexion a Internet detectada."
    Else
        Debug.Print "[RED] No hay conexion a Internet."
    End If
    Exit Function
ErrHandler:
    Debug.Print "[ERROR] Error en HayConexionInternet: " & Err.Description
    HayConexionInternet = False
End Function

' Funcion para medir la latencia (ping) a un host.
' Devuelve: Latencia en ms. -1 si hay error o no se encuentra el host.
Public Function ObtenerLatencia(Optional ByVal host As String = "8.8.8.8") As Long
    On Error GoTo ErrHandler
    
    Dim objShell As Object
    Dim objExec As Object
    Dim strResult As String
    Dim arrLines() As String
    Dim line As Variant
    Dim timePart As String
    
    Set objShell = CreateObject("WScript.Shell")
    ' Ejecuta el comando ping con 1 intento y un timeout de 1000ms (1 segundo)
    Set objExec = objShell.Exec("ping -n 1 -w 1000 " & host)
    
    ' Espera a que el comando termine y lee el resultado
    strResult = LCase(objExec.StdOut.ReadAll)
    Debug.Print "[PING] Resultado para " & host & ": " & vbCrLf & strResult
    
    ' Valor por defecto si no se encuentra la latencia
    ObtenerLatencia = -1
    
    arrLines = Split(strResult, vbLf)
    For Each line In arrLines
        ' Busca una linea que contenga "time=" o "tiempo=" para compatibilidad de idioma
        If InStr(line, "time=") > 0 Or InStr(line, "tiempo=") > 0 Then
            ' Extrae la parte del tiempo
            If InStr(line, "time=") > 0 Then
                timePart = Mid(line, InStr(line, "time=") + 5)
            Else
                timePart = Mid(line, InStr(line, "tiempo=") + 7)
            End If
            
            ' Limpia y convierte a numero
            timePart = Split(timePart, "ms")(0)
            timePart = Trim(Replace(timePart, "<", ""))
            If IsNumeric(timePart) Then
                ObtenerLatencia = CLng(timePart)
                Debug.Print "[PING] Latencia extraida: " & ObtenerLatencia & "ms"
                Exit For ' Sale del bucle una vez encontrada
            End If
        End If
    Next line

    Set objShell = Nothing
    Set objExec = Nothing
    Exit Function
    
ErrHandler:
    Debug.Print "[ERROR] Error en ObtenerLatencia: " & Err.Description
    ObtenerLatencia = -1
    Set objShell = Nothing
    Set objExec = Nothing
End Function
