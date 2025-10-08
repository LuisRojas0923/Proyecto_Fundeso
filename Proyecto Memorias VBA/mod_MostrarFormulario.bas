' Attribute VB_Name = "mod_MostrarFormulario"
' ?? Funci�n reutilizable para centrar cualquier UserForm en la pantalla donde est� Excel

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

Public Sub CentrarFormularioEnExcel(ByRef formulario As Object)
    Dim hWnd As LongPtr
    Dim rect As rect
    Dim hMonitor As LongPtr
    Dim mi As MONITORINFO
    Dim LeftPos As Single, TopPos As Single
    Dim AnchoMonitor As Long, AltoMonitor As Long

    Call LogInfo("Buscando ventana de Excel...")
    hWnd = FindWindow("XLMAIN", Application.Caption)
    If hWnd = 0 Then
        MsgBox "No se pudo encontrar la ventana de Excel", vbExclamation
        Call LogError("No se encontró la ventana de Excel.")
        Exit Sub
    End If
    Call LogOK("Ventana Excel encontrada. Handle: " & hWnd)

    GetWindowRect hWnd, rect

    ' Obtener el monitor donde est� la ventana de Excel
    hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
    mi.cbSize = Len(mi)
    If GetMonitorInfo(hMonitor, mi) = 0 Then
        MsgBox "No se pudo obtener la informaci�n del monitor", vbExclamation
        Exit Sub
    End If

    ' Dimensiones del monitor
    AnchoMonitor = mi.rcWork.Right - mi.rcWork.Left
    AltoMonitor = mi.rcWork.Bottom - mi.rcWork.Top

    With formulario
        .StartUpPosition = 0
        LeftPos = mi.rcWork.Left + (AnchoMonitor - .Width) / 2
        TopPos = mi.rcWork.Top + (AltoMonitor - .Height) / 2.5

        .Left = LeftPos
        .Top = TopPos

            Call LogOK("Formulario centrado en el monitor de Excel.")
    End With
End Sub





Sub MostrarFormulario()
frm_MultiEmpleados.Show
End Sub

