' VERSION 5.00
' Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNavegadorHojas 
'    Caption         =   "Menu Memorias"
'    ClientHeight    =   2115
'    ClientLeft      =   120
'    ClientTop       =   465
'    ClientWidth     =   4770
'    OleObjectBlob   =   "FrmNavegadorHoja.frx":0000
'    StartUpPosition =   1  'Centrar en propietario
' End
' Attribute VB_Name = "frmNavegadorHojas"
' Attribute VB_GlobalNameSpace = False
' Attribute VB_Creatable = False
' Attribute VB_PredeclaredId = True
' Attribute VB_Exposed = False
' C�digo del formulario
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim descripcion As String
    Me.Seleccionar.Clear
    For Each ws In ThisWorkbook.Worksheets
        nombreHoja = ws.Name
        ' Verifica si el nombre contiene al menos un n�mero
        If nombreHoja Like "*[0-9]*" Then
            On Error Resume Next
            descripcion = ws.Range("D7").Value
            On Error GoTo 0
            Me.Seleccionar.AddItem nombreHoja & " - " & descripcion
        End If
    Next ws
    Me.Seleccionar.ListIndex = -1
End Sub

Private Sub ir_Click()
    Dim seleccion As String
    Dim nombreHoja As String
    If Me.Seleccionar.ListIndex = -1 Then
        MsgBox "Por favor selecciona una hoja.", vbExclamation
        Exit Sub
    End If
    seleccion = Me.Seleccionar.Value
    ' Extrae el nombre de la hoja antes del primer " - "
    nombreHoja = Trim(Split(seleccion, " - ")(0))
    On Error Resume Next
    ThisWorkbook.Sheets(nombreHoja).Activate
    On Error GoTo 0
    Unload Me
End Sub
