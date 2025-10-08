' VERSION 5.00
' Begin VB.UserForm FrmSelector 
'    Caption         =   "Selector de Hojas"
'    ClientHeight    =   3000
'    ClientLeft      =   60
'    ClientTop       =   345
'    ClientWidth     =   4800
'    StartUpPosition =   1  'Centrar
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Ir a Hoja16"
      Height          =   400
      Left            =   120
      Top             =   240
      Width           =   1200
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Ir a Hoja9"
      Height          =   400
      Left            =   1320
      Top             =   240
      Width           =   1200
   End
   Begin VB.CommandButton CommandButton3 
      Caption         =   "Ir a Hoja8"
      Height          =   400
      Left            =   2520
      Top             =   240
      Width           =   1200
   End
   Begin VB.CommandButton CommandButton4 
      Caption         =   "Ir a Hoja10"
      Height          =   400
      Left            =   3720
      Top             =   240
      Width           =   1200
   End
   Begin VB.ComboBox Seleccionar 
      Height          =   350
      Left            =   120
      Top             =   900
      Width           =   3000
   End
   Begin VB.CommandButton ir 
      Caption         =   "Ir"
      Height          =   350
      Left            =   3200
      Top             =   900
      Width           =   800
   End
' End
' Attribute VB_Name = "FrmSelector"
' Attribute VB_GlobalNameSpace = False
' Attribute VB_Creatable = False
' Attribute VB_PredeclaredId = True
' Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Hoja16.Select
End Sub

Private Sub CommandButton2_Click()
    Hoja9.Select
End Sub

Private Sub CommandButton3_Click()
    Hoja8.Select
End Sub

Private Sub CommandButton4_Click()
    Hoja10.Select
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim descripcion As String
    Me.Seleccionar.Clear
    For Each ws In ThisWorkbook.Worksheets
        nombreHoja = ws.Name
        If nombreHoja Like "*[0-9]*" Then
            On Error Resume Next
            descripcion = ws.Range("D7").Value
            On Error GoTo 0
            Me.Seleccionar.AddItem nombreHoja & " - " & descripcion
        End If
    Next ws
    Me.Seleccionar.ListIndex = -1
    ' Colores personalizados para los botones
    Me.CommandButton1.BackColor = RGB(226, 128, 15) ' E2800F
    Me.CommandButton2.BackColor = RGB(47, 103, 225) ' 2F67E1
    Me.CommandButton3.BackColor = RGB(35, 196, 23)  ' 23C417
    Me.CommandButton4.BackColor = RGB(233, 15, 15)  ' E90F0F
End Sub

Private Sub ir_Click()
    Dim seleccion As String
    Dim nombreHoja As String
    If Me.Seleccionar.ListIndex = -1 Then
        MsgBox "Por favor selecciona una hoja.", vbExclamation
        Exit Sub
    End If
    seleccion = Me.Seleccionar.Value
    nombreHoja = Trim(Split(seleccion, " - ")(0))
    On Error Resume Next
    ThisWorkbook.Sheets(nombreHoja).Activate
    On Error GoTo 0
    Unload Me
End Sub 