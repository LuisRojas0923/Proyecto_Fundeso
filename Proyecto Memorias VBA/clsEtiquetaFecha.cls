' VERSION 1.0 CLASS
' BEGIN
'   MultiUse = -1  'True
' END
' Attribute VB_Name = "clsEtiquetaFecha"
' Attribute VB_GlobalNameSpace = False
' Attribute VB_Creatable = False
' Attribute VB_PredeclaredId = False
' Attribute VB_Exposed = False
Public WithEvents Etiqueta As MSForms.Label
' Attribute Etiqueta.VB_VarHelpID = -1
Private mes As Integer, anio As Integer

' Evento que se dispara cuando se asigna la etiqueta
Private Sub Class_Initialize()
    Call LogDebug("clsEtiquetaFecha: Nueva instancia creada")
End Sub

Private Sub Class_Terminate()
    Call LogDebug("clsEtiquetaFecha: Instancia destruida")
End Sub

Public Property Let mesActual(valor As Integer)
    mes = valor
    Call LogDebug("clsEtiquetaFecha: mesActual establecido a " & valor)
End Property

Public Property Let anioActual(valor As Integer)
    anio = valor
    Call LogDebug("clsEtiquetaFecha: anioActual establecido a " & valor)
End Property

Private Sub Etiqueta_Click()
    If Etiqueta.Caption <> "" Then
        Call LogDebug("=== clsEtiquetaFecha Click INICIADO ===")
        Call LogDebug("Etiqueta clickeada: " & Etiqueta.Name & ", Valor: " & Etiqueta.Caption)
        
        Dim fechaSeleccionada As Date
        fechaSeleccionada = DateSerial(anio, mes, CInt(Etiqueta.Caption))
        Call LogDebug("Fecha calculada en clase: " & Format(fechaSeleccionada, "dd/mm/yyyy"))
        
        ' Asignar la fecha a la variable pública del formulario (para creación rápida)
        frmCalendario_.FechaSeleccionada = Format(fechaSeleccionada, "dd/mm/yyyy")
        Call LogDebug("Fecha asignada a frmCalendario_.FechaSeleccionada: " & frmCalendario_.FechaSeleccionada)
        
        ' Si hay un control seleccionado (para campos F_Desde, F_Hasta), asignar también ahí
        If Not SelectedTextbox Is Nothing Then
            SelectedTextbox.Value = Format(fechaSeleccionada, "dd/mm/yyyy")
            Call LogDebug("Fecha asignada al control SelectedTextbox: " & SelectedTextbox.Value)
        Else
            Call LogDebug("SelectedTextbox es Nothing - solo usando FechaSeleccionada")
        End If
        
        ' Cerrar el formulario
        Call LogDebug("Cerrando formulario desde clsEtiquetaFecha...")
        Unload frmCalendario_
        
        Call LogDebug("=== clsEtiquetaFecha Click FINALIZADO ===")
    Else
        Call LogDebug("Etiqueta vacía en clsEtiquetaFecha, no se puede seleccionar")
    End If
End Sub

