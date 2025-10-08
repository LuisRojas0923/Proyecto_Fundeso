' VERSION 5.00
' Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario_ 
'    Caption         =   "Seleccionar Fecha"
'    ClientHeight    =   3195
'    ClientLeft      =   120
'    ClientTop       =   465
'    ClientWidth     =   3360
'    OleObjectBlob   =   "frmCalendario_.frx":0000
'    StartUpPosition =   1  'Centrar en propietario
' End
' Attribute VB_Name = "frmCalendario_"
' Attribute VB_GlobalNameSpace = False
' Attribute VB_Creatable = False
' Attribute VB_PredeclaredId = True
' Attribute VB_Exposed = False

Option Explicit
Private mesActual As Integer, anioActual As Integer

Dim etiquetas(1 To 42) As clsEtiquetaFecha

' Variable para almacenar la fecha seleccionada
Public FechaSeleccionada As String



Private Sub UserForm_Initialize()
     Me.StartUpPosition = 1
      Call CentrarFormularioEnExcel(Me)
    mesActual = Month(Date): anioActual = Year(Date)
    MostrarCalendario mesActual, anioActual
End Sub

Private Sub btnAnterior_Click()
    mesActual = mesActual - 1: If mesActual = 0 Then mesActual = 12: anioActual = anioActual - 1
    MostrarCalendario mesActual, anioActual
End Sub

Private Sub btnSiguiente_Click()
    mesActual = mesActual + 1: If mesActual = 13 Then mesActual = 1: anioActual = anioActual + 1
    MostrarCalendario mesActual, anioActual
End Sub

Private Sub MostrarCalendario(mes As Integer, anio As Integer)
    Dim i As Integer, diaInicio As Integer, diasMes As Integer, fechaInicio As Date, diaSemana As Integer
    Dim ctrl As Control
    
    ' Limpiar etiquetas anteriores
    Call LogDebug("Limpiando etiquetas anteriores...")
    For i = 1 To 42
        Set etiquetas(i) = Nothing
    Next i
    
    fechaInicio = DateSerial(anio, mes, 1)
    diaInicio = Weekday(fechaInicio, vbSunday) - 1
    diasMes = Day(DateSerial(anio, mes + 1, 1) - 1)
    Me.lblMesAnio.Caption = UCase(Format(fechaInicio, "MMMM YYYY"))
    
    Call LogDebug("Mostrando calendario para: " & Format(fechaInicio, "MMMM YYYY"))

    For i = 1 To 42
        With Me.Controls("lblFecha" & i)
            .BackColor = RGB(255, 255, 255)
            .Font.Bold = False
            .TextAlign = fmTextAlignCenter

            If i > diaInicio And i <= diaInicio + diasMes Then
                .Caption = i - diaInicio
                .Enabled = True
                .BackColor = RGB(240, 240, 240)

                diaSemana = (i - 1) Mod 7 ' 0=Domingo, 6=S�bado
                If diaSemana = 0 Or diaSemana = 6 Then
                    .BackColor = RGB(255, 230, 230) ' Colorear fines de semana
                End If

                ' ?? Sombrar el d�a actual
                If Day(Date) = (i - diaInicio) And Month(Date) = mes And Year(Date) = anio Then
                    .BackColor = RGB(0, 32, 96) ' Azul claro
                    .Font.Bold = True
                    .ForeColor = RGB(255, 255, 255) ' Blanco
                End If
            Else
                .Caption = ""
                .Enabled = False
                .BackColor = RGB(255, 255, 255)
            End If

            ' Asignar clase con eventos
            On Error GoTo ErrorAsignacion
            Call LogDebug("Asignando clase a etiqueta " & i)
            Set etiquetas(i) = New clsEtiquetaFecha
            
            ' Intentar asignar el control con manejo de errores
            Call LogDebug("Buscando control: lblFecha" & i)
            Set ctrl = Me.Controls("lblFecha" & i)
            Call LogDebug("Control lblFecha" & i & " encontrado exitosamente")
            
            ' Verificar que el control es un Label
            Call LogDebug("Control encontrado: lblFecha" & i & ", Tipo: " & TypeName(ctrl))
            Call LogDebug("Nombre del control: " & ctrl.Name)
            Call LogDebug("Tipo exacto: " & TypeOf ctrl Is MSForms.Label)
            
            ' Intentar conversión a MSForms.Label
            If TypeOf ctrl Is MSForms.Label Then
                Call LogDebug("Control es MSForms.Label - Iniciando asignación a clase...")
                
                ' Paso 1: Asignar etiqueta
                Call LogDebug("Paso 1: Asignando ctrl a etiquetas(" & i & ").Etiqueta...")
                Set etiquetas(i).Etiqueta = ctrl
                Call LogDebug("Paso 1: Etiqueta asignada correctamente")
                
                ' Paso 2: Asignar mes
                Call LogDebug("Paso 2: Asignando mesActual = " & mes)
                etiquetas(i).mesActual = mes
                Call LogDebug("Paso 2: mesActual asignado correctamente")
                
                ' Paso 3: Asignar año
                Call LogDebug("Paso 3: Asignando anioActual = " & anio)
                etiquetas(i).anioActual = anio
                Call LogDebug("Paso 3: anioActual asignado correctamente")
                
                Call LogDebug("*** CLASE ASIGNADA EXITOSAMENTE A ETIQUETA " & i & " ***")
            Else
                Call LogWarn("*** CONTROL lblFecha" & i & " NO ES UN MSFORMS.LABEL ***")
                Call LogWarn("Tipo encontrado: " & TypeName(ctrl))
                Call LogWarn("Nombre del control: " & ctrl.Name)
                Call LogWarn("Es MSForms.Label: " & (TypeOf ctrl Is MSForms.Label))
                Set etiquetas(i) = Nothing
                Call LogDebug("Etiqueta " & i & " establecida como Nothing")
            End If
            On Error GoTo 0
        End With
    Next i
    Exit Sub
    
ErrorAsignacion:
    Call LogError("*** ERROR EN ASIGNACIÓN DE ETIQUETA ***")
    Call LogError("Etiqueta número: " & i)
    Call LogError("Descripción del error: " & Err.Description)
    Call LogError("Número de error: " & Err.Number)
    Call LogError("Fuente del error: " & Err.Source)
    
    ' Información adicional del estado
    Call LogDebug("Estado de etiquetas(" & i & "): " & IIf(etiquetas(i) Is Nothing, "Nothing", "Instanciado"))
    Call LogDebug("Mes actual: " & mes & ", Año actual: " & anio)
    
    ' Limpiar la etiqueta que falló
    Call LogDebug("Limpiando etiqueta " & i & " que falló...")
    Set etiquetas(i) = Nothing
    Call LogDebug("Etiqueta " & i & " limpiada")
    
    ' Continuar con la siguiente etiqueta
    Call LogDebug("Continuando con siguiente etiqueta...")
    Resume Next
End Sub

Private Sub lblFecha_Click()
    Call LogDebug("=== lblFecha_Click INICIADO ===")
    Dim i As Integer
    For i = 1 To 42
        If TypeName(Me.ActiveControl) = "Label" Then
            If Me.ActiveControl.Name = "lblFecha" & i Then
                Call LogDebug("Label clickeado: " & Me.ActiveControl.Name)
                If Me.Controls("lblFecha" & i).Caption <> "" Then
                    Dim fechaSeleccionada As Date
                    fechaSeleccionada = DateSerial(anioActual, mesActual, CInt(Me.Controls("lblFecha" & i).Caption))
                    Call LogDebug("Fecha calculada: " & Format(fechaSeleccionada, "dd/mm/yyyy"))
                    Call LogDebug("Llamando a AsignarFechaAControl...")
                    Call AsignarFechaAControl(fechaSeleccionada)
                    Exit Sub
                Else
                    Call LogDebug("Label vacío, no se puede seleccionar")
                End If
            End If
        End If
    Next i
    Call LogDebug("=== lblFecha_Click FINALIZADO ===")
End Sub


Private Sub lblFecha_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call lblFecha_Click
End Sub

' Función para asignar la fecha seleccionada al control correspondiente
' NOTA: Esta función ya no se usa porque clsEtiquetaFecha maneja directamente los clics
' Se mantiene por compatibilidad en caso de que se use desde lblFecha_Click
Private Sub AsignarFechaAControl(fechaSeleccionada As Date)
    On Error GoTo ErrorHandler
    
    Call LogDebug("=== ASIGNARFECHACONTROL INICIADO (MÉTODO LEGACY) ===")
    Call LogDebug("Fecha seleccionada: " & Format(fechaSeleccionada, "dd/mm/yyyy"))
    Call LogDebug("SelectedTextbox es Nothing: " & (SelectedTextbox Is Nothing))
    
    ' Almacenar la fecha en la variable pública del formulario
    FechaSeleccionada = Format(fechaSeleccionada, "dd/mm/yyyy")
    Call LogDebug("Fecha almacenada en FechaSeleccionada: " & FechaSeleccionada)
    
    ' Si hay un control seleccionado, asignar la fecha
    If Not SelectedTextbox Is Nothing Then
        SelectedTextbox.Value = Format(fechaSeleccionada, "dd/mm/yyyy")
        Call LogDebug("Fecha asignada al control: " & Format(fechaSeleccionada, "dd/mm/yyyy"))
        Call LogDebug("Valor del control después de asignar: " & SelectedTextbox.Value)
    Else
        Call LogWarn("SelectedTextbox es Nothing - usando variable FechaSeleccionada")
    End If
    
    ' Cerrar el formulario
    Call LogDebug("Cerrando formulario de calendario...")
    Unload Me
    
    Exit Sub
ErrorHandler:
    Call LogError("Error en AsignarFechaAControl: " & Err.Description)
    Unload Me
End Sub



