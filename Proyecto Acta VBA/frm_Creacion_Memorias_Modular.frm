VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Creacion_Memorias 
   Caption         =   "Creacion Presupuesto"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18210
   OleObjectBlob   =   "frm_Creacion_Memorias_Modular.frx":0000
   StartUpPosition =   3  'Predeterminado de Widnows
End
Attribute VB_Name = "frm_Creacion_Memorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Creacion_Memorias
'   Caption = "Creacion De Memorias"
'   ClientHeight = 10245
'   ClientLeft = 120
'   ClientTop = 465
'   ClientWidth = 17370
'   OleObjectBlob   =   "frm_Creacion_Memorias.frx":0000
'   StartUpPosition = 3    'Predeterminado de Widnows
'End
'Attribute VB_Name = "frm_Creacion_Memorias"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False

' === FORMULARIO PRINCIPAL MODULAR ===
' Este formulario usa los m�dulos: Modulo_ComboBoxes, Modulo_ListBoxes, Modulo_Consecutivos, Modulo_Trabajo, Modulo_Exportacion

Private Sub btn_LimpiarCampos_Click()
    On Error GoTo ErrHandler
    
    ' Limpiar todos los ListBoxes del nuevo sistema
    Call LimpiarTodosLosListBoxes(Me)
    
    ' Limpiar controles de la p�gina 1
    Me.Palabra_Clave.Value = ""
    Me.cmb_Area.Value = ""
    Me.cmb_Capitulos.Value = ""
    
    ' Recargar datos en la p�gina 1
    Call FiltrarYCargarListBox(Me)
    
    ' Cambiar a la p�gina 1
    Me.MultiPage1.Value = 0
    
    MsgBox "Todos los campos han sido limpiados. El sistema est� listo para una nueva sesi�n.", vbInformation, "Campos Limpiados"
    Debug.Print "Sistema limpiado completamente"
    Exit Sub
ErrHandler:
    Debug.Print "Error en btn_LimpiarCampos_Click: " & Err.Description
End Sub

Private Sub btn_Desmarcar_Click()
    On Error GoTo ErrHandler
    Dim item As Long, i As Long
    Titulo = "Fundeso"
    item = 0
    With Me.Listbox_Registros
        If .ListCount > 0 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    .Selected(i) = False
                    item = item + 1
                End If
            Next i
            If item = 0 Then
                MsgBox "No hay registros seleccionados", vbExclamation, Titulo
                Exit Sub
            End If
        Else
            MsgBox "No hay registros para deseleccionar", vbExclamation, Titulo
        End If
    End With
    Exit Sub
ErrHandler:
    Debug.Print "Error en btn_Desmarcar_Click: " & Err.Description
End Sub

Private Sub btn_Marcar_Click()
    On Error GoTo ErrHandler
    Dim i As Long
    Titulo = "Fundeso"
    If Me.Listbox_Registros.ListCount > 0 Then
        For i = 0 To Me.Listbox_Registros.ListCount - 1
            Me.Listbox_Registros.Selected(i) = True
        Next i
    Else
        MsgBox "No hay registros para seleccionar", vbExclamation, Titulo
    End If
    Exit Sub
ErrHandler:
    Debug.Print "Error en btn_Marcar_Click: " & Err.Description
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    ' === CENTRAR FORMULARIO ===
    Call CentrarFormularioSimple(Me)
    
    ' === VALIDAR CONTROLES DEL FORMULARIO ===
    Call ValidarControlesFormulario(Me)
    
    ' === CONFIGURAR ORDEN DE NAVEGACION POR TABS ===
    Call ConfigurarOrdenTabs(Me)
    
    ' Cargar el ComboBox de �reas
    Call CargarComboBoxAreas(Me)
    
    ' Cargar el ComboBox de Cap�tulos
    Call CargarComboBoxCapitulos(Me)
    
    ' Configurar ListBoxes del nuevo sistema MultiPage
    Call ConfigurarListBoxTrabajo_Principal(Me)
    Call ConfigurarListBoxExportados(Me)
    
    ' Cargar datos iniciales en la p�gina 1
    Me.Listbox_Registros.Clear
    Call AjustarAnchoListBox(Me)
    Call FiltrarYCargarListBox(Me)
    
    Exit Sub
ErrHandler:
    Debug.Print "ERROR en UserForm_Initialize: " & Err.Description
    Debug.Print "N�mero de error: " & Err.Number
End Sub

Private Sub Palabra_Clave_Change()
    Call FiltrarYCargarListBox(Me)
    ' Eliminar llamada a AjustarAnchoListBox aqu� para que el ancho no cambie al cargar datos
End Sub

' Funci�n para agregar registros seleccionados al ListBox de Trabajo
Private Sub btn_AgregarATrabajo_Click()
    On Error GoTo ErrHandler
    
    Dim registrosSeleccionados As Long
    Dim areaSeleccionada As String, capituloSeleccionado As String
    Dim cantidadIngresada As String
    Dim i As Long
    
    ' Validar que se haya seleccionado un �rea y cap�tulo
    areaSeleccionada = Me.cmb_Area.Value
    capituloSeleccionado = Me.cmb_Capitulos.Value
    
    If areaSeleccionada = "" Then
        MsgBox "Debe seleccionar un �rea antes de agregar registros", vbExclamation, "Validaci�n"
        Exit Sub
    End If
    
    If capituloSeleccionado = "" Then
        MsgBox "Debe seleccionar un Cap�tulo antes de agregar registros", vbExclamation, "Validaci�n"
        Exit Sub
    End If
    
    ' Contar registros seleccionados
    registrosSeleccionados = 0
    With Me.Listbox_Registros
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                registrosSeleccionados = registrosSeleccionados + 1
            End If
        Next i
    End With
    
    If registrosSeleccionados = 0 Then
        MsgBox "Debe seleccionar al menos un registro para agregar", vbExclamation, "Validaci�n"
        Exit Sub
    End If
    
    ' Obtener cantidad del usuario (cantidad general para todos los registros seleccionados)
    cantidadIngresada = InputBox("Ingrese la cantidad para los " & registrosSeleccionados & " registros seleccionados:" & vbCrLf & vbCrLf & "Esta cantidad se aplicar� a todos los registros seleccionados.", "Cantidad General", "1")
    If cantidadIngresada = "" Then Exit Sub
    
    If Not IsNumeric(cantidadIngresada) Or CDbl(cantidadIngresada) <= 0 Then
        MsgBox "La cantidad debe ser un n�mero mayor a 0", vbExclamation, "Validaci�n"
        Exit Sub
    End If
    
    ' Agregar registros al ListBox de Trabajo usando el m�dulo
    Call AgregarRegistrosATrabajo(Me, areaSeleccionada, capituloSeleccionado, CDbl(cantidadIngresada))
    
    ' Cambiar a la p�gina de trabajo
    Me.MultiPage1.Value = 1
    
    ' Limpiar selecciones en ListBox principal
    With Me.Listbox_Registros
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
    End With
    
    Debug.Print "Registros agregados al trabajo: " & registrosSeleccionados & " con cantidad: " & cantidadIngresada
    Exit Sub
ErrHandler:
    Debug.Print "Error en btn_AgregarATrabajo_Click: " & Err.Description
End Sub

' === NUEVA FUNCI�N: Asignar cantidad individual en P�gina 2 ===
Private Sub btn_AsignarCantidad_Click()
    On Error GoTo ErrHandler
    
    Dim cantidadIngresada As String
    Dim cantidadNumerica As Double
    
    ' Obtener cantidad del txt_Cantidad o InputBox
    If Me.txt_Cantidad.Value <> "" Then
        cantidadIngresada = Me.txt_Cantidad.Value
    Else
        cantidadIngresada = InputBox("Ingrese la cantidad para los registros seleccionados:", "Cantidad", "1")
        If cantidadIngresada = "" Then Exit Sub
    End If
    
    ' Validar cantidad
    If Not IsNumeric(cantidadIngresada) Or CDbl(cantidadIngresada) <= 0 Then
        MsgBox "La cantidad debe ser un n�mero mayor a 0", vbExclamation, "Validaci�n"
        Exit Sub
    End If
    
    cantidadNumerica = CDbl(cantidadIngresada)
    
    ' Asignar cantidad usando el m�dulo
    Call AsignarCantidadATrabajo(Me, cantidadNumerica)
    
    ' Limpiar campo de cantidad
    Me.txt_Cantidad.Value = ""
    
    Exit Sub
ErrHandler:
    Debug.Print "Error en btn_AsignarCantidad_Click: " & Err.Description
    MsgBox "Error al asignar cantidad: " & Err.Description, vbCritical, "Error"
End Sub

' Exportar registros del ListBox de Trabajo a la hoja
Private Sub btn_Exportar_Click()
    On Error GoTo ErrHandler
    
    ' Validar que hay registros para exportar
    If Me.Listbox_Trabajo.ListCount = 0 Then
        MsgBox "No hay registros en el �rea de trabajo para exportar", vbExclamation, "Validaci�n"
        Exit Sub
    End If
    
    ' Exportar usando el m�dulo
    Call ExportarDatosATrabajo(Me)
    
    ' Cambiar a la p�gina de revisi�n
    Me.MultiPage1.Value = 2
    
    ' Cargar datos exportados en la p�gina de revisi�n
    Call CargarDatosExportadosEnRevision(Me)
    
    Exit Sub
ErrHandler:
    Debug.Print "Error en btn_Exportar_Click: " & Err.Description
    MsgBox "Error durante la exportaci�n: " & Err.Description, vbCritical, "Error"
End Sub

' Eliminar registros seleccionados del ListBox de Trabajo
Private Sub btn_EliminarSeleccionado_Click()
    On Error GoTo ErrHandler
    
    ' Eliminar registros usando el m�dulo
    Call EliminarRegistrosTrabajo(Me)
    
    Exit Sub
ErrHandler:
    Debug.Print "Error en btn_EliminarSeleccionado_Click: " & Err.Description
    MsgBox "Error durante la eliminaci�n: " & Err.Description, vbCritical, "Error"
End Sub

' Evento cuando se cambia a la p�gina de revisi�n
Private Sub MultiPage1_Change()
    On Error GoTo ErrHandler
    
    ' Si se cambia a la p�gina 3 (Revisi�n), cargar datos autom�ticamente
    If Me.MultiPage1.Value = 2 Then ' P�gina 3 (�ndice 2)
        Call CargarDatosExportadosEnRevision(Me)
        Call MostrarEstadisticasExportados(Me)
        Call MostrarResumenConsecutivos ' Mostrar resumen de consecutivos
        Debug.Print "Cambio a p�gina de revisi�n - datos cargados autom�ticamente"
    End If
    
    Exit Sub
ErrHandler:
    Debug.Print "Error en MultiPage1_Change: " & Err.Description
End Sub

' Evento de doble clic en ListBox de Trabajo para editar cantidad
Private Sub Listbox_Trabajo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrHandler
    
    Dim filaSeleccionada As Long
    filaSeleccionada = Me.Listbox_Trabajo.ListIndex
    
    If filaSeleccionada = -1 Then Exit Sub ' No hay nada seleccionado
    
    Dim nuevaCantidadStr As String
    nuevaCantidadStr = InputBox("Ingrese la nueva cantidad para el registro seleccionado:", "Editar Cantidad", Me.Listbox_Trabajo.List(filaSeleccionada, 8))
    
    If nuevaCantidadStr = "" Then Exit Sub ' Usuario cancelo
    
    If IsNumeric(nuevaCantidadStr) Then
        Call EditarCantidadTrabajo(Me, filaSeleccionada, CDbl(nuevaCantidadStr))
    Else
        MsgBox "Por favor, ingrese un valor numerico valido.", vbExclamation, "Entrada no valida"
    End If
    
    Exit Sub
ErrHandler:
    Debug.Print "Error en Listbox_Trabajo_DblClick: " & Err.Description
End Sub

' Evento que se dispara al hacer doble clic en un registro de la lista de datos ya exportados.
' Requiere autenticacion antes de permitir modificar o eliminar el registro.
Private Sub Listbox_Exportados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrHandler
    
    Dim filaSeleccionada As Long
    filaSeleccionada = Me.Listbox_Exportados.ListIndex
    
    ' Salir si no hay nada seleccionado
    If filaSeleccionada = -1 Then Exit Sub
    
    Debug.Print "Doble clic detectado en la fila " & filaSeleccionada & " de la lista de exportados."
    
    ' PASO 1: Autenticar al usuario
    If AutenticarUsuario() Then
        ' Si la autenticacion es exitosa...
        
        ' PASO 2: Preguntar al usuario que accion desea realizar
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("Que desea hacer con el registro seleccionado?" & vbCrLf & vbCrLf & _
                          "Si = Modificar cantidad" & vbCrLf & _
                          "No = Eliminar registro", _
                          vbQuestion + vbYesNoCancel, "Accion Requerida")
        
        ' PASO 3: Ejecutar la accion seleccionada
        Select Case respuesta
            Case vbYes ' Modificar Cantidad
                If MsgBox("Esta seguro que desea modificar la cantidad de este registro?", vbQuestion + vbYesNo, "Confirmar Modificacion") = vbYes Then
                    Call ModificarCantidadExportada(Me, filaSeleccionada)
                End If
                
            Case vbNo ' Eliminar Registro
                If MsgBox("Esta seguro que desea eliminar permanentemente este registro?", vbQuestion + vbYesNo, "Confirmar Eliminacion") = vbYes Then
                    Call EliminarRegistroExportado(Me, filaSeleccionada)
                End If
                
            Case vbCancel ' Cancelar
                Debug.Print "Accion cancelada por el usuario."
        End Select
        
    End If
    
    Exit Sub
    
ErrHandler:
    Debug.Print "Error en Listbox_Exportados_DblClick: " & Err.Description
End Sub

' Mostrar resumen del sistema
Private Sub MostrarResumenSistema()
    On Error GoTo ErrHandler
    
    Dim resumen As String
    
    resumen = "=== RESUMEN DEL SISTEMA ===" & vbCrLf & vbCrLf
    resumen = resumen & "P�gina 1 (Selecci�n):" & vbCrLf
    resumen = resumen & "- Registros disponibles: " & Me.Listbox_Registros.ListCount & vbCrLf & vbCrLf
    
    resumen = resumen & "P�gina 2 (Trabajo):" & vbCrLf
    resumen = resumen & "- Registros en trabajo: " & Me.Listbox_Trabajo.ListCount & vbCrLf & vbCrLf
    
    resumen = resumen & "P�gina 3 (Revisi�n):" & vbCrLf
    resumen = resumen & "- Registros exportados: " & Me.Listbox_Exportados.ListCount & vbCrLf
    
    MsgBox resumen, vbInformation, "Resumen del Sistema"
    Exit Sub
ErrHandler:
    Debug.Print "Error en MostrarResumenSistema: " & Err.Description
End Sub

' === FUNCION PARA CONFIGURAR ORDEN DE NAVEGACION POR TABS ===
' Configura el orden lógico de navegación según el flujo del proceso
Private Sub ConfigurarOrdenTabs(ByVal frm As Object)
    On Error GoTo ErrHandler
    
    ' === PAGINA 1: SELECCION DE REGISTROS ===
    ' Orden lógico: Palabra Clave -> Area -> Capitulo -> ListBox -> Botones de acción
    
    ' Controles principales de la página 1
    frm.Palabra_Clave.TabIndex = 1
    frm.cmb_Area.TabIndex = 2
    frm.cmb_Capitulos.TabIndex = 3
    frm.Listbox_Registros.TabIndex = 4
    
    ' Botones de la página 1
    frm.btn_Marcar.TabIndex = 5
    frm.btn_Desmarcar.TabIndex = 6
    frm.btn_AgregarATrabajo.TabIndex = 7
    frm.btn_LimpiarCampos.TabIndex = 8
    
    ' === PAGINA 2: AREA DE TRABAJO ===
    ' Orden lógico: Campo cantidad -> Botón asignar -> ListBox trabajo -> Botones de trabajo
    
    frm.txt_Cantidad.TabIndex = 10
    frm.btn_AsignarCantidad.TabIndex = 11
    frm.Listbox_Trabajo.TabIndex = 12
    frm.btn_Exportar.TabIndex = 13
    frm.btn_EliminarSeleccionado.TabIndex = 14
    
    ' === PAGINA 3: REVISION DE EXPORTADOS ===
    ' Orden lógico: ListBox exportados -> Botones de revisión
    
    frm.Listbox_Exportados.TabIndex = 20
    
    ' === CONTROLES GENERALES ===
    ' MultiPage siempre al final para no interferir con la navegación interna
    frm.MultiPage1.TabIndex = 100
    
    Exit Sub
    
ErrHandler:
    Debug.Print "Error en ConfigurarOrdenTabs: " & Err.Description
End Sub

