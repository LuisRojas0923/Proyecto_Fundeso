' === CONSTANTES GLOBALES ===
Private Const ANCHO_LISTBOX As Single = 800
Private Const ANCHOS_COLUMNAS As String = "40 pt;40 pt;350 pt;40 pt;55 pt;55 pt;80 pt;0 pt"

' === DECLARACIÓN DE API WINDOWS PARA DETECCIÓN DE DOBLE CLIC ===
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
' Columnas del ListBox (8 columnas, la última oculta):
' Col 1: Concatenado (40pt) - Col1.Col3.Col5
' Col 2: Codigo Actividad (40pt) - Columna 6
' Col 3: Actividad (350pt) - Columna 7
' Col 4: Unidad (40pt) - Columna 8
' Col 5: Fecha Desde (55pt) - Ingresada por usuario
' Col 6: Fecha Hasta (55pt) - Ingresada por usuario
' Col 7: Observaciones (80pt) - Ingresada por usuario
' Col 8: Area (0pt, oculta) - Columna 2 de EXPORTE_PRESUPUESTO

Private Sub btn_LimpiarCampos_Click()
    On Error GoTo ErrHandler
    Call LimpiarControlesFormulario(Me)
    Call ActualizarControlesOpciones
    Exit Sub
ErrHandler:
    Call LogErrorVBA("btn_LimpiarCampos_Click", Err.Description)
End Sub

Private Sub CommandButton1_Click()
    On Error GoTo ErrHandler
    Call ExportarListboxACrearHojas(Me)
    Exit Sub
ErrHandler:
    Call LogErrorVBA("CommandButton1_Click", Err.Description)
End Sub

Private Sub CommandButton2_Click()
    On Error GoTo ErrHandler
    Call RegistrarAusentismoPorRango(Me)
    Exit Sub
ErrHandler:
    Call LogErrorVBA("CommandButton2_Click", Err.Description)
End Sub

Private Sub CommandButton3_Click()
    On Error GoTo ErrHandler
    Call RegistrarAusentismoPorRango(Me)
    Exit Sub
ErrHandler:
    Call LogErrorVBA("CommandButton3_Click", Err.Description)
End Sub

Private Sub F_Desde_Enter()
    On Error Resume Next
    Set SelectedTextbox = Me.F_Desde
    frmCalendario_.Show
    If Err.Number = 0 Then
        Call GuardarFDesde(Me.F_Desde.Value)
    Else
        Call LogErrorVBA("F_Desde_Enter", Err.Description)
    End If
End Sub

Private Sub F_Hasta_Enter()
    On Error Resume Next
    Set SelectedTextbox = Me.F_Hasta
    frmCalendario_.Show
    If Err.Number = 0 Then
        Call GuardarFHasta(Me.F_Hasta.Value)
    Else
        Call LogErrorVBA("F_Hasta_Enter", Err.Description)
    End If
End Sub

Private Sub btn_Desmarcar_Click()
    On Error GoTo ErrHandler
    Dim item As Long, i As Long
    Titulo = "Refridcol S.A."
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
    Call LogErrorVBA("btn_Desmarcar_Click", Err.Description)
End Sub

Private Sub btn_Marcar_Click()
    On Error GoTo ErrHandler
    Dim i As Long
    Titulo = "Refridcol S.A."
    If Me.Listbox_Registros.ListCount > 0 Then
        For i = 0 To Me.Listbox_Registros.ListCount - 1
            Me.Listbox_Registros.Selected(i) = True
        Next i
    Else
        MsgBox "No hay registros para seleccionar", vbExclamation, Titulo
    End If
    Exit Sub
ErrHandler:
    Call LogErrorVBA("btn_Marcar_Click", Err.Description)
End Sub

' === AJUSTE: Fechas solo en el ListBox, NO en la hoja de datos ===
' Elimina cualquier código que escriba en la hoja EXPORTE_PRESUPUESTO y asegura que las fechas solo se muestren en el ListBox.

' Guardar el valor de F_Desde cuando se activa el control
Private Sub GuardarFDesde(valor As Variant)
    Static FDesdeValor As Variant
    FDesdeValor = valor
    Call LogDebug("FDesdeValor guardado: " & FDesdeValor)
End Sub

' Guardar el valor de F_Hasta cuando se activa el control
Private Sub GuardarFHasta(valor As Variant)
    Static FHastaValor As Variant
    FHastaValor = valor
    Call LogDebug("FHastaValor guardado: " & FHastaValor)
End Sub

' Registrar los valores en las columnas siguientes del ListBox (solo visualización)
Public Sub RegistrarFechasEnListBox(frm As Object)
    On Error GoTo ErrorHandlerGeneral
    Dim i As Long, j As Long, datos() As Variant
    Dim desde As Variant, hasta As Variant, obs As String
    desde = frm.F_Desde.Value
    hasta = frm.F_Hasta.Value
    obs = frm.Observaciones.Value
    Dim nFilas As Long, nCols As Long
    nFilas = frm.Listbox_Registros.ListCount
    nCols = frm.Listbox_Registros.ColumnCount
    If nFilas = 0 Then Exit Sub
    datos = frm.Listbox_Registros.List
    Dim algunoSeleccionado As Boolean: algunoSeleccionado = False
    For i = 0 To nFilas - 1
        If frm.Listbox_Registros.Selected(i) Then
            datos(i, 4) = desde ' Fecha Desde en columna 5
            datos(i, 5) = hasta ' Fecha Hasta en columna 6
            datos(i, 6) = obs   ' Observación en columna 7
            algunoSeleccionado = True
        End If
    Next i
    If algunoSeleccionado Then
        With frm.Listbox_Registros
            .Clear
            .ColumnCount = nCols
            ' Establecer ancho ANTES de asignar ColumnWidths y datos
            .Width = ANCHO_LISTBOX ' Ancho fijo en puntos
            .ColumnWidths = ANCHOS_COLUMNAS
            .List = datos
        End With
    Else
        MsgBox "Debes seleccionar al menos una fila antes de registrar las fechas.", vbExclamation
    End If
    Exit Sub
ErrorHandlerGeneral:
    MsgBox "Error general en RegistrarFechasEnListBox: " & Err.Description, vbCritical
    Call LogErrorVBA("RegistrarFechasEnListBox", Err.Description)
End Sub

Private Sub btn_RegistrarDatos_Click()
    Dim fechaDesde As Date
    Dim fechaHasta As Date
    ' Intenta convertir los valores a fecha
    On Error GoTo FechaInvalida
    fechaDesde = CDate(Me.F_Desde.Value)
    fechaHasta = CDate(Me.F_Hasta.Value)
    On Error GoTo 0
    ' Validar que la fecha de inicio sea menor o igual a la de fin
    If fechaDesde > fechaHasta Then
        MsgBox "La fecha de inicio no puede ser mayor que la fecha de fin.", vbExclamation, "Validación de Fechas"
        Exit Sub
    End If
    ' Si pasa la validación, continúa con el proceso normal
    Call RegistrarFechasEnListBox(Me)
    Exit Sub
FechaInvalida:
    MsgBox "Una o ambas fechas no son válidas. Por favor, verifica el formato.", vbExclamation, "Error de Fecha"
End Sub

' === BOTÓN DE PRUEBA TEMPORAL PARA DEBUGGING ===
Private Sub btn_PruebaLog_Click()
    Call LogInfo("=== PRUEBA DE LOGGING ===")
    Call LogDebug("Este es un mensaje de debug")
    Call LogWarn("Este es un mensaje de advertencia")
    Call LogError("Este es un mensaje de error de prueba")
    Call LogInfo("=== FIN PRUEBA DE LOGGING ===")
    MsgBox "Logs de prueba enviados. Revisa la ventana de depuración (Ctrl+G)", vbInformation, "Prueba de Logging"
End Sub

' === BOTÓN PARA CREACIÓN RÁPIDA DE MEMORIA ===
Private Sub btn_CrearMemoriaRapida_Click()
    On Error GoTo ErrHandler
    Dim indiceSeleccionado As Long
    
    Call LogInfo("=== BOTÓN CREAR MEMORIA RÁPIDA PRESIONADO ===")
    
    ' Verificar que hay una selección válida
    indiceSeleccionado = Me.Listbox_Registros.ListIndex
    Call LogDebug("Indice seleccionado: " & indiceSeleccionado)
    
    If indiceSeleccionado = -1 Then
        MsgBox "Por favor selecciona una fila del ListBox antes de crear la memoria.", vbExclamation, "Selección requerida"
        Call LogWarn("No hay fila seleccionada")
        Exit Sub
    End If
    
    Call LogInfo("Procediendo con creación rápida para fila: " & indiceSeleccionado)
    Call CrearMemoriaRapida(indiceSeleccionado)
    
    Exit Sub
ErrHandler:
    Call LogErrorVBA("btn_CrearMemoriaRapida_Click", Err.Description)
    MsgBox "Error en la creación rápida: " & Err.Description, vbCritical, "Error"
End Sub

' === Variables para detección de doble clic ===
Private ultimoClick As Long
Private ultimoIndice As Long

' === Variable para el formulario de calendario ===
Public SelectedTextbox As Object

' === EVENTO DE CLICK SIMPLE (para referencia) ===
Private Sub Listbox_Registros_Click()
    Call LogDebug("Click simple detectado en ListBox")
End Sub

' === EVENTO DE DOBLE CLIC CORRECTO ===
Private Sub Listbox_Registros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrHandler
    Dim indiceSeleccionado As Long
    
    Call LogInfo("*** DOBLE CLIC DETECTADO CORRECTAMENTE ***")
    
    ' Verificar que hay una selección válida
    indiceSeleccionado = Me.Listbox_Registros.ListIndex
    Call LogDebug("Indice seleccionado en doble clic: " & indiceSeleccionado)
    
    If indiceSeleccionado = -1 Then
        MsgBox "Por favor selecciona una fila haciendo clic en ella antes de hacer doble clic.", vbExclamation, "Selección requerida"
        Call LogWarn("No hay fila seleccionada para doble clic")
        Exit Sub
    End If
    
    ' Llamar directamente a la función de procesamiento
    Call ProcesarDobleClic(indiceSeleccionado)
    
    Exit Sub
ErrHandler:
    Call LogErrorVBA("Listbox_Registros_DblClick", Err.Description)
    Call LogError("Error en Listbox_Registros_DblClick: " & Err.Description)
    MsgBox "Error en el doble clic: " & Err.Description, vbCritical, "Error"
End Sub

' === EVENTO GENÉRICO PARA CUALQUIER LISTBOX (respaldo) ===
Private Sub ListBox_Click()
    Call LogDebug("Evento click genérico detectado")
End Sub

' === EVENTO DE CAMBIO DE SELECCIÓN (DESHABILITADO) ===
Private Sub Listbox_Registros_Change()
    ' Call LogInfo("*** EVENTO CHANGE DETECTADO ***")
    ' Call Listbox_Registros_Click_Handler()
    ' DESHABILITADO: El evento Change interfiere con la detección de doble clic
End Sub

' === EVENTO DE TECLADO PARA CREACIÓN RÁPIDA ===
Private Sub Listbox_Registros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrHandler
    
    ' Detectar Enter o F2 para creación rápida
    If KeyCode = 13 Or KeyCode = 113 Then ' Enter = 13, F2 = 113
        Call LogInfo("*** TECLA PRESIONADA: " & KeyCode & " ***")
        Call LogInfo("Iniciando creación rápida desde teclado")
        
        Dim indiceSeleccionado As Long
        indiceSeleccionado = Me.Listbox_Registros.ListIndex
        
        If indiceSeleccionado = -1 Then
            MsgBox "Por favor selecciona una fila del ListBox antes de crear la memoria.", vbExclamation, "Selección requerida"
            Exit Sub
        End If
        
        Call CrearMemoriaRapida(indiceSeleccionado)
    End If
    
    Exit Sub
ErrHandler:
    Call LogErrorVBA("Listbox_Registros_KeyDown", Err.Description)
End Sub

' === NUEVO: Función para procesar doble clic ===
Private Sub ProcesarDobleClic(indiceSeleccionado As Long)
    On Error GoTo ErrHandler
    Dim nombreHojaEsperado As String
    Dim ws As Worksheet
    
    Call LogInfo("=== INICIANDO PROCESARDOBLECLIC ===")
    Call LogDebug("Indice seleccionado: " & indiceSeleccionado)
    
    ' Verificar que hay una selección válida
    If indiceSeleccionado = -1 Then
        Call LogWarn("Indice seleccionado es -1, no hay selección válida")
        MsgBox "Por favor selecciona una fila haciendo clic en ella antes de hacer doble clic.", vbExclamation, "Selección requerida"
        Exit Sub
    End If
    
    Call LogInfo("Procesando doble clic en fila: " & indiceSeleccionado)
    
    ' Verificar que el ListBox tiene datos
    If Me.Listbox_Registros.ListCount = 0 Then
        Call LogError("ListBox está vacío, no se puede procesar doble clic")
        MsgBox "No hay datos en el ListBox para procesar.", vbExclamation, "ListBox vacío"
        Exit Sub
    End If
    
    Call LogDebug("ListBox tiene " & Me.Listbox_Registros.ListCount & " elementos")
    
    ' Construir el nombre de la hoja según la convención (Col1.Col3.Col5)
    On Error Resume Next
    nombreHojaEsperado = CStr(Me.Listbox_Registros.List(indiceSeleccionado, 0)) ' Ya viene concatenado
    If Err.Number <> 0 Then
        Call LogError("Error al acceder al ListBox: " & Err.Description)
        MsgBox "Error al acceder a los datos del ListBox: " & Err.Description, vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    Call LogDebug("Nombre de hoja esperado: '" & nombreHojaEsperado & "'")
    
    ' Verificar si la hoja ya existe
    Call LogDebug("Verificando si la hoja existe...")
    If SheetExists(nombreHojaEsperado) Then
        ' La hoja existe - navegar a ella y cerrar formulario
        Call LogInfo("Hoja '" & nombreHojaEsperado & "' ya existe. Navegando a la hoja.")
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHojaEsperado)
        If Err.Number <> 0 Then
            Call LogError("Error al acceder a la hoja: " & Err.Description)
            MsgBox "Error al acceder a la hoja: " & Err.Description, vbCritical, "Error"
            Exit Sub
        End If
        On Error GoTo ErrHandler
        ws.Activate
        Unload Me
        MsgBox "La memoria '" & nombreHojaEsperado & "' ya existe. Navegando a la hoja.", vbInformation, "Memoria existente"
    Else
        ' La hoja no existe - proceder con creación rápida
        Call LogInfo("Hoja '" & nombreHojaEsperado & "' no existe. Iniciando creación rápida.")
        Call CrearMemoriaRapida(indiceSeleccionado)
    End If
    
    Call LogInfo("=== FIN PROCESARDOBLECLIC ===")
    Exit Sub
ErrHandler:
    Call LogErrorVBA("ProcesarDobleClic", Err.Description)
    Call LogError("Error en ProcesarDobleClic: " & Err.Description)
    MsgBox "Error en la creación rápida de memoria: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub txt_FechaInicio_Change()
    On Error Resume Next
    frmCalendario_.Show
    If Err.Number <> 0 Then Call LogErrorVBA("txt_FechaInicio_Change", Err.Description)
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    Call LogInfo("=== INICIANDO USERFORM_INITIALIZE ===")
    Call LogDebug("Inicializando variables de doble clic...")
    
    ' Inicializar variables de doble clic
    ultimoClick = 0
    ultimoIndice = -1
    Call LogDebug("Variables de doble clic inicializadas")
    
    ' Centrar el formulario en la pantalla principal
    Me.StartUpPosition = 0 ' Manual
    Me.Left = (Application.Width - Me.Width) / 2 + Application.Left
    Me.Top = (Application.Height - Me.Height) / 2 + Application.Top
    ' === COLOR DE FONDO PERSONALIZADO ===
    ' Me.BackColor = RGB(&H2F, &H67, &HE1) ' Color #2F67E1
    ' Llenar cmb_ITEMS con combinaciones únicas de columnas 1 y 2 de EXPORTE_PRESUPUESTO
    Dim ws As Worksheet, ultFila As Long, i As Long, dict As Object, clave As Variant, concatClave As String
    
    Call LogIniciar("UserForm_Initialize - Iniciando inicializacion...")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Verificar que la hoja existe
    If Not SheetExists("EXPORTE_PRESUPUESTO") Then
        Call LogError("UserForm_Initialize - La hoja EXPORTE_PRESUPUESTO no existe")
        MsgBox "Error: La hoja EXPORTE_PRESUPUESTO no existe en este libro.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("EXPORTE_PRESUPUESTO")
    
    ' Verificar que la hoja no esté vacía
    If ws.Cells(1, 1).Value = "" Then
        Call LogWarn("UserForm_Initialize - La hoja EXPORTE_PRESUPUESTO parece estar vacia")
        MsgBox "Advertencia: La hoja EXPORTE_PRESUPUESTO parece estar vacía.", vbExclamation, "Advertencia"
        Exit Sub
    End If
    
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Call LogDebug("UserForm_Initialize - Ultima fila encontrada: " & ultFila)
    
    ' Verificar que hay datos para procesar
    If ultFila < 1 Then
        Call LogError("UserForm_Initialize - No hay filas de datos para procesar")
        MsgBox "Error: No hay filas de datos para procesar en EXPORTE_PRESUPUESTO.", vbCritical, "Error"
        Exit Sub
    End If
    
    Me.cmb_ITEMS.Clear
    
    ' Procesar solo si hay datos válidos, comenzando desde la fila 2
    For i = 2 To ultFila
        ' Verificar que las celdas existen antes de acceder a ellas
        If Not IsEmpty(ws.Cells(i, 1)) And Not IsEmpty(ws.Cells(i, 2)) Then
            If ws.Cells(i, 1).Value <> "" And ws.Cells(i, 2).Value <> "" Then
                concatClave = ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value
                If Not dict.Exists(concatClave) Then 
                    dict.Add concatClave, ws.Cells(i, 1).Value & "|" & ws.Cells(i, 2).Value
                    Call LogDebug("UserForm_Initialize - Agregado item: " & concatClave)
                End If
            End If
        End If
    Next i
    
    Call LogInfo("UserForm_Initialize - Total de items unicos encontrados: " & dict.Count)
    
    For Each clave In dict.Keys
        Me.cmb_ITEMS.AddItem clave
    Next clave
    
    ' No seleccionar ningún ITEM al iniciar
    Me.cmb_ITEMS.ListIndex = -1
    Me.cmb_Capitulo.ListIndex = -1
    Me.cmb_Capitulo.Enabled = False

    ' Configurar ListBox con ancho fijo desde el inicio
    With Me.Listbox_Registros
        .Clear
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .ColumnCount = 8
        .Width = ANCHO_LISTBOX ' Ancho fijo desde la inicialización
        .ColumnWidths = ANCHOS_COLUMNAS
    End With
    
    Call LogFinalizar("UserForm_Initialize - Inicializacion completada exitosamente")
    
    ' === PRUEBA: Verificar que los eventos funcionan ===
    Call LogInfo("=== PRUEBA DE EVENTOS ===")
    Call LogInfo("Formulario inicializado correctamente")
    Call LogInfo("Variables de doble clic: ultimoClick=" & ultimoClick & ", ultimoIndice=" & ultimoIndice)
    Call LogInfo("=== FIN PRUEBA DE EVENTOS ===")
    
    Exit Sub
ErrHandler:
    Call LogErrorVBA("UserForm_Initialize", Err.Description, Erl)
End Sub

' === FUNCIÓN AUXILIAR: Verificar si una hoja existe ===
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' === FUNCIÓN AUXILIAR: Leer datos de hoja en bloque (OPTIMIZACIÓN) ===
Private Function LeerDatosHojaEnBloque(ws As Worksheet, ultFila As Long) As Variant
    On Error GoTo ErrorHandler
    ' Leer todo el rango de datos de una vez (mucho más eficiente)
    Dim rangoDatos As Range
    Set rangoDatos = ws.Range("A2:H" & ultFila) ' Columnas A-H desde fila 2
    LeerDatosHojaEnBloque = rangoDatos.Value
    Call LogDebug("Datos leidos en bloque: " & ultFila - 1 & " filas x 8 columnas")
    Exit Function
ErrorHandler:
    Call LogError("Error al leer datos en bloque: " & Err.Description)
    LeerDatosHojaEnBloque = Empty
End Function

' === NUEVA FUNCIÓN: Creación rápida de memoria con flujo guiado ===
Private Sub CrearMemoriaRapida(indiceFila As Long)
    On Error GoTo ErrHandler
    
    Call LogInfo("=== INICIANDO CREARMEMORIARAPIDA (SIMULANDO FLUJO MANUAL) ===")
    Call LogInfo("Iniciando creación rápida de memoria para fila: " & indiceFila)
    Call LogDebug("Indice fila recibido: " & indiceFila)
    
    ' PASO 1: Asignar el campo F_Desde y abrir calendario
    Call LogInfo("=== PASO 1: SOLICITANDO FECHA DE INICIO (F_Desde) ===")
    Call LogInfo("Simulando entrada en campo F_Desde...")
    
    ' Simular el comportamiento de F_Desde_Enter()
    Set SelectedTextbox = Me.F_Desde
    Call LogDebug("SelectedTextbox asignado a F_Desde")
    
    Call LogInfo("Abriendo formulario de calendario para F_Desde...")
    frmCalendario_.Show
    
    ' Verificar si se asignó la fecha
    If Me.F_Desde.Value = "" Then
        Call LogWarn("*** USUARIO CANCELÓ LA SELECCIÓN DE FECHA DE INICIO ***")
        Exit Sub
    End If
    
    Call LogInfo("*** FECHA DE INICIO ASIGNADA A F_Desde: " & Me.F_Desde.Value & " ***")
    
    ' PASO 2: Asignar el campo F_Hasta y abrir calendario
    Call LogInfo("=== PASO 2: SOLICITANDO FECHA DE FIN (F_Hasta) ===")
    Call LogInfo("Simulando entrada en campo F_Hasta...")
    
    ' Simular el comportamiento de F_Hasta_Enter()
    Set SelectedTextbox = Me.F_Hasta
    Call LogDebug("SelectedTextbox asignado a F_Hasta")
    
    Call LogInfo("Abriendo formulario de calendario para F_Hasta...")
    frmCalendario_.Show
    
    ' Verificar si se asignó la fecha
    If Me.F_Hasta.Value = "" Then
        Call LogWarn("*** USUARIO CANCELÓ LA SELECCIÓN DE FECHA DE FIN ***")
        Exit Sub
    End If
    
    Call LogInfo("*** FECHA DE FIN ASIGNADA A F_Hasta: " & Me.F_Hasta.Value & " ***")
    
    ' PASO 3: Ejecutar la lógica de btn_RegistrarDatos_Click()
    Call LogInfo("=== PASO 3: EJECUTANDO LÓGICA DE REGISTRO DE DATOS ===")
    Call LogInfo("Simulando clic en btn_RegistrarDatos...")
    
    Dim fechaDesde As Date
    Dim fechaHasta As Date
    
    ' Intenta convertir los valores a fecha
    On Error GoTo FechaInvalida
    fechaDesde = CDate(Me.F_Desde.Value)
    fechaHasta = CDate(Me.F_Hasta.Value)
    On Error GoTo ErrHandler
    
    Call LogDebug("Fechas convertidas - Desde: " & fechaDesde & ", Hasta: " & fechaHasta)
    
    ' Validar que la fecha de inicio sea menor o igual a la de fin
    If fechaDesde > fechaHasta Then
        MsgBox "La fecha de inicio no puede ser mayor que la fecha de fin.", vbExclamation, "Validación de Fechas"
        Call LogWarn("Fechas inválidas - Inicio mayor que fin")
        Exit Sub
    End If
    
    Call LogInfo("*** FECHAS VALIDADAS CORRECTAMENTE ***")
    
    ' Si pasa la validación, continúa con el proceso normal
    Call LogInfo("Llamando a RegistrarFechasEnListBox...")
    Call RegistrarFechasEnListBox(Me)
    Call LogInfo("*** FECHAS REGISTRADAS EN LISTBOX EXITOSAMENTE ***")
    
    ' PASO 4: Mostrar confirmación para crear la memoria
    Call LogInfo("=== PASO 4: CONFIRMACIÓN DE CREACIÓN DE MEMORIA ===")
    
    ' Obtener datos de la fila seleccionada para mostrar en la confirmación
    Dim actividad As String, area As String, unidad As String
    Dim fechaDesdeStr As String, fechaHastaStr As String
    
    actividad = Me.Listbox_Registros.List(indiceFila, 2) ' Columna Actividad
    area = Me.Listbox_Registros.List(indiceFila, 7)      ' Columna Area (oculta)
    unidad = Me.Listbox_Registros.List(indiceFila, 3)    ' Columna Unidad
    fechaDesdeStr = Me.Listbox_Registros.List(indiceFila, 4) ' Fecha Desde
    fechaHastaStr = Me.Listbox_Registros.List(indiceFila, 5) ' Fecha Hasta
    
    Call LogDebug("Datos para confirmación - Actividad: " & actividad & ", Area: " & area & ", Unidad: " & unidad)
    Call LogDebug("Fechas - Desde: " & fechaDesdeStr & ", Hasta: " & fechaHastaStr)
    
    ' Crear mensaje de confirmación detallado
    Dim mensajeConfirmacion As String
    mensajeConfirmacion = "¿Deseas crear la memoria con los siguientes datos?" & vbCrLf & vbCrLf & _
                         "ACTIVIDAD: " & actividad & vbCrLf & _
                         "AREA: " & area & vbCrLf & _
                         "UNIDAD: " & unidad & vbCrLf & _
                         "FECHA DESDE: " & fechaDesdeStr & vbCrLf & _
                         "FECHA HASTA: " & fechaHastaStr & vbCrLf & vbCrLf & _
                         "Se creara una nueva hoja de memoria con estos datos."
    
    Call LogInfo("Mostrando mensaje de confirmación de creación...")
    Dim respuestaCreacion As VbMsgBoxResult
    respuestaCreacion = MsgBox(mensajeConfirmacion, vbYesNo + vbQuestion, "Confirmar Creación de Memoria")
    
    Call LogDebug("Respuesta del usuario: " & IIf(respuestaCreacion = vbYes, "SÍ", "NO"))
    
    If respuestaCreacion = vbNo Then
        Call LogWarn("*** USUARIO CANCELÓ LA CREACIÓN DE MEMORIA ***")
        Call LogInfo("Proceso cancelado. Las fechas quedan registradas en el ListBox.")
        Exit Sub
    End If
    
    ' PASO 5: Ejecutar creación de memoria (igual que CommandButton1_Click)
    Call LogInfo("=== PASO 5: CREANDO MEMORIA ===")
    Call LogInfo("*** USUARIO CONFIRMÓ LA CREACIÓN DE MEMORIA ***")
    
    ' IMPORTANTE: Seleccionar la fila antes de exportar
    Call LogDebug("Seleccionando fila " & indiceFila & " en el ListBox...")
    Me.Listbox_Registros.Selected(indiceFila) = True
    Me.Listbox_Registros.ListIndex = indiceFila
    Call LogDebug("Fila " & indiceFila & " seleccionada correctamente")
    
    Call LogInfo("Ejecutando ExportarListboxACrearHojas (igual que CommandButton1_Click)...")
    
    On Error GoTo ErrCreacion
    Call ExportarListboxACrearHojas(Me)
    Call LogInfo("*** MEMORIA CREADA EXITOSAMENTE ***")
    On Error GoTo ErrHandler
    
    Call LogInfo("*** CREACIÓN RÁPIDA COMPLETADA EXITOSAMENTE ***")
    Call LogInfo("=== FIN CREARMEMORIARAPIDA ===")
    Exit Sub
    
ErrCreacion:
    Call LogErrorVBA("CrearMemoriaRapida - ExportarListboxACrearHojas", Err.Description)
    Call LogError("Error al crear memoria: " & Err.Description)
    MsgBox "Error al crear la memoria: " & Err.Description, vbCritical, "Error en Creación"
    Resume Next
    
FechaInvalida:
    MsgBox "Una o ambas fechas no son válidas. Por favor, verifica el formato.", vbExclamation, "Error de Fecha"
    Call LogError("Error de fecha inválida en CrearMemoriaRapida")
    Exit Sub
    
ErrHandler:
    Call LogErrorVBA("CrearMemoriaRapida", Err.Description)
    Call LogError("Error en CrearMemoriaRapida: " & Err.Description)
    MsgBox "Error en la creación rápida de memoria: " & Err.Description, vbCritical, "Error"
End Sub

' === AJUSTE FINAL: Código limpio y unificado bajo cmb_ITEMS ===
' El procedimiento CargarRegistros y eventos relacionados quedan obsoletos.

' === CORRECCIÓN FILTRO: Mostrar columnas 3 a 5 y filtrar correctamente por ITEMS (antes área) ===
Private Sub cmb_ITEMS_Change()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, ultFila As Long, i As Long, datos(), filaDestino As Long
    Dim filtroITEMS As String, valorCol1 As String, valorCol2 As String
    
    Call LogIniciar("cmb_ITEMS_Change - Iniciando filtro...")
    
    ' Limpiar controles dependientes
    Me.Listbox_Registros.Clear
    
    Application.EnableEvents = False
    Me.cmb_Capitulo.Clear
    Me.cmb_Capitulo.Enabled = False
    Application.EnableEvents = True
    
    If Me.cmb_ITEMS.ListIndex = -1 Then
        Call LogDebug("cmb_ITEMS_Change - No hay item seleccionado")
        Exit Sub
    End If
    
    filtroITEMS = Me.cmb_ITEMS.Value
    Call LogDebug("cmb_ITEMS_Change - Filtro seleccionado: " & filtroITEMS)
    
    ' Extraer los valores de las dos primeras columnas
    valorCol1 = Trim(Split(filtroITEMS, " - ")(0))
    valorCol2 = Trim(Mid(filtroITEMS, InStr(filtroITEMS, " - ") + 3))
    
    Call LogDebug("cmb_ITEMS_Change - Valor Col1: '" & valorCol1 & "'")
    Call LogDebug("cmb_ITEMS_Change - Valor Col2: '" & valorCol2 & "'")
    
    ' Verificar que la hoja existe
    If Not SheetExists("EXPORTE_PRESUPUESTO") Then
        Call LogError("cmb_ITEMS_Change - La hoja EXPORTE_PRESUPUESTO no existe")
        MsgBox "Error: La hoja EXPORTE_PRESUPUESTO no existe en este libro.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("EXPORTE_PRESUPUESTO")
    
    ' Verificar que la hoja no esté vacía
    If ws.Cells(1, 1).Value = "" Then
        Call LogWarn("cmb_ITEMS_Change - La hoja EXPORTE_PRESUPUESTO parece estar vacia")
        MsgBox "Advertencia: La hoja EXPORTE_PRESUPUESTO parece estar vacía.", vbExclamation, "Advertencia"
        Exit Sub
    End If
    
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Call LogDebug("cmb_ITEMS_Change - Ultima fila encontrada: " & ultFila)
    
    ' Verificar que hay datos para procesar
    If ultFila < 2 Then ' Se necesita al menos 1 fila de datos más el encabezado
        Call LogError("cmb_ITEMS_Change - No hay filas de datos para procesar")
        MsgBox "Error: No hay filas de datos para procesar en EXPORTE_PRESUPUESTO.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' --- Cargar ListBox y cmb_Capitulo simultáneamente (OPTIMIZADO) ---
    Dim dictCapitulos As Object
    Set dictCapitulos = CreateObject("Scripting.Dictionary")
    
    ' Leer todos los datos de la hoja en bloque (MUCHO MÁS EFICIENTE)
    Dim datosHoja As Variant
    datosHoja = LeerDatosHojaEnBloque(ws, ultFila)
    
    If IsEmpty(datosHoja) Then
        Call LogError("No se pudieron leer los datos de la hoja")
        Exit Sub
    End If
    
    Call LogDebug("cmb_ITEMS_Change - Procesando " & UBound(datosHoja, 1) & " filas de datos")
    
    filaDestino = 0
    ' Procesar datos en memoria (mucho más rápido)
    For i = 1 To UBound(datosHoja, 1)
        If Not IsEmpty(datosHoja(i, 1)) And Not IsEmpty(datosHoja(i, 2)) Then
            If Trim(CStr(datosHoja(i, 1))) = valorCol1 And Trim(CStr(datosHoja(i, 2))) = valorCol2 Then
                filaDestino = filaDestino + 1
                ' Añadir capitulo al diccionario
                Dim claveCapitulo As String
                If Not IsEmpty(datosHoja(i, 3)) And Not IsEmpty(datosHoja(i, 4)) Then
                    claveCapitulo = Trim(CStr(datosHoja(i, 3))) & " - " & Trim(CStr(datosHoja(i, 4)))
                    If Not dictCapitulos.Exists(claveCapitulo) Then
                        dictCapitulos.Add claveCapitulo, 1
                    End If
                End If
            End If
        End If
    Next i
    
    If filaDestino = 0 Then
        Call LogWarn("Filtro: No se encontraron coincidencias para '" & valorCol1 & " - " & valorCol2 & "'")
        Exit Sub
    End If
    
    ' Llenar cmb_Capitulo
    If dictCapitulos.Count > 0 Then
        Dim clave As Variant
        For Each clave In dictCapitulos.Keys
            Me.cmb_Capitulo.AddItem clave
        Next clave
        Me.cmb_Capitulo.Enabled = True
        Call LogInfo("cmb_ITEMS_Change - Se cargaron " & dictCapitulos.Count & " capitulos")
    End If
    
    ' Llenar el ListBox con todos los resultados del primer filtro (OPTIMIZADO)
    ReDim datos(1 To filaDestino, 1 To 8)
    filaDestino = 1
    ' Procesar datos en memoria (mucho más rápido)
    For i = 1 To UBound(datosHoja, 1)
        If Not IsEmpty(datosHoja(i, 1)) And Not IsEmpty(datosHoja(i, 2)) Then
            If Trim(CStr(datosHoja(i, 1))) = valorCol1 And Trim(CStr(datosHoja(i, 2))) = valorCol2 Then
                ' Nuevo orden: Concatenado, Codigo Actividad, Actividad, Unidad, Fecha Desde, Fecha Hasta, Observaciones
                Dim concatenado As String
                concatenado = CStr(datosHoja(i, 1)) & "." & CStr(datosHoja(i, 3)) & "." & CStr(datosHoja(i, 5))
                datos(filaDestino, 1) = concatenado
                If Not IsEmpty(datosHoja(i, 6)) Then datos(filaDestino, 2) = datosHoja(i, 6) Else datos(filaDestino, 2) = "" ' Codigo Actividad
                If Not IsEmpty(datosHoja(i, 7)) Then datos(filaDestino, 3) = datosHoja(i, 7) Else datos(filaDestino, 3) = "" ' Actividad
                If Not IsEmpty(datosHoja(i, 8)) Then datos(filaDestino, 4) = datosHoja(i, 8) Else datos(filaDestino, 4) = "" ' Unidad
                datos(filaDestino, 5) = "" ' Fecha Desde
                datos(filaDestino, 6) = "" ' Fecha Hasta
                datos(filaDestino, 7) = "" ' Observación
                If Not IsEmpty(datosHoja(i, 2)) Then datos(filaDestino, 8) = datosHoja(i, 2) Else datos(filaDestino, 8) = "" ' Area (oculta)
                filaDestino = filaDestino + 1
            End If
        End If
    Next i
    With Me.Listbox_Registros
        .Clear
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .ColumnCount = 8
        ' Establecer ancho ANTES de asignar ColumnWidths y datos
        .Width = ANCHO_LISTBOX ' Ancho fijo en puntos
        .ColumnWidths = ANCHOS_COLUMNAS
        .List = datos
        Call LogDebug("cmb_ITEMS_Change - ColumnWidths: " & .ColumnWidths)
        Call LogDebug("cmb_ITEMS_Change - ListBox.Width: " & .Width)
    End With
    Call LogInfo("Filtro: Se encontraron " & filaDestino - 1 & " coincidencias para '" & valorCol1 & " - " & valorCol2 & "'")
    Exit Sub
ErrHandler:
    Call LogErrorVBA("cmb_ITEMS_Change", Err.Description)
End Sub

Private Sub cmb_Capitulo_Change()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, ultFila As Long, i As Long, datos(), filaDestino As Long
    
    ' Valores del primer filtro (ITEMS)
    Dim filtroITEMS As String, valorCol1 As String, valorCol2 As String
    
    ' Valores del segundo filtro (Capitulo)
    Dim filtroCapitulo As String, valorCol3 As String, valorCol4 As String

    Call LogIniciar("cmb_Capitulo_Change - Iniciando segundo filtro...")
    Me.Listbox_Registros.Clear

    ' --- Validar que ambos filtros tengan una selección ---
    If Me.cmb_ITEMS.ListIndex = -1 Then
        Call LogWarn("cmb_Capitulo_Change - No hay un ITEM seleccionado. Filtro abortado.")
        Exit Sub
    End If
    
    If Me.cmb_Capitulo.ListIndex = -1 Then
        ' Si se deselecciona el capítulo, se debería recargar el ListBox con el filtro de ITEMS
        Call cmb_ITEMS_Change
        Exit Sub
    End If

    ' --- Obtener valores de ambos filtros ---
    filtroITEMS = Me.cmb_ITEMS.Value
    valorCol1 = Trim(Split(filtroITEMS, " - ")(0))
    valorCol2 = Trim(Mid(filtroITEMS, InStr(filtroITEMS, " - ") + 3))

    filtroCapitulo = Me.cmb_Capitulo.Value
    valorCol3 = Trim(Split(filtroCapitulo, " - ")(0))
    valorCol4 = Trim(Mid(filtroCapitulo, InStr(filtroCapitulo, " - ") + 3))

    Call LogDebug("Filtro 1 (ITEMS): " & valorCol1 & " - " & valorCol2)
    Call LogDebug("Filtro 2 (Capitulo): " & valorCol3 & " - " & valorCol4)

    ' --- Aplicar filtro combinado a la hoja (OPTIMIZADO) ---
    Set ws = ThisWorkbook.Sheets("EXPORTE_PRESUPUESTO")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Leer todos los datos de la hoja en bloque (MUCHO MÁS EFICIENTE)
    Dim datosHoja As Variant
    datosHoja = LeerDatosHojaEnBloque(ws, ultFila)
    
    If IsEmpty(datosHoja) Then
        Call LogError("No se pudieron leer los datos de la hoja")
        Exit Sub
    End If
    
    Call LogDebug("cmb_Capitulo_Change - Procesando " & UBound(datosHoja, 1) & " filas de datos")

    ' Contar coincidencias (procesando en memoria)
    filaDestino = 0
    For i = 1 To UBound(datosHoja, 1)
        If Trim(CStr(datosHoja(i, 1))) = valorCol1 And Trim(CStr(datosHoja(i, 2))) = valorCol2 Then
            ' Aplicar segundo filtro
            If Trim(CStr(datosHoja(i, 3))) = valorCol3 And Trim(CStr(datosHoja(i, 4))) = valorCol4 Then
                filaDestino = filaDestino + 1
            End If
        End If
    Next i

    If filaDestino = 0 Then
        Call LogWarn("Filtro combinado no encontró resultados.")
        Exit Sub
    End If

    ' Llenar el ListBox con los resultados del filtro combinado (OPTIMIZADO)
    ReDim datos(1 To filaDestino, 1 To 8)
    filaDestino = 1
    ' Procesar datos en memoria (mucho más rápido)
    For i = 1 To UBound(datosHoja, 1)
        If Trim(CStr(datosHoja(i, 1))) = valorCol1 And Trim(CStr(datosHoja(i, 2))) = valorCol2 Then
            If Trim(CStr(datosHoja(i, 3))) = valorCol3 And Trim(CStr(datosHoja(i, 4))) = valorCol4 Then
                ' Mapeo de datos
                Dim concatenado As String
                concatenado = CStr(datosHoja(i, 1)) & "." & CStr(datosHoja(i, 3)) & "." & CStr(datosHoja(i, 5))
                datos(filaDestino, 1) = concatenado
                If Not IsEmpty(datosHoja(i, 6)) Then datos(filaDestino, 2) = datosHoja(i, 6) Else datos(filaDestino, 2) = ""
                If Not IsEmpty(datosHoja(i, 7)) Then datos(filaDestino, 3) = datosHoja(i, 7) Else datos(filaDestino, 3) = ""
                If Not IsEmpty(datosHoja(i, 8)) Then datos(filaDestino, 4) = datosHoja(i, 8) Else datos(filaDestino, 4) = ""
                datos(filaDestino, 5) = "" ' Fecha Desde
                datos(filaDestino, 6) = "" ' Fecha Hasta
                datos(filaDestino, 7) = "" ' Observación
                If Not IsEmpty(datosHoja(i, 2)) Then datos(filaDestino, 8) = datosHoja(i, 2) Else datos(filaDestino, 8) = "" ' Area (oculta)
                filaDestino = filaDestino + 1
            End If
        End If
    Next i

    ' Poblar ListBox
    With Me.Listbox_Registros
        .Clear
        .ColumnCount = 8
        .Width = ANCHO_LISTBOX
        .ColumnWidths = ANCHOS_COLUMNAS
        .List = datos
    End With
    
    Call LogInfo("Filtro combinado encontró " & filaDestino - 1 & " resultados.")
    Exit Sub
ErrHandler:
    Call LogErrorVBA("cmb_Capitulo_Change", Err.Description)
End Sub





