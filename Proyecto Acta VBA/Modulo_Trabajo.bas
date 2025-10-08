Attribute VB_Name = "Modulo_Trabajo"

' === MÓDULO PARA GESTIÓN DEL LISTBOX DE TRABAJO ===

' Función para agregar registros al ListBox de Trabajo con consecutivos automáticos
Public Sub AgregarRegistrosATrabajo(frm As Object, areaCompleta As String, capituloCompleto As String, cantidad As Double)
    On Error GoTo ErrHandler
    
    Dim i As Long, j As Long
    Dim registrosAgregados As Long
    Dim consecutivoActividad As Long
    Dim consActividadFinal As Long ' <-- Variable añadida
    Dim datosTrabajo() As Variant
    Dim filaDestino As Long
    
    ' Validar que el formulario y el ListBox existan
    If frm Is Nothing Then
        Debug.Print "ERROR: El formulario es Nothing"
        Exit Sub
    End If
    
    If frm.Listbox_Registros Is Nothing Then
        Debug.Print "ERROR: Listbox_Registros es Nothing"
        Exit Sub
    End If
    
    ' Validar que el ListBox esté correctamente configurado
    If Not ValidarListBoxRegistros(frm) Then
        Debug.Print "ERROR: ListBox no está correctamente configurado, intentando reinicializar..."
        Call ReinicializarListBoxRegistros(frm)
        
        ' Validar nuevamente después de la reinicialización
        If Not ValidarListBoxRegistros(frm) Then
            MsgBox "Error: No se pudo configurar correctamente el ListBox. Por favor, reinicie la aplicación.", vbCritical, "Error de Configuración"
            Exit Sub
        End If
    End If
    
    ' Extraer consecutivo y área del valor seleccionado (ej: "1 - UBA")
    Dim consecutivoArea As String, area As String
    Call ObtenerValorComboBoxArea(areaCompleta, consecutivoArea, area)
    
    ' Extraer consecutivo y capítulo del valor seleccionado (ej: "2 - ESTRUCTURA")
    Dim consecutivoCapitulo As String, capitulo As String
    Call ObtenerValorComboBoxCapitulo(capituloCompleto, consecutivoCapitulo, capitulo)
    
    ' Obtener solo el consecutivo de la actividad (el de capítulo ya es fijo)
    consecutivoActividad = ObtenerConsecutivoActividad(frm, area, capitulo)
    
    ' Contar registros que se van a agregar
    registrosAgregados = 0
    With frm.Listbox_Registros
        Debug.Print "ListBox_Registros tiene " & .ListCount & " registros totales"
        
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                registrosAgregados = registrosAgregados + 1
                Debug.Print "Registro " & (i + 1) & " está seleccionado"
            End If
        Next i
    End With
    
    Debug.Print "Total de registros seleccionados: " & registrosAgregados
    
    ' Si no se seleccionó nada, salir
    If registrosAgregados = 0 Then 
        MsgBox "No se ha seleccionado ningún registro para agregar.", vbInformation, "Selección Vacía"
        Exit Sub
    End If
    
    ' Si la cantidad no es válida, salir
    If cantidad <= 0 Then
        MsgBox "La cantidad debe ser un número mayor que cero.", vbExclamation, "Cantidad no válida"
        Exit Sub
    End If
    
    ' --- INICIO DE PROCESAMIENTO DE REGISTROS SELECCIONADOS ---
    ' Inicializar la colección para almacenar los registros que se agregarán.
    Dim nuevosRegistros As New Collection
    Dim nuevoRegistro() As Variant
    
    ' Procesar registros seleccionados
    With frm.Listbox_Registros
        Debug.Print "Iniciando procesamiento de " & .ListCount & " registros del ListBox_Registros"
        Debug.Print "ColumnCount del ListBox_Registros: " & .ColumnCount
        
        ' Validaciones básicas del ListBox de origen
        If .ListCount = 0 Then GoTo FinProcesamiento
        If .ColumnCount < 4 Then GoTo FinProcesamiento
        
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Debug.Print "Procesando registro seleccionado " & (i + 1) & " del ListBox_Registros"
                
                ' Extraer datos del registro seleccionado
                Dim codigoActividad As String, descripcionActividad As String, unidadActividad As String, precioActividad As String
                
                ' Índices actualizados para leer de la ListBox de 4 columnas:
                ' Col 1 (índice 0): Código
                ' Col 2 (índice 1): Descripción
                ' Col 3 (índice 2): Unidad
                ' Col 4 (índice 3): Precio
                codigoActividad = ExtraerDatoListBox(frm.Listbox_Registros, i, 0)
                descripcionActividad = ExtraerDatoListBox(frm.Listbox_Registros, i, 1)
                unidadActividad = ExtraerDatoListBox(frm.Listbox_Registros, i, 2)
                precioActividad = ExtraerDatoListBox(frm.Listbox_Registros, i, 3)
                
                Debug.Print "  Código: '" & codigoActividad & "', Desc: '" & descripcionActividad & "'" & vbCrLf & _
                "  Precio Original en ListBox_Registros: '" & precioActividad & "'"

                ' --- VALIDACIÓN DE ACTIVIDAD DUPLICADA ---
                If ManejarActividadExistente(frm, area, capitulo, codigoActividad, descripcionActividad, cantidad) Then
                    ' Actividad ya existía y fue manejada (sumada o ignorada). Saltar al siguiente.
                    Debug.Print "  Actividad '" & codigoActividad & "' ya existía. Saltando..."
                Else
                    ' Es una actividad nueva. Preparar un array para la fila y agregarlo a la colección.
                    Debug.Print "  Actividad '" & codigoActividad & "' es nueva. Agregando a la colección..."
                    ReDim nuevoRegistro(1 To 11)
                    
                    ' Llenar la fila de datos
                    Debug.Print "    -> Concatenando: Area=" & consecutivoArea & ", Capitulo=" & consecutivoCapitulo & ", Actividad=" & consecutivoActividad
                    nuevoRegistro(1) = consecutivoArea & "." & consecutivoCapitulo & "." & consecutivoActividad ' <--- CAMBIO: Consecutivo concatenado
                    nuevoRegistro(2) = area
                    nuevoRegistro(3) = consecutivoCapitulo
                    nuevoRegistro(4) = capitulo
                    nuevoRegistro(5) = consecutivoActividad
                    nuevoRegistro(6) = codigoActividad
                    nuevoRegistro(7) = descripcionActividad
                    nuevoRegistro(8) = unidadActividad
                    nuevoRegistro(9) = cantidad
                    nuevoRegistro(10) = precioActividad
                    
                    ' Calcular Valor Parcial
                    Dim precioUnitario As Double
                    Dim precioTexto As String: precioTexto = precioActividad
                    If precioTexto <> "" And Left(precioTexto, 1) = "$" Then precioTexto = Mid(precioTexto, 2)
                    ' Reemplazar la coma por punto para el procesamiento decimal
                    ' Procesar el formato del precio correctamente
                ' 1. Primero quitamos los separadores de miles (puntos)
                precioTexto = Replace(precioTexto, ".", "")
                ' 2. Reemplazamos la coma decimal por punto
                precioTexto = Replace(precioTexto, ",", ".")
                ' 3. Si termina en .0, .00, etc., lo dividimos por la potencia de 10 correspondiente
                If InStr(precioTexto, ".") > 0 Then
                    Dim decimales As Integer
                    decimales = Len(precioTexto) - InStr(precioTexto, ".") ' Número de decimales
                    If decimales > 0 Then
                        precioTexto = CDbl(precioTexto) / (10 ^ decimales)
                    End If
                End If
                    If IsNumeric(precioTexto) Then precioUnitario = CDbl(precioTexto) Else precioUnitario = 0
                    
                    ' CORRECCION: Formatear los valores monetarios usando punto como separador decimal
                    ' y coma como separador de miles para consistencia
                    nuevoRegistro(10) = Format(precioUnitario, "$#,##0.00")  ' Formato VR. UNITARIO
                    nuevoRegistro(11) = Format(precioUnitario * cantidad, "$#,##0.00")  ' Formato VR. PARCIAL
                    
                    ' Asegurar que los separadores sean consistentes
                    nuevoRegistro(10) = Replace(Replace(nuevoRegistro(10), ".", ","), ",", ".")
                    nuevoRegistro(11) = Replace(Replace(nuevoRegistro(11), ".", ","), ",", ".")
                    
                    ' Log de transformación de precio
                    Debug.Print "=== TRANSFORMACIÓN DE PRECIO ==="
                    Debug.Print "Precio Original: " & precioActividad
                    Debug.Print "Después de limpiar $: " & precioTexto
                    Debug.Print "Convertido a Número: " & precioUnitario
                    Debug.Print "Formato Final VR. UNITARIO: " & nuevoRegistro(10)
                    Debug.Print "Formato Final VR. PARCIAL: " & nuevoRegistro(11)
                    Debug.Print "=============================="
                    
                    ' Agregar la fila a la coleccion
                    nuevosRegistros.Add nuevoRegistro
                    
                    ' Incrementar el consecutivo de actividad solo para registros nuevos
                    consecutivoActividad = consecutivoActividad + 1
                End If
            End If
        Next i
    End With

FinProcesamiento:
    ' --- FIN DE PROCESAMIENTO DE REGISTROS SELECCIONADOS ---
    Debug.Print "--- Fin del bucle de procesamiento. Registros procesados: " & registrosAgregados & " ---"
    
    ' --- VOLCADO DE DATOS DE LA COLECCION AL ARRAY FINAL ---
    Dim datosNuevos() As Variant
    Dim contadorNuevos As Long
    contadorNuevos = nuevosRegistros.Count
    
    ' Validar que la coleccion tenga datos
    If contadorNuevos = 0 Then
        Debug.Print "La coleccion 'nuevosRegistros' esta vacia. No hay nada que agregar al ListBox de Trabajo."
        ' Mostrar el MsgBox aqui tambien para informar al usuario
        MsgBox "Registros agregados al trabajo: " & registrosAgregados & " con cantidad: " & cantidad, vbInformation
        Exit Sub
    End If
        
    ' Dimensionar el array final
    ReDim datosNuevos(1 To contadorNuevos, 1 To 11)
    
    ' Volcar los datos de la coleccion al array 2D
    For i = 1 To contadorNuevos
        Dim registroActual As Variant
        registroActual = nuevosRegistros(i)
        For j = 1 To 11
            datosNuevos(i, j) = registroActual(j)
        Next j
    Next i
    
    ' Debug: Verificar contenido del array datosTrabajo
    Debug.Print "=== CONTENIDO DEL ARRAY datosTrabajo ==="
    Debug.Print "Dimensiones del array: " & UBound(datosNuevos, 1) & " x " & UBound(datosNuevos, 2)
    Debug.Print "contadorNuevos: " & contadorNuevos
    
    For i = 1 To contadorNuevos
        Debug.Print "Registro " & i & ":"
        For j = 1 To 11
            If IsEmpty(datosNuevos(i, j)) Then
                Debug.Print "  Col " & j & ": EMPTY"
            ElseIf IsNull(datosNuevos(i, j)) Then
                Debug.Print "  Col " & j & ": NULL"
            Else
                Debug.Print "  Col " & j & ": '" & CStr(datosNuevos(i, j)) & "'"
            End If
        Next j
    Next i
    Debug.Print "========================================"
    
    ' Validar que el array tenga datos validos
    If contadorNuevos <= 0 Then
        Debug.Print "ERROR: No hay registros para agregar (contadorNuevos = " & contadorNuevos & ")"
        Exit Sub
    End If
    
    ' Validar que el ListBox de Trabajo este correctamente configurado
    If Not ValidarListBoxTrabajo(frm) Then
        Debug.Print "ERROR: ListBox_Trabajo no esta correctamente configurado, intentando reinicializar..."
        Call ReinicializarListBoxTrabajo(frm)
        
        ' Validar nuevamente despues de la reinicializacion
        If Not ValidarListBoxTrabajo(frm) Then
            MsgBox "Error: No se pudo configurar correctamente el ListBox de Trabajo. Por favor, reinicie la aplicacion.", vbCritical, "Error de Configuracion"
            Exit Sub
        End If
        End If
        
    ' Validar que el array este correctamente dimensionado
    If UBound(datosNuevos, 1) < contadorNuevos Or UBound(datosNuevos, 2) < 11 Then
        Debug.Print "ERROR: Array datosNuevos mal dimensionado"
        Debug.Print "  UBound(1): " & UBound(datosNuevos, 1) & " (esperado >= " & contadorNuevos & ")"
        Debug.Print "  UBound(2): " & UBound(datosNuevos, 2) & " (esperado >= 11)"
        Exit Sub
    End If
    
    ' --- NUEVA LOGICA: ASIGNACION EN BLOQUE ---
    
    ' PASO 1: Obtener datos existentes del ListBox de Trabajo
    Dim datosActuales() As Variant
    Dim contadorActuales As Long
    With frm.Listbox_Trabajo
        If .ListCount > 0 Then
            datosActuales = .List ' Esto es un array 0-based
            contadorActuales = .ListCount
        Else
            contadorActuales = 0
                    End If
    End With
    
    ' PASO 2: Crear el array final combinado
    Dim datosFinales() As Variant
    Dim contadorFinal As Long
    contadorFinal = contadorActuales + contadorNuevos
    ' El array debe ser Variant para asignarlo a .List
    ReDim datosFinales(0 To contadorFinal - 1, 0 To 10)
    
    ' PASO 3: Copiar datos existentes al array final
    If contadorActuales > 0 Then
        For i = 0 To contadorActuales - 1
            For j = 0 To 10
                datosFinales(i, j) = datosActuales(i, j)
            Next j
        Next i
    End If
    
    ' PASO 4: Copiar datos nuevos al array final
    ' datosNuevos es 1-based, datosFinales es 0-based. Se ajustan indices.
    For i = 1 To contadorNuevos
        For j = 1 To 11
            datosFinales(contadorActuales + i - 1, j - 1) = datosNuevos(i, j)
        Next j
    Next i
    
    ' PASO 5: Asignar el array final al ListBox de una sola vez
    With frm.Listbox_Trabajo
        .Clear
        If contadorFinal > 0 Then
            .List = datosFinales
        End If
    End With
    
    Debug.Print "Asignacion de datos al ListBox completada. Total de registros: " & contadorFinal
    
    ' --- FIN DE LA NUEVA LOGICA ---
    
    ' Actualizar consecutivos
    consActividadFinal = consecutivoActividad - 1
    Call ActualizarConsecutivos(area, CLng(consecutivoCapitulo), consActividadFinal)
    
    Debug.Print "Agregados " & contadorNuevos & " registros al trabajo"
    Debug.Print "Area: " & area & ", Capitulo: " & capitulo
    Debug.Print "Consecutivo Capitulo: " & consecutivoCapitulo & ", Consecutivo Actividad: " & (consActividadFinal)
    
    Exit Sub
ErrHandler:
    Debug.Print "Error en AgregarRegistrosATrabajo: " & Err.Description
End Sub

' Asignar cantidad a registros seleccionados en el ListBox de Trabajo
Public Sub AsignarCantidadATrabajo(frm As Object, cantidad As Double)
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim registrosSeleccionados As Long
    
    ' Contar registros seleccionados en ListBox de Trabajo
    registrosSeleccionados = 0
    With frm.Listbox_Trabajo
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                registrosSeleccionados = registrosSeleccionados + 1
            End If
        Next i
    End With
    
    If registrosSeleccionados = 0 Then
        MsgBox "Debe seleccionar al menos un registro en el área de trabajo para asignar cantidad", vbExclamation, "Validación"
        Exit Sub
    End If
    
    ' Asignar cantidad a registros seleccionados
    With frm.Listbox_Trabajo
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                ' Actualizar cantidad (columna 9, índice 8)
                .List(i, 8) = cantidad
                
                ' Recalcular VR. PARCIAL (columna 11, índice 10)
                Dim precioUnitario As Double
                Dim precioTexto As String
                precioTexto = CStr(.List(i, 9)) ' VR. UNITARIO (columna 10, índice 9)
                
                ' Limpiar formato de precio (quitar $ y comas)
                If Left(precioTexto, 1) = "$" Then
                    precioTexto = Mid(precioTexto, 2)
                End If
                precioTexto = Replace(precioTexto, ",", "")
                
                If IsNumeric(precioTexto) Then
                    precioUnitario = CDbl(precioTexto)
                Else
                    precioUnitario = 0
                End If
                .List(i, 10) = precioUnitario * cantidad
            End If
        Next i
    End With
    
    MsgBox "Se asignó la cantidad " & cantidad & " a " & registrosSeleccionados & " registro(s) seleccionado(s)", vbInformation, "Cantidad Asignada"
    Debug.Print "Cantidad asignada: " & cantidad & " a " & registrosSeleccionados & " registros"
    Exit Sub
ErrHandler:
    Debug.Print "Error en AsignarCantidadATrabajo: " & Err.Description
    MsgBox "Error al asignar cantidad: " & Err.Description, vbCritical, "Error"
End Sub

' Editar cantidad de un registro específico en el ListBox de Trabajo
Public Sub EditarCantidadTrabajo(frm As Object, filaSeleccionada As Long, nuevaCantidad As Double)
    On Error GoTo ErrHandler
    
    ' Actualizar cantidad
    frm.Listbox_Trabajo.List(filaSeleccionada, 8) = nuevaCantidad
    
    ' Recalcular VR. PARCIAL
    Dim precioUnitario As Double
    Dim precioTexto As String
    precioTexto = CStr(frm.Listbox_Trabajo.List(filaSeleccionada, 9)) ' VR. UNITARIO
    
    ' Limpiar formato de precio (quitar $ y comas)
    If Left(precioTexto, 1) = "$" Then
        precioTexto = Mid(precioTexto, 2)
    End If
    precioTexto = Replace(precioTexto, ",", "")
    
    If IsNumeric(precioTexto) Then
        precioUnitario = CDbl(precioTexto)
    Else
        precioUnitario = 0
    End If
    frm.Listbox_Trabajo.List(filaSeleccionada, 10) = precioUnitario * nuevaCantidad
    
    Debug.Print "Cantidad editada en fila " & (filaSeleccionada + 1) & ": " & nuevaCantidad
    Exit Sub
ErrHandler:
    Debug.Print "Error en EditarCantidadTrabajo: " & Err.Description
End Sub

' Eliminar registros seleccionados del ListBox de Trabajo
Public Sub EliminarRegistrosTrabajo(frm As Object)
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim registrosEliminados As Long
    
    ' Contar registros seleccionados
    registrosEliminados = 0
    With frm.Listbox_Trabajo
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                registrosEliminados = registrosEliminados + 1
            End If
        Next i
    End With
    
    If registrosEliminados = 0 Then
        MsgBox "No hay registros seleccionados para eliminar", vbExclamation, "Validación"
        Exit Sub
    End If
    
    ' Confirmar eliminación
    If MsgBox("¿Está seguro de que desea eliminar " & registrosEliminados & " registro(s) del área de trabajo?", vbQuestion + vbYesNo, "Confirmar Eliminación") = vbNo Then
        Exit Sub
    End If
    
    ' Eliminar registros de abajo hacia arriba para evitar problemas con índices
    With frm.Listbox_Trabajo
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                .RemoveItem i
            End If
        Next i
    End With
    
    MsgBox "Se eliminaron " & registrosEliminados & " registro(s) del área de trabajo", vbInformation, "Eliminación Completada"
    Debug.Print "Eliminados " & registrosEliminados & " registros del área de trabajo"
    Exit Sub
ErrHandler:
    Debug.Print "Error en EliminarRegistrosTrabajo: " & Err.Description
    MsgBox "Error durante la eliminación: " & Err.Description, vbCritical, "Error"
End Sub

' Validar datos antes de exportar
Public Function ValidarDatosParaExportar(frm As Object) As Boolean
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim errores As String
    Dim totalRegistros As Long
    Dim registrosConErrores As Long
    
    errores = ""
    totalRegistros = 0
    registrosConErrores = 0
    
    With frm.Listbox_Trabajo
        For i = 0 To .ListCount - 1
            totalRegistros = totalRegistros + 1
            
            ' Verificar campos obligatorios
            If .List(i, 1) = "" Then ' Área vacía
                errores = errores & "Fila " & (i + 1) & ": Área vacía" & vbCrLf
                registrosConErrores = registrosConErrores + 1
            End If
            
            If .List(i, 3) = "" Then ' Capítulo vacío
                errores = errores & "Fila " & (i + 1) & ": Capítulo vacío" & vbCrLf
                registrosConErrores = registrosConErrores + 1
            End If
            
            ' Validar cantidad específicamente
            If .List(i, 8) = "" Then ' Cantidad vacía
                errores = errores & "Fila " & (i + 1) & ": Cantidad vacía" & vbCrLf
                registrosConErrores = registrosConErrores + 1
            ElseIf Not IsNumeric(.List(i, 8)) Then ' Cantidad no numérica
                errores = errores & "Fila " & (i + 1) & ": Cantidad no es un número válido" & vbCrLf
                registrosConErrores = registrosConErrores + 1
            ElseIf CDbl(.List(i, 8)) <= 0 Then ' Cantidad menor o igual a 0
                errores = errores & "Fila " & (i + 1) & ": Cantidad debe ser mayor a 0" & vbCrLf
                registrosConErrores = registrosConErrores + 1
            End If
        Next i
    End With
    
    If errores <> "" Then
        Dim mensajeError As String
        mensajeError = "Se encontraron errores en " & registrosConErrores & " de " & totalRegistros & " registros:" & vbCrLf & vbCrLf & errores & vbCrLf & "Use el botón 'Asignar Cantidad' para corregir las cantidades."
        MsgBox mensajeError, vbExclamation, "Validación de Datos"
        ValidarDatosParaExportar = False
    Else
        ValidarDatosParaExportar = True
    End If
    
    Exit Function
ErrHandler:
    Debug.Print "Error en ValidarDatosParaExportar: " & Err.Description
    ValidarDatosParaExportar = False
End Function

' Función auxiliar para extraer datos del ListBox de forma segura
Private Function ExtraerDatoListBox(listBox As Object, fila As Long, columna As Long) As String
    On Error Resume Next
    
    Dim valor As Variant
    valor = ""
    
    ' Validar índices
    If fila < 0 Or fila >= listBox.ListCount Then
        Debug.Print "  ERROR: Índice de fila fuera de rango: " & fila & " (ListCount: " & listBox.ListCount & ")"
        ExtraerDatoListBox = ""
        Exit Function
    End If
    
    If columna < 0 Or columna >= listBox.ColumnCount Then
        Debug.Print "  ERROR: Índice de columna fuera de rango: " & columna & " (ColumnCount: " & listBox.ColumnCount & ")"
        ExtraerDatoListBox = ""
        Exit Function
    End If
    
    ' Intentar extraer el valor
    valor = listBox.List(fila, columna)
    
    ' Verificar si hubo error
    If Err.Number <> 0 Then
        Debug.Print "  ERROR: No se pudo acceder a .List(" & fila & ", " & columna & ") - " & Err.Description
        Err.Clear
        ExtraerDatoListBox = ""
        Exit Function
    End If
    
    ' Convertir a string de forma segura
    If IsEmpty(valor) Or IsNull(valor) Then
        ExtraerDatoListBox = ""
    Else
        ExtraerDatoListBox = CStr(valor)
    End If
    
    On Error GoTo 0
End Function

' === FUNCIONES DE VALIDACIÓN ===

' Validar que el ListBox esté correctamente configurado
Public Function ValidarListBoxRegistros(frm As Object) As Boolean
    On Error GoTo ErrHandler
    
    If frm Is Nothing Then
        Debug.Print "ERROR: El formulario es Nothing"
        ValidarListBoxRegistros = False
        Exit Function
    End If
    
    If frm.Listbox_Registros Is Nothing Then
        Debug.Print "ERROR: Listbox_Registros es Nothing"
        ValidarListBoxRegistros = False
        Exit Function
    End If
    
    With frm.Listbox_Registros
        Debug.Print "=== VALIDACIÓN DEL LISTBOX ==="
        Debug.Print "ListCount: " & .ListCount
        Debug.Print "ColumnCount: " & .ColumnCount
        Debug.Print "MultiSelect: " & .MultiSelect
        Debug.Print "ListStyle: " & .ListStyle
        Debug.Print "============================="
        
        ' Validar que tenga datos
        If .ListCount = 0 Then
            Debug.Print "ERROR: ListBox está vacío"
            ValidarListBoxRegistros = False
            Exit Function
        End If
        
            ' Validar que tenga suficientes columnas (ahora son 4)
    If .ColumnCount < 4 Then
        Debug.Print "ERROR: ListBox no tiene suficientes columnas. ColumnCount=" & .ColumnCount
        ValidarListBoxRegistros = False
        Exit Function
    End If
        
        ' Validar que se pueda acceder a los datos
        On Error Resume Next
        Dim testValue As Variant
        testValue = .List(0, 0)
        If Err.Number <> 0 Then
            Debug.Print "ERROR: No se puede acceder a los datos del ListBox - " & Err.Description
            Err.Clear
            ValidarListBoxRegistros = False
            Exit Function
        End If
        On Error GoTo ErrHandler
        
        ValidarListBoxRegistros = True
    End With
    
    Exit Function
ErrHandler:
    Debug.Print "Error en ValidarListBoxRegistros: " & Err.Description
    ValidarListBoxRegistros = False
End Function

' Función para reinicializar el ListBox si es necesario
Public Sub ReinicializarListBoxRegistros(frm As Object)
    On Error GoTo ErrHandler
    
    Debug.Print "Reinicializando ListBox_Registros..."
    
    ' Limpiar y reconfigurar el ListBox
    With frm.Listbox_Registros
        .Clear
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .ColumnCount = 7
        .ColumnWidths = "40 pt;40 pt;40 pt;60 pt;220 pt;60 pt;60 pt"
    End With
    
    ' Recargar datos
    Call FiltrarYCargarListBox(frm)
    
    Debug.Print "ListBox_Registros reinicializado correctamente"
    Exit Sub
ErrHandler:
    Debug.Print "Error en ReinicializarListBoxRegistros: " & Err.Description
End Sub

' Validar que el ListBox de Trabajo esté correctamente configurado
Public Function ValidarListBoxTrabajo(frm As Object) As Boolean
    On Error GoTo ErrHandler
    
    Debug.Print "=== VALIDACIÓN DEL LISTBOX_TRABAJO ==="
    
    If frm Is Nothing Then
        Debug.Print "ERROR: El formulario es Nothing"
        ValidarListBoxTrabajo = False
        Exit Function
    End If
    
    If frm.Listbox_Trabajo Is Nothing Then
        Debug.Print "ERROR: Listbox_Trabajo es Nothing"
        ValidarListBoxTrabajo = False
        Exit Function
    End If
    
    With frm.Listbox_Trabajo
        Debug.Print "ListCount: " & .ListCount
        Debug.Print "ColumnCount: " & .ColumnCount
        Debug.Print "MultiSelect: " & .MultiSelect
        Debug.Print "ListStyle: " & .ListStyle
        Debug.Print "Width: " & .Width
        Debug.Print "Height: " & .Height
        Debug.Print "Visible: " & .Visible
        Debug.Print "Enabled: " & .Enabled
        
        ' Validar que tenga suficientes columnas
        If .ColumnCount < 11 Then
            Debug.Print "ERROR: ListBox_Trabajo no tiene suficientes columnas. ColumnCount=" & .ColumnCount
            ValidarListBoxTrabajo = False
            Exit Function
        End If
        
        ' Validar que se pueda acceder a los datos
        On Error Resume Next
        Dim testValue As Variant
        If .ListCount > 0 Then
            testValue = .List(0, 0)
            If Err.Number <> 0 Then
                Debug.Print "ERROR: No se puede acceder a los datos del ListBox_Trabajo - " & Err.Description
                Err.Clear
                ValidarListBoxTrabajo = False
                Exit Function
            End If
        End If
        On Error GoTo ErrHandler
        
        ' Validar que se pueda agregar una fila
        On Error Resume Next
        .AddItem ""
        If Err.Number <> 0 Then
            Debug.Print "ERROR: No se puede agregar fila al ListBox_Trabajo - " & Err.Description
            Err.Clear
            ValidarListBoxTrabajo = False
            Exit Function
        End If
        ' Remover la fila de prueba
        If .ListCount > 0 Then
            .RemoveItem 0
        End If
        On Error GoTo ErrHandler
        
        ValidarListBoxTrabajo = True
    End With
    
    Debug.Print "ListBox_Trabajo está correctamente configurado"
    Debug.Print "============================================="
    
    Exit Function
ErrHandler:
    Debug.Print "Error en ValidarListBoxTrabajo: " & Err.Description
    ValidarListBoxTrabajo = False
End Function

' Función para reinicializar el ListBox de Trabajo si es necesario
Public Sub ReinicializarListBoxTrabajo(frm As Object)
    On Error GoTo ErrHandler
    
    Debug.Print "Reinicializando ListBox_Trabajo..."
    
    ' Limpiar y reconfigurar el ListBox
    With frm.Listbox_Trabajo
        .Clear
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .ColumnCount = 11
        .ColumnWidths = "40 pt;40 pt;40 pt;120 pt;40 pt;60 pt;220 pt;40 pt;60 pt;80 pt;80 pt"
    End With
    
    Debug.Print "ListBox_Trabajo reinicializado correctamente"
    Exit Sub
ErrHandler:
    Debug.Print "Error en ReinicializarListBoxTrabajo: " & Err.Description
End Sub

'**************************************************************************************************
'*** NUEVAS FUNCIONES PARA MANEJAR ACTIVIDADES DUPLICADAS (2024-07-29) ***
'**************************************************************************************************

' Función principal para buscar y gestionar una actividad que ya podría existir.
Private Function ManejarActividadExistente(frm As Object, area As String, capitulo As String, codigoActividad As String, descripcionActividad As String, cantidadASumar As Double) As Boolean
    On Error GoTo ErrHandler
    ManejarActividadExistente = False ' Por defecto, la actividad es nueva

    Dim i As Long
    Dim respuesta As VbMsgBoxResult
    Dim cantidadExistente As Double
    Dim precioUnitario As Double
    
    ' --- PASO 1: Buscar en la Lista de Trabajo (lo que está en el formulario) ---
    With frm.Listbox_Trabajo
        If .ListCount > 0 Then
            For i = 0 To .ListCount - 1
                ' Columnas: 2=Area, 4=Capitulo, 6=Codigo Actividad
                If .List(i, 1) = area And .List(i, 3) = capitulo And .List(i, 5) = codigoActividad Then
                    respuesta = MsgBox("La actividad '" & descripcionActividad & "' ya existe en la Lista de Trabajo." & vbCrLf & vbCrLf & _
                                       "¿Desea sumar la nueva cantidad (" & cantidadASumar & ") a la ya existente?", _
                                       vbQuestion + vbYesNo, "Actividad Duplicada")
                    
                    If respuesta = vbYes Then
                        cantidadExistente = CDbl(.List(i, 8))
                        .List(i, 8) = cantidadExistente + cantidadASumar ' Actualizar Cantidad (Col 9)
                        
                        precioUnitario = CDbl(Replace(Replace(.List(i, 9), "$", ""), ",", ""))
                        .List(i, 10) = .List(i, 8) * precioUnitario ' Actualizar Vr. Parcial (Col 11)
                        
                        Debug.Print "CANTIDAD ACTUALIZADA en ListBox para actividad '" & codigoActividad & "'"
                    End If
                    
                    ManejarActividadExistente = True ' Actividad encontrada y manejada (se haya sumado o no)
                    Exit Function
                End If
            Next i
        End If
    End With
    
    ' --- PASO 2: Si no está en la lista, buscar en la Hoja de Excel (lo ya exportado) ---
    Dim ws As Worksheet
    Set ws = ObtenerHojaDestino()
    
    If Not ws Is Nothing Then
        For i = 2 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            ' Columnas: 2=AREA, 4=CAPITULO, 6=CODIGO ACTIVIDAD
            If ws.Cells(i, 2).Value = area And ws.Cells(i, 4).Value = capitulo And ws.Cells(i, 6).Value = codigoActividad Then
                respuesta = MsgBox("La actividad '" & descripcionActividad & "' ya existe en la hoja 'Acta-Presupuesto' (ya fue exportada)." & vbCrLf & vbCrLf & _
                                   "¿Desea sumar la nueva cantidad (" & cantidadASumar & ") a la ya existente?", _
                                   vbQuestion + vbYesNo, "Actividad Duplicada")
                
                If respuesta = vbYes Then
                    cantidadExistente = CDbl(ws.Cells(i, 9).Value)
                    ws.Cells(i, 9).Value = cantidadExistente + cantidadASumar ' Actualizar Cantidad (Col 9)
                    
                    precioUnitario = CDbl(ws.Cells(i, 10).Value)
                    ws.Cells(i, 11).Value = ws.Cells(i, 9).Value * precioUnitario ' Actualizar Vr. Parcial (Col 11)
                    
                    Debug.Print "CANTIDAD ACTUALIZADA en Hoja para actividad '" & codigoActividad & "'"
                End If

                ManejarActividadExistente = True ' Actividad encontrada y manejada
                Exit Function
            End If
        Next i
    End If
    
    Exit Function
ErrHandler:
    Debug.Print "Error en ManejarActividadExistente: " & Err.Description
    ManejarActividadExistente = False ' En caso de error, se asume que no existe para no bloquear al usuario.
End Function

' Función auxiliar para obtener la hoja de destino.
Private Function ObtenerHojaDestino() As Worksheet
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Acta-Presupuesto" Then
            Set ObtenerHojaDestino = ws
            Exit Function
        End If
    Next ws
    
    ' Si no existe, no la crea, solo devuelve Nothing.
    ' La creación se maneja en el módulo de exportación.
    Set ObtenerHojaDestino = Nothing
    Exit Function
ErrHandler:
    Debug.Print "Error en ObtenerHojaDestino (Módulo Trabajo): " & Err.Description
    Set ObtenerHojaDestino = Nothing
End Function
