Attribute VB_Name = "Modulo_Exportacion"

' === MÓDULO PARA GESTIÓN DE EXPORTACIÓN === 7:43 PM 22/08/2025

' Exportar datos del ListBox de Trabajo a la hoja "Acta-Presupuesto"
Public Sub ExportarDatosATrabajo(frm As Object)
    On Error GoTo ErrHandler
    
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim filaDestino As Long
    Dim registrosExportados As Long
    Dim listData As Variant
    
    ' Validar datos antes de exportar
    If Not ValidarDatosParaExportar(frm) Then Exit Sub
    
    ' Obtener hoja destino
    Set ws = ObtenerHojaDestino(True)
    If ws Is Nothing Then
        MsgBox "No se pudo obtener o crear la hoja de destino.", vbCritical, "Error de Exportación"
        Exit Sub
    End If
    
    ' --- CAMBIO CLAVE: Volcar el ListBox a un array para estabilidad ---
    If frm.Listbox_Trabajo.ListCount = 0 Then
        Call RegistrarLog("Nada que exportar, ListBox de trabajo está vacío.")
        Exit Sub
    End If
    listData = frm.Listbox_Trabajo.List
    
    ' Encontrar la siguiente fila disponible en la hoja
    filaDestino = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call RegistrarLog("Iniciando exportación a fila: " & filaDestino)
    
    ' Exportar datos desde el array
    registrosExportados = 0
    For i = LBound(listData, 1) To UBound(listData, 1)
        ' --- LÓGICA PARA SEPARAR EL CONSECUTIVO CONCATENADO ---
        Dim codigoCompleto As String
        Dim partes() As String
        Dim consecutivoArea As String
        
        codigoCompleto = CStr(listData(i, 0)) ' Columna 1: Código combinado (ej: "1.2.14")
        
        If InStr(codigoCompleto, ".") > 0 Then
            partes = Split(codigoCompleto, ".")
            consecutivoArea = partes(0) ' Extraer solo la primera parte (el "1")
        Else
            consecutivoArea = codigoCompleto ' Fallback por si no tiene el formato esperado
        End If
        ' --- FIN DE LA LÓGICA ---

        registrosExportados = registrosExportados + 1
        
        For j = LBound(listData, 2) To UBound(listData, 2)
            Dim valorParaExportar As Variant
            
            ' Decidir qué valor exportar a la hoja de Excel
            If j = 0 Then
                ' Para la primera columna (CONSECUTIVO AREA), usar el número extraído.
                valorParaExportar = consecutivoArea
            Else
                ' Para todas las demás columnas, usar el valor que ya está en la lista.
                valorParaExportar = listData(i, j)
                
                ' Log para columnas de precio (9 y 10 son VR. UNITARIO y VR. PARCIAL)
                If j = 9 Or j = 10 Then
                    Call RegistrarLog("=== EXPORTACIÓN DE PRECIO ===")
                    Call RegistrarLog("Columna " & j + 1 & ": " & IIf(j = 9, "VR. UNITARIO", "VR. PARCIAL"))
                    Call RegistrarLog("Valor Original: " & valorParaExportar)
                End If
            End If
            
            ' Limpiar formato de moneda para las columnas de precio antes de exportar
            If j = 9 Or j = 10 Then ' Col 10 (VR. UNITARIO) y Col 11 (VR. PARCIAL)
                Dim precioTexto As String
                precioTexto = CStr(valorParaExportar)
                
                Call RegistrarLog("=== PROCESAMIENTO DE PRECIO ===")
                Call RegistrarLog("Precio inicial: " & precioTexto)
                
                ' 1. Quitar el símbolo de moneda si existe
                If Left(precioTexto, 1) = "$" Then
                    precioTexto = Mid(precioTexto, 2)
                End If
                Call RegistrarLog("Sin símbolo $: " & precioTexto)
                
                ' 2. Quitar espacios en blanco
                precioTexto = Trim(precioTexto)
                
                ' 3. Extraer el número puro
                Dim numeroStr As String
                numeroStr = Replace(Replace(precioTexto, ".", ""), ",", "")
                
                ' 4. Determinar si es un precio con decimales
                Dim tieneDecimales As Boolean
                tieneDecimales = (InStr(precioTexto, ",") > 0) Or (InStrRev(precioTexto, ".") = Len(precioTexto) - 2)
                
                ' 5. Si tiene decimales, dividir por 100
                If tieneDecimales Then
                    precioTexto = CDbl(numeroStr) / 100
                Else
                    precioTexto = CDbl(numeroStr)
                End If
                
                Call RegistrarLog("Número procesado: " & precioTexto)
                
                If IsNumeric(precioTexto) Then
                    valorParaExportar = CDbl(precioTexto)
                    
                    ' Log del proceso de transformación
                    If j = 9 Or j = 10 Then
                        Call RegistrarLog("Después de limpiar $: " & precioTexto)
                        Call RegistrarLog("Valor Final: " & valorParaExportar)
                        Call RegistrarLog("==============================")
                    End If
                End If
            End If
            
            ' Escribir el valor final en la celda
            ws.Cells(filaDestino, j + 1).Value = valorParaExportar
            
            ' Log del valor escrito en Excel (solo para precios)
            If j = 9 Or j = 10 Then
                Call RegistrarLog("=== VALOR ESCRITO EN EXCEL ===")
                Call RegistrarLog("Columna " & j + 1 & ": " & ws.Cells(filaDestino, j + 1).Value)
                Call RegistrarLog("==============================")
            End If
        Next j
        
        filaDestino = filaDestino + 1
    Next i
    
    ' Limpiar ListBox de Trabajo después de exportar
    frm.Listbox_Trabajo.Clear
    
    ' Reconfigurar ListBox de Trabajo
    Call ConfigurarListBoxTrabajo_Principal(frm)
    
    MsgBox "Se exportaron " & registrosExportados & " registros a la hoja 'Acta-Presupuesto'", vbInformation, "Exportación Completada"
    Call RegistrarLog("Exportación completada: " & registrosExportados & " registros")
    Exit Sub
ErrHandler:
    Call RegistrarLog("Error en ExportarDatosATrabajo: " & Err.Description)
    MsgBox "Error durante la exportación: " & Err.Description, vbCritical, "Error de Exportación"
End Sub

' Cargar datos exportados en el ListBox de Revisión
Public Sub CargarDatosExportadosEnRevision(frm As Object)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim dataRange As Range
    Dim dataArray As Variant
    
    ' Obtener hoja destino
    Set ws = ObtenerHojaDestino()
    If ws Is Nothing Then
        Call RegistrarLog("No hay hoja de datos exportados para cargar.")
        frm.Listbox_Exportados.Clear
        Call ConfigurarListBoxExportados(frm) ' Reconfigurar por si tenía datos
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Si solo hay encabezados o está vacía, limpiar y salir.
    If lastRow <= 1 Then
        Call RegistrarLog("La hoja de datos exportados está vacía.")
        frm.Listbox_Exportados.Clear
        Call ConfigurarListBoxExportados(frm)
        Exit Sub
    End If
    
    ' --- CAMBIO CLAVE: Volcar el rango de Excel a un array ---
    Set dataRange = ws.Range("A2:K" & lastRow) ' Desde la fila 2 hasta la última
    dataArray = dataRange.Value
    
    ' Crear un nuevo array para los datos que se mostrarán en el ListBox
    Dim listData() As Variant
    Dim numRows As Long, numCols As Long
    numRows = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    numCols = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    ReDim listData(1 To numRows, 1 To numCols)
    
    ' Procesar el array para crear el código combinado y formatear precios
    Dim filaDestino As Long
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        filaDestino = i - LBound(dataArray, 1) + 1
        
        ' --- LÓGICA PARA CREAR EL CONSECUTIVO CONCATENADO ---
        Dim consecArea As String, consecCapitulo As String, consecActividad As String
        consecArea = CStr(dataArray(i, 1))
        consecCapitulo = CStr(dataArray(i, 3))
        consecActividad = CStr(dataArray(i, 5))
        listData(filaDestino, 1) = consecArea & "." & consecCapitulo & "." & consecActividad
        ' --- FIN DE LA LÓGICA ---
        
        ' Copiar el resto de los datos
        For j = 2 To numCols
            listData(filaDestino, j) = dataArray(i, j)
        Next j
        
        ' Formatear columnas de precio como moneda para visualización
        If IsNumeric(listData(filaDestino, 10)) Then
            listData(filaDestino, 10) = Format(listData(filaDestino, 10), "$#,##0.00")
        End If
        If IsNumeric(listData(filaDestino, 11)) Then
            listData(filaDestino, 11) = Format(listData(filaDestino, 11), "$#,##0.00")
        End If
    Next i
    
    ' Cargar el array procesado en el ListBox
    With frm.Listbox_Exportados
        .Clear
        .List = listData
    End With
    
    Call RegistrarLog("Carga completada: " & numRows & " registros en ListBox de Exportados")
    
    ' Limpiar ListBox de Trabajo después de exportar
    frm.Listbox_Trabajo.Clear
    
    ' Reconfigurar ListBox de Trabajo
    Call ConfigurarListBoxTrabajo_Principal(frm)
    
    Exit Sub
ErrHandler:
    Call RegistrarLog("Error en CargarDatosExportadosEnRevision: " & Err.Description)
    frm.Listbox_Exportados.Clear
    Call ConfigurarListBoxExportados(frm)
End Sub

' Mostrar estadísticas de datos exportados
Public Sub MostrarEstadisticasExportados(frm As Object)
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim totalRegistros As Long
    Dim totalValor As Double
    Dim areas As Object, capitulos As Object
    Dim listData As Variant
    
    ' Volcar ListBox a un array para estabilidad
    If frm.Listbox_Exportados.ListCount = 0 Then
        Call RegistrarLog("No hay datos exportados para mostrar estadísticas.")
        Exit Sub
    End If
    listData = frm.Listbox_Exportados.List
    
    Set areas = CreateObject("Scripting.Dictionary")
    Set capitulos = CreateObject("Scripting.Dictionary")
    totalValor = 0
    
    For i = LBound(listData, 1) To UBound(listData, 1)
        ' Columna 2 (índice 1) es Área
        If Not IsEmpty(listData(i, 1)) And listData(i, 1) <> "" Then
            ' Contar áreas únicas
            If Not areas.Exists(listData(i, 1)) Then areas.Add listData(i, 1), 1
            
            ' Contar capítulos únicos (Columna 4, índice 3)
            If Not capitulos.Exists(listData(i, 3)) Then capitulos.Add listData(i, 3), 1
            
            ' Sumar valor parcial (Columna 11, índice 10)
            Dim valorParcial As String
            valorParcial = CStr(listData(i, 10))
            ' Procesar el formato del precio correctamente
            valorParcial = Replace(valorParcial, "$", "")  ' Quitar el símbolo de moneda
            ' 1. Primero quitamos los separadores de miles (puntos)
            valorParcial = Replace(valorParcial, ".", "")
            ' 2. Reemplazamos la coma decimal por punto
            valorParcial = Replace(valorParcial, ",", ".")
            ' 3. Si termina en .0, .00, etc., lo dividimos por la potencia de 10 correspondiente
            If InStr(valorParcial, ".") > 0 Then
                Dim decimalesValor As Integer
                decimalesValor = Len(valorParcial) - InStr(valorParcial, ".") ' Número de decimales
                If decimalesValor > 0 Then
                    valorParcial = CDbl(valorParcial) / (10 ^ decimalesValor)
                End If
            End If
            If IsNumeric(valorParcial) Then totalValor = totalValor + CDbl(valorParcial)
        End If
    Next i
    
    totalRegistros = areas.Count
    
    ' Mostrar estadísticas
    Call RegistrarLog("=== ESTADÍSTICAS DE DATOS EXPORTADOS ===")
    Call RegistrarLog("Total de registros: " & totalRegistros)
    Call RegistrarLog("Total de áreas únicas: " & areas.Count)
    Call RegistrarLog("Total de capítulos únicos: " & capitulos.Count)
    Call RegistrarLog("Valor total: $" & Format(totalValor, "#,##0.00"))
    Call RegistrarLog("=========================================")
    
    ' Mostrar áreas únicas
    Call RegistrarLog("Áreas encontradas:")
    For i = 0 To areas.Count - 1
        Call RegistrarLog("  - " & areas.Keys()(i))
    Next i
    
    ' Mostrar capítulos únicos
    Call RegistrarLog("Capítulos encontrados:")
    For i = 0 To capitulos.Count - 1
        Call RegistrarLog("  - " & capitulos.Keys()(i))
    Next i
    
    Exit Sub
ErrHandler:
    Call RegistrarLog("Error en MostrarEstadisticasExportados: " & Err.Description)
End Sub

' Función auxiliar para obtener hoja destino
Private Function ObtenerHojaDestino(Optional crearSiNoExiste As Boolean = False) As Worksheet
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    ' Buscar la hoja "Acta-Presupuesto"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Acta-Presupuesto" Then
            Set ObtenerHojaDestino = ws
            Exit Function
        End If
    Next ws
    
    ' Si no existe y se solicita crear
    If crearSiNoExiste Then
        If MsgBox("La hoja 'Acta-Presupuesto' no existe. ¿Desea crearla?", vbQuestion + vbYesNo, "Hoja No Encontrada") = vbYes Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = "Acta-Presupuesto"
            
            ' Crear encabezados
            ws.Cells(1, 1).Value = "CONSECUTIVO AREA"
            ws.Cells(1, 2).Value = "AREA"
            ws.Cells(1, 3).Value = "CONSECUTIVO CAPITULO"
            ws.Cells(1, 4).Value = "DESCRIPCION CAPITULO"
            ws.Cells(1, 5).Value = "CONSECUTIVO ACTIVIDAD"
            ws.Cells(1, 6).Value = "CODIGO ACTIVIDAD"
            ws.Cells(1, 7).Value = "ACTIVIDAD"
            ws.Cells(1, 8).Value = "UND"
            ws.Cells(1, 9).Value = "CANTIDAD"
            ws.Cells(1, 10).Value = "VR. UNITARIO"
            ws.Cells(1, 11).Value = "VR. PARCIAL"
            
            Set ObtenerHojaDestino = ws
        Else
            Set ObtenerHojaDestino = Nothing
        End If
    Else
        ' Solo consultar, no crear
        Set ObtenerHojaDestino = Nothing
    End If
    
    Exit Function
ErrHandler:
    Call RegistrarLog("Error en ObtenerHojaDestino: " & Err.Description)
    Set ObtenerHojaDestino = Nothing
End Function

' === FUNCIONES CRUD PARA DATOS EXPORTADOS (CON SEGURIDAD) ===

' Modifica la cantidad de un registro ya exportado, previa autenticación.
Public Sub ModificarCantidadExportada(frm As Object, filaSeleccionada As Long)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim filaEnHoja As Long
    Dim nuevaCantidadStr As String, nuevaCantidad As Double
    
    ' PASO 1: Pedir la nueva cantidad al usuario
    nuevaCantidadStr = InputBox("Ingrese la nueva cantidad para el registro seleccionado:", "Modificar Cantidad Exportada")
    
    ' Validar entrada
    If nuevaCantidadStr = "" Then Exit Sub ' Usuario canceló
    If Not IsNumeric(nuevaCantidadStr) Then
        MsgBox "Por favor, ingrese un valor numerico valido.", vbExclamation, "Entrada no valida"
        Exit Sub
    End If
    nuevaCantidad = CDbl(nuevaCantidadStr)
    
    ' PASO 2: Identificar el registro en la hoja de Excel
    Set ws = ObtenerHojaDestino()
    If ws Is Nothing Then Exit Sub ' La hoja no existe
    
    filaEnHoja = EncontrarFilaEnHoja(frm.Listbox_Exportados, filaSeleccionada, ws)
    
    If filaEnHoja = 0 Then
        MsgBox "No se pudo encontrar el registro seleccionado en la hoja 'Acta-Presupuesto'.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' PASO 3: Actualizar los datos en Excel
    With ws
        .Cells(filaEnHoja, 9).Value = nuevaCantidad ' Actualizar Cantidad (Col 9)
        
        Dim precioUnitario As Double
        precioUnitario = .Cells(filaEnHoja, 10).Value
        .Cells(filaEnHoja, 11).Value = nuevaCantidad * precioUnitario ' Recalcular Vr. Parcial (Col 11)
    End With
    
    ' PASO 4: Refrescar la vista
    Call CargarDatosExportadosEnRevision(frm)
    MsgBox "El registro ha sido actualizado correctamente.", vbInformation, "Operacion Exitosa"
    
    Exit Sub
ErrHandler:
    Call RegistrarLog("Error en ModificarCantidadExportada: " & Err.Description)
End Sub

' Elimina un registro ya exportado, previa autenticacion.
Public Sub EliminarRegistroExportado(frm As Object, filaSeleccionada As Long)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim filaEnHoja As Long
    
    ' NOTA: La confirmacion ahora se hace en el formulario frm_Accion
    ' antes de llamar a esta subrutina.
    
    ' PASO 1: Identificar el registro en la hoja de Excel
    Set ws = ObtenerHojaDestino()
    If ws Is Nothing Then Exit Sub ' La hoja no existe
    
    filaEnHoja = EncontrarFilaEnHoja(frm.Listbox_Exportados, filaSeleccionada, ws)
    
    If filaEnHoja = 0 Then
        MsgBox "No se pudo encontrar el registro seleccionado en la hoja 'Acta-Presupuesto'.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' PASO 2: Eliminar la fila en Excel
    ws.Rows(filaEnHoja).Delete
    
    ' PASO 3: Refrescar la vista
    Call CargarDatosExportadosEnRevision(frm)
    MsgBox "El registro ha sido eliminado correctamente.", vbInformation, "Operacion Exitosa"
    
    Exit Sub
ErrHandler:
    Call RegistrarLog("Error en EliminarRegistroExportado: " & Err.Description)
End Sub

' Función auxiliar para encontrar la fila correspondiente en la hoja de Excel
Private Function EncontrarFilaEnHoja(listBox As Object, filaSeleccionada As Long, ws As Worksheet) As Long
    On Error GoTo ErrHandler
    
    Dim i As Long
    ' Obtener identificadores únicos del registro en el ListBox
    Dim codigoActividad As String: codigoActividad = listBox.List(filaSeleccionada, 5) ' Col 6: CODIGO ACTIVIDAD
    Dim descCapitulo As String: descCapitulo = listBox.List(filaSeleccionada, 3)  ' Col 4: DESCRIPCION CAPITULO
    Dim area As String: area = listBox.List(filaSeleccionada, 1) ' Col 2: AREA
    
    ' Recorrer la hoja para encontrar la fila que coincida
    For i = 2 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If ws.Cells(i, 6).Value = codigoActividad And _
           ws.Cells(i, 4).Value = descCapitulo And _
           ws.Cells(i, 2).Value = area Then
            
            EncontrarFilaEnHoja = i
            Exit Function
        End If
    Next i
    
    ' Si no se encuentra, devuelve 0
    EncontrarFilaEnHoja = 0
    Exit Function
    
ErrHandler:
    Call RegistrarLog("Error en EncontrarFilaEnHoja: " & Err.Description)
    EncontrarFilaEnHoja = 0
End Function
