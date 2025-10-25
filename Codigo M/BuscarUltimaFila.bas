'Attribute VB_Name = "BuscarUltimaFila"
Option Explicit

' ============================================
' FUNCIÓN: LogOperacion
' Registra operaciones para validar el proceso
' ============================================

Sub LogOperacion(mensaje As String)
    ' Escribir mensaje con timestamp en Debug.Print
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & mensaje
End Sub

' ============================================
' FUNCIÓN: UltimaFilaVacia
' Busca la última fila vacía de una hoja, considerando TODAS las columnas
' No importa si hay filas vacías intermedias, busca realmente la última con datos
' ============================================

Function UltimaFilaVacia(ws As Worksheet) As Long
    On Error Resume Next
    
    Dim ultimaCelda As Range
    Dim ultimaFila As Long
    
    Set ultimaCelda = ws.Cells.Find(What:="*", _
                                     After:=ws.Range("A1"), _
                                     LookIn:=xlFormulas, _
                                     LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlPrevious)
    
    If Not ultimaCelda Is Nothing Then
        ultimaFila = ultimaCelda.Row + 1
    Else
        ultimaFila = 1
    End If
    
    UltimaFilaVacia = ultimaFila
    
    On Error GoTo 0
End Function


Sub Ubicartotales()
    Dim ws As Worksheet
    Dim filaVacia As Long
    Dim filaInicio As Long
    Dim resultado As Boolean
    
    ' Iniciar log
    LogOperacion "=== INICIO Ubicartotales ==="
    
    ' Desactivar actualización de pantalla para mejor rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("VISTA_CLIENTE")
    LogOperacion "Hoja VISTA_CLIENTE seleccionada"
    
    filaVacia = UltimaFilaVacia(ws)
    LogOperacion "Última fila vacía encontrada: " & filaVacia
    
    ' Llama automáticamente a la función Copiar
    filaInicio = Copiar()
    LogOperacion "Función Copiar ejecutada. Fila inicio: " & filaInicio
    
    If filaInicio > 0 Then
        LogOperacion "Iniciando formateo de tabla en fila: " & filaInicio
        ' Aplica formato a la tabla
        resultado = FormatearTabla(ws, filaInicio)
        LogOperacion "Formateo de tabla completado. Resultado: " & resultado
        
        ' Agrega fila TOTAL con sumas de columnas VR
        If resultado Then
            LogOperacion "Iniciando agregado de fila TOTAL"
            resultado = AgregarFilaTotal(ws, filaInicio)
            LogOperacion "Fila TOTAL agregada. Resultado: " & resultado
            
            ' Crear tabla resumen de actas
            If resultado Then
                LogOperacion "Iniciando creación de tabla resumen"
                resultado = CrearTablaResumen(ws, filaInicio)
                LogOperacion "Tabla resumen creada. Resultado: " & resultado
            End If
        End If
        
        ' Activar la hoja donde se pegó
        ws.Activate
        ws.Cells(filaInicio, 1).Select
        LogOperacion "Hoja activada y celda seleccionada: " & filaInicio & ",1"
    Else
        LogOperacion "ERROR: No se pudo copiar datos"
        MsgBox "Error al pegar sub totales", vbCritical
    End If
    
    ' Rehabilitar actualización de pantalla
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    LogOperacion "=== FIN Ubicartotales ==="
    
    Exit Sub
    
ErrorHandler:
    ' Rehabilitar actualización de pantalla en caso de error
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    LogOperacion "ERROR en Ubicartotales: " & Err.Description
    MsgBox "Error en Ubicartotales: " & Err.Description, vbCritical
End Sub


Function Copiar() As Long
    Dim wsDestino As Worksheet
    Dim tablaTOTALES As ListObject
    Dim rngDatos As Range
    Dim filaVacia As Long
    Dim filaDestino As Long
    
    On Error GoTo ErrorHandler
    
    LogOperacion "=== INICIO Función Copiar ==="
    
    ' Establece la hoja destino
    Set wsDestino = ThisWorkbook.Worksheets("VISTA_CLIENTE")
    LogOperacion "Hoja destino VISTA_CLIENTE establecida"
    
    ' Verifica si la hoja TOTALES está oculta y la hace visible temporalmente
    Dim wsOrigen As Worksheet
    Dim hojaOculta As Boolean
    
    ' Verifica que la hoja TOTALES existe
    On Error Resume Next
    Set wsOrigen = ThisWorkbook.Worksheets("TOTALES")
    If wsOrigen Is Nothing Then
        LogOperacion "ERROR: No se encontró la hoja 'TOTALES'"
        MsgBox "Error: No se encontró la hoja 'TOTALES'", vbCritical
        Copiar = 0
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    LogOperacion "Hoja TOTALES encontrada"
    
    hojaOculta = (wsOrigen.Visible = xlSheetHidden Or wsOrigen.Visible = xlSheetVeryHidden)
    
    If hojaOculta Then
        wsOrigen.Visible = xlSheetVisible
        LogOperacion "Hoja TOTALES estaba oculta, ahora visible"
    End If
    
    ' Encuentra la tabla TOTALES (consulta de Power Query)
    Set tablaTOTALES = wsOrigen.ListObjects("TOTALES")
    LogOperacion "Tabla TOTALES encontrada"
    
    ' Obtiene solo el rango de datos (sin encabezados)
    Set rngDatos = tablaTOTALES.DataBodyRange
    LogOperacion "Rango de datos obtenido"
    
    ' Verifica que haya datos
    If rngDatos Is Nothing Then
        LogOperacion "ERROR: La tabla TOTALES está vacía"
        Copiar = 0
        Exit Function
    End If
    
    ' Encuentra la fila destino
    filaVacia = UltimaFilaVacia(wsDestino)
    filaDestino = filaVacia + 3
    LogOperacion "Fila destino calculada: " & filaDestino
    
    ' Copia los encabezados
    tablaTOTALES.HeaderRowRange.Copy
    wsDestino.Cells(filaDestino, 2).PasteSpecial Paste:=xlPasteValues
    LogOperacion "Encabezados copiados en coordenadas: " & filaDestino & ",2"
    
    ' Agrega columnas vacías con encabezados invisibles
    'Call AgregarColumnasVacias(wsDestino, filaDestino)
    
    ' Copia TODAS las filas de datos EXCEPTO la fila "TOTAL" original
    Dim i As Long
    Dim concepto As String
    Dim filaActual As Long
    Dim filasCopiadas As Long
    
    filaActual = filaDestino + 1
    filasCopiadas = 0
    
    For i = 1 To rngDatos.Rows.Count
        concepto = Trim(rngDatos.Cells(i, 1).Value)
        LogOperacion "Procesando fila " & i & ": " & concepto
        ' Copia TODAS las filas incluyendo TOTAL
        rngDatos.Rows(i).Copy
        wsDestino.Cells(filaActual, 2).PasteSpecial Paste:=xlPasteValues
        LogOperacion "Fila copiada: " & concepto & " en coordenadas: " & filaActual & ",2"
        filaActual = filaActual + 1
        filasCopiadas = filasCopiadas + 1
    Next i
    
    LogOperacion "Total de filas copiadas: " & filasCopiadas
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
    
    ' Restaura el estado de visibilidad de la hoja origen si estaba oculta
    If hojaOculta Then
        wsOrigen.Visible = xlSheetHidden
        LogOperacion "Hoja TOTALES vuelta a ocultar"
    End If
    
    LogOperacion "=== FIN Función Copiar ==="
    Copiar = filaDestino
    
    Exit Function
    
ErrorHandler:
    Application.CutCopyMode = False
    LogOperacion "ERROR en función Copiar: " & Err.Description
    ' Restaura el estado de visibilidad en caso de error
    If hojaOculta Then
        wsOrigen.Visible = xlSheetHidden
    End If
    Copiar = 0
End Function


Function FormatearTabla(wsDestino As Worksheet, filaInicio As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rngHeaders As Range
    Dim numColumnas As Long
    Dim i As Long
    Dim j As Long
    Dim ultimaFila As Long
    Dim concepto As String
    
    ' Determina el número de columnas (ajustado para empezar en columna B)
    numColumnas = 20
    
    ' Encuentra la última fila con datos
    ultimaFila = wsDestino.Cells(wsDestino.Rows.Count, 2).End(xlUp).Row
    LogOperacion "Última fila con datos para formateo: " & ultimaFila
    
    ' Define el rango de encabezados (empezando en columna B)
    Set rngHeaders = wsDestino.Range(wsDestino.Cells(filaInicio, 2), wsDestino.Cells(filaInicio, 2 + numColumnas - 1))
    LogOperacion "Rango de encabezados: " & filaInicio & ",2 hasta " & filaInicio & "," & (2 + numColumnas - 1)
    
    ' --- Formato de Encabezados ---
    With rngHeaders
        .Interior.Color = RGB(48, 84, 150) ' Azul #305496
        .Font.Color = RGB(255, 255, 255)   ' Blanco (por defecto)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200) ' Gris claro
        .Borders.Weight = xlThin
    End With
    LogOperacion "Formato de encabezados aplicado"
    
    ' --- Encabezados invisibles (columnas 2-4) ---
    ' Texto del mismo color que el fondo para que parezcan invisibles
    For i = 2 To 4
        wsDestino.Cells(filaInicio, i).Font.Color = RGB(48, 84, 150) ' Mismo azul que el fondo #305496
    Next i
    
    ' --- Formato de TODAS las filas de datos ---
    For j = filaInicio + 1 To ultimaFila
        Dim rngFila As Range
        Set rngFila = wsDestino.Range(wsDestino.Cells(j, 2), wsDestino.Cells(j, 2 + numColumnas - 1))
        
        ' Obtiene el concepto de la primera columna (ahora columna B)
        concepto = Trim(wsDestino.Cells(j, 2).Value)
        
        ' Determina si es una fila de totales (TOTAL AIU, TOTAL, etc.)
        If UCase(concepto) Like "*TOTAL*" Then
            ' --- Formato de Fila de Totales ---
            With rngFila
                .Interior.Color = RGB(48, 84, 150) ' Azul #305496
                .Font.Color = RGB(255, 255, 255)   ' Blanco
                .Font.Bold = True
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200) ' Gris claro
                .Borders.Weight = xlThin
            End With
            
            ' Alineación específica para la primera columna de filas de totales
            wsDestino.Cells(j, 2).HorizontalAlignment = xlLeft
        Else
            ' --- Formato de Fila de Datos Normal ---
            With rngFila
                .Interior.Color = RGB(255, 255, 255) ' Blanco
                .Font.Color = RGB(0, 0, 0)          ' Negro
                .Font.Bold = False
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200) ' Gris claro
                .Borders.Weight = xlThin
            End With
        End If
        
        ' --- Formato de Números para TODAS las filas ---
        ' Columnas de porcentaje: 2, 6, 9, 12, 15, 18 (nueva estructura)
        Dim percentCols As Variant
        percentCols = Array(4, 8, 11, 14, 17, 20)
        
        For i = LBound(percentCols) To UBound(percentCols)
            wsDestino.Cells(j, percentCols(i)).NumberFormat = "0.00%"
        Next i
        LogOperacion "Formatos de porcentaje aplicados en fila " & j
        
        ' Columnas de moneda: 3, 4, 7, 10, 13, 16, 19, 20 (nueva estructura)
        Dim currencyCols As Variant
        currencyCols = Array(6,9,12,15,18,21)
        
        For i = LBound(currencyCols) To UBound(currencyCols)
            wsDestino.Cells(j, currencyCols(i)).NumberFormat = "$ #,##0.00"
        Next i
        LogOperacion "Formatos de moneda aplicados en fila " & j
    Next j
    
    ' Autoajuste de columnas removido según solicitud
    
    FormatearTabla = True
    Exit Function
    
ErrorHandler:
    FormatearTabla = False
End Function


' ============================================
' FUNCIÓN: AgregarFilaTotal
' Agrega una fila TOTAL al final con sumas de las columnas VR
' ============================================

Function AgregarFilaTotal(wsDestino As Worksheet, filaInicio As Long) As Boolean
    On Error GoTo ErrorHandler
    
    LogOperacion "=== INICIO AgregarFilaTotal ==="
    
    Dim ultimaFila As Long
    Dim filaTotal As Long
    Dim numColumnas As Long
    Dim i As Long
    Dim j As Long
    Dim suma As Double
    Dim concepto As String
    
    ' Determina el número de columnas (ajustado para empezar en columna B)
    numColumnas = 20
    
    ' Encuentra la última fila con datos (después de copiar)
    ultimaFila = wsDestino.Cells(wsDestino.Rows.Count, 2).End(xlUp).Row
    LogOperacion "Última fila con datos encontrada: " & ultimaFila
    
    ' La fila TOTAL será la siguiente
    filaTotal = ultimaFila + 1
    LogOperacion "Fila TOTAL calculada: " & filaTotal
    
    ' Coloca "TOTAL" en la columna B
    wsDestino.Cells(filaTotal, 2).Value = "TOTAL"
    LogOperacion "TOTAL colocado en coordenadas: " & filaTotal & ",2"
    
    ' Columnas VR que necesitan suma (ajustadas para nueva estructura)
    ' Las columnas VR ahora son: 6, 9, 12, 15, 18, 21
    Dim columnasVR As Variant
    columnasVR = Array(6,9,12,15,18,21)
    
    ' Calcula las sumas para cada columna VR
    For i = LBound(columnasVR) To UBound(columnasVR)
       
        suma = 0
        LogOperacion "Calculando suma para columna " & columnasVR(i)
        For j = filaInicio + 1 To ultimaFila
            concepto = Trim(wsDestino.Cells(j, 2).Value)
            LogOperacion "Revisando fila " & j & ": " & concepto
            ' Suma explícitamente solo estas 4 filas específicas
            If concepto = "COSTOS DIRECTO" Or _
               concepto = "ADMINISTRACION" Or _
               concepto = "IMPREVISTOS" Or _
               concepto = "UTILIDAD" Then
                 If IsNumeric(wsDestino.Cells(j, columnasVR(i)).Value) Then
                     suma = suma + wsDestino.Cells(j, columnasVR(i)).Value
                     LogOperacion "Sumando " & wsDestino.Cells(j, columnasVR(i)).Value & " de " & concepto
                 End If
             End If
         Next j
         wsDestino.Cells(filaTotal, columnasVR(i)).Value = suma
         LogOperacion "Suma " & suma & " colocada en coordenadas: " & filaTotal & "," & columnasVR(i)
     Next i
    
    ' Aplica formato de totales a la nueva fila
    Dim rngTotal As Range
    Set rngTotal = wsDestino.Range(wsDestino.Cells(filaTotal, 2), wsDestino.Cells(filaTotal, 2 + numColumnas - 1))
    LogOperacion "Aplicando formato a rango: " & filaTotal & ",2 hasta " & filaTotal & "," & numColumnas
    
    With rngTotal
        .Interior.Color = RGB(48, 84, 150) ' Azul #305496
        .Font.Color = RGB(255, 255, 255)   ' Blanco
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200) ' Gris claro
        .Borders.Weight = xlThin
    End With
    
    ' Alineación específica para la primera columna
    wsDestino.Cells(filaTotal, 2).HorizontalAlignment = xlLeft
    LogOperacion "Alineación izquierda aplicada en: " & filaTotal & ",2"
    
    ' Formato de números para las columnas VR
    For i = LBound(columnasVR) To UBound(columnasVR)
        wsDestino.Cells(filaTotal, columnasVR(i)).NumberFormat = "$ #,##0.00"
        LogOperacion "Formato moneda aplicado en: " & filaTotal & "," & columnasVR(i)
    Next i
    
    LogOperacion "=== FIN AgregarFilaTotal ==="
    AgregarFilaTotal = True
    Exit Function
    
ErrorHandler:
    LogOperacion "ERROR en AgregarFilaTotal: " & Err.Description
    AgregarFilaTotal = False
End Function


' ============================================
' FUNCIÓN: CrearTablaResumen
' Crea una tabla resumen de actas en la columna R
' ============================================

Function CrearTablaResumen(wsDestino As Worksheet, filaInicio As Long) As Boolean
    On Error GoTo ErrorHandler
    
    LogOperacion "=== INICIO CrearTablaResumen ==="
    
    Dim ultimaFila As Long
    Dim filaResumen As Long
    Dim columnaInicio As Long
    Dim i As Long
    
    ' Encuentra la última fila con datos de la tabla principal
    ultimaFila = wsDestino.Cells(wsDestino.Rows.Count, 2).End(xlUp).Row
    LogOperacion "Última fila de tabla principal: " & ultimaFila
    
    ' La tabla resumen va dos filas debajo de la tabla principal (con fila de separación)
    filaResumen = ultimaFila + 2
    LogOperacion "Fila de inicio para tabla resumen: " & filaResumen
    
    ' Columna R (columna 18)
    columnaInicio = 18
    LogOperacion "Columna de inicio: " & columnaInicio
    
    ' --- ENCABEZADOS ---
    wsDestino.Cells(filaResumen, columnaInicio).Value = "PARCIAL"
    wsDestino.Cells(filaResumen, columnaInicio + 1).Value = "%"
    wsDestino.Cells(filaResumen, columnaInicio + 2).Value = "FECHA DE ACTA"
    wsDestino.Cells(filaResumen, columnaInicio + 3).Value = "VALOR ACTA"
    LogOperacion "Encabezados colocados en fila: " & filaResumen
    
    ' --- DATOS ---
    ' CTA No.1 - Fecha G3, Valor I de fila TOTAL
    wsDestino.Cells(filaResumen + 1, columnaInicio).Value = "CTA No.1"
    wsDestino.Cells(filaResumen + 1, columnaInicio + 1).Value = "" ' Vacío
    wsDestino.Cells(filaResumen + 1, columnaInicio + 2).Formula = "=+G3" ' Referencia a G3 (fecha)
    wsDestino.Cells(filaResumen + 1, columnaInicio + 3).Formula = "=I" & ultimaFila ' Traer total de la fila TOTAL (columna I)
    
    ' CTA No.2 - Fecha J3, Valor L de fila TOTAL
    wsDestino.Cells(filaResumen + 2, columnaInicio).Value = "CTA No.2"
    wsDestino.Cells(filaResumen + 2, columnaInicio + 1).Value = "" ' Vacío
    wsDestino.Cells(filaResumen + 2, columnaInicio + 2).Formula = "=+J3" ' Referencia a J3 (fecha)
    wsDestino.Cells(filaResumen + 2, columnaInicio + 3).Formula = "=L" & ultimaFila ' Traer total de la fila TOTAL (columna L)
    
    ' CTA No.3 - Fecha M3, Valor O de fila TOTAL
    wsDestino.Cells(filaResumen + 3, columnaInicio).Value = "CTA No.3"
    wsDestino.Cells(filaResumen + 3, columnaInicio + 1).Value = "" ' Vacío
    wsDestino.Cells(filaResumen + 3, columnaInicio + 2).Formula = "=+M3" ' Referencia a M3 (fecha)
    wsDestino.Cells(filaResumen + 3, columnaInicio + 3).Formula = "=O" & ultimaFila ' Traer total de la fila TOTAL (columna O)
    
    ' CTA No.4 - Fecha P3, Valor R de fila TOTAL
    wsDestino.Cells(filaResumen + 4, columnaInicio).Value = "CTA No.4"
    wsDestino.Cells(filaResumen + 4, columnaInicio + 1).Value = "" ' Vacío
    wsDestino.Cells(filaResumen + 4, columnaInicio + 2).Formula = "=+P3" ' Referencia a P3 (fecha)
    wsDestino.Cells(filaResumen + 4, columnaInicio + 3).Formula = "=R" & ultimaFila ' Traer total de la fila TOTAL (columna R)
    
    ' ACTA FINAL - Fecha S3, Valor U de fila TOTAL
    wsDestino.Cells(filaResumen + 5, columnaInicio).Value = "ACTA FINAL"
    wsDestino.Cells(filaResumen + 5, columnaInicio + 1).Value = "" ' Vacío
    wsDestino.Cells(filaResumen + 5, columnaInicio + 2).Formula = "=+S3" ' Referencia a S3 (fecha)
    wsDestino.Cells(filaResumen + 5, columnaInicio + 3).Formula = "=U" & ultimaFila ' Traer total de la fila TOTAL (columna U)
    
    LogOperacion "Datos de tabla resumen colocados"
    
    ' --- FILA TOTAL ---
    wsDestino.Cells(filaResumen + 6, columnaInicio).Value = "TOTAL"
    ' Columna % (columnaInicio + 1) - Suma desde CTA No.1 hasta ACTA FINAL usando referencias relativas
    wsDestino.Cells(filaResumen + 6, columnaInicio + 1).FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)" ' Suma las 5 filas anteriores
    LogOperacion "Fórmula % con referencias relativas aplicada"
    
    wsDestino.Cells(filaResumen + 6, columnaInicio + 2).Value = "" ' Vacío
    
    ' Columna VALOR ACTA (columnaInicio + 3) - Suma desde CTA No.1 hasta ACTA FINAL usando referencias relativas
    wsDestino.Cells(filaResumen + 6, columnaInicio + 3).FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)" ' Suma las 5 filas anteriores
    LogOperacion "Fórmula VALOR con referencias relativas aplicada"
    
    LogOperacion "Fila TOTAL agregada"
    
    ' --- APLICAR FORMATOS ---
    ' Encabezados
    Dim rngHeaders As Range
    Set rngHeaders = wsDestino.Range(wsDestino.Cells(filaResumen, columnaInicio), wsDestino.Cells(filaResumen, columnaInicio + 3))
    With rngHeaders
        .Interior.Color = RGB(217, 217, 217) ' Gris
        .Font.Color = RGB(0, 0, 0) ' Negro
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Borders.Weight = xlThin
    End With
    LogOperacion "Formato de encabezados aplicado"
    
    ' Datos (filas 2-6)
    For i = 1 To 5
        Dim rngFila As Range
        Set rngFila = wsDestino.Range(wsDestino.Cells(filaResumen + i, columnaInicio), wsDestino.Cells(filaResumen + i, columnaInicio + 3))
        With rngFila
            .Interior.Color = RGB(255, 255, 255) ' Blanco
            .Font.Color = RGB(0, 0, 0) ' Negro
            .Font.Bold = False
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            .Borders.Weight = xlThin
        End With
    Next i
    LogOperacion "Formato de datos aplicado"
    
    ' Fila TOTAL
    Dim rngTotal As Range
    Set rngTotal = wsDestino.Range(wsDestino.Cells(filaResumen + 6, columnaInicio), wsDestino.Cells(filaResumen + 6, columnaInicio + 3))
    With rngTotal
        .Interior.Color = RGB(48, 84, 150) ' Azul #305496
        .Font.Color = RGB(255, 255, 255) ' Blanco
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Borders.Weight = xlThin
    End With
    wsDestino.Cells(filaResumen + 6, columnaInicio).HorizontalAlignment = xlLeft ' Primera columna alineada a la izquierda
    LogOperacion "Formato de fila TOTAL aplicado"
    
    ' Formatos numéricos
    ' Columna % (columnaInicio + 1) - formato porcentaje
    wsDestino.Cells(filaResumen + 6, columnaInicio + 1).NumberFormat = "0.00%"
    
    ' Columna VALOR ACTA (columnaInicio + 3) - formato moneda
    For i = 1 To 6
        wsDestino.Cells(filaResumen + i, columnaInicio + 3).NumberFormat = "$ #,##0.00"
    Next i
    LogOperacion "Formatos numéricos aplicados"
    
    LogOperacion "=== FIN CrearTablaResumen ==="
    CrearTablaResumen = True
    Exit Function
    
ErrorHandler:
    LogOperacion "ERROR en CrearTablaResumen: " & Err.Description
    CrearTablaResumen = False
End Function



