' Attribute VB_Name = "ListarConsultas"
Option Explicit


Sub ListarTablasConexionesYRutaPQ()
    ' Declaraci�n de variables
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim consultaNombre As String
    Dim hojaResultados As Worksheet
    Dim fila As Long
    Dim consulta As WorkbookQuery
    Dim formulaM As String
    Dim rutaOrigen As String
    Dim cargadaEnTabla As Boolean
    Dim pasoNavegacion As String
    Dim posInicio As Long, posFin As Long
    Dim rangoDatos As Range
    Dim nombreTabla As String

    ' Crear una nueva hoja para listar las tablas y las consultas
    Set hojaResultados = ThisWorkbook.Sheets.Add
    hojaResultados.Name = "TablasYConsultasPowerQuery"

    ' Encabezados para la lista con el nuevo orden
    hojaResultados.Cells(1, 1).Value = "Nombre de la Consulta"
    hojaResultados.Cells(1, 2).Value = "Ruta de Origen"
    hojaResultados.Cells(1, 3).Value = "Paso de Navegaci�n"
    hojaResultados.Cells(1, 4).Value = "Hoja"
    hojaResultados.Cells(1, 5).Value = "Cargada en Tabla"

    fila = 2 ' Empezar en la fila 2 para los resultados

    ' Recorrer todas las consultas del libro
    For Each consulta In ThisWorkbook.Queries
        formulaM = consulta.Formula
        rutaOrigen = ""
        pasoNavegacion = ""

        ' Verificar si contiene una referencia a un archivo externo (File.Contents)
        If InStr(formulaM, "File.Contents") > 0 Then
            rutaOrigen = Mid(formulaM, InStr(formulaM, "File.Contents(") + 14)
            rutaOrigen = Left(rutaOrigen, InStr(rutaOrigen, ")") - 1)
        Else
            rutaOrigen = "$Workbook$ (Interna)"
        End If

        ' Buscar el paso de navegaci�n (Item="nombre", Kind="Table" o Kind="Sheet")
        If InStr(formulaM, "Item=") > 0 Then
            posInicio = InStr(formulaM, "Item=") + 5
            posFin = InStr(formulaM, "Kind=") - 2
            pasoNavegacion = Mid(formulaM, posInicio, posFin - posInicio + 1)
            pasoNavegacion = pasoNavegacion & " (" & Mid(formulaM, posFin + 8, 5) & ")"
        End If

        ' Inicializar como no cargada en tabla
        cargadaEnTabla = False
        
        ' Verificar si la consulta est� cargada en una tabla
        For Each ws In ThisWorkbook.Sheets
            For Each lo In ws.ListObjects
                If lo.SourceType = xlSrcQuery Then
                    ' Usar coincidencia flexible en el nombre de la conexi�n
                    consultaNombre = lo.QueryTable.WorkbookConnection.Name
                    If InStr(1, consultaNombre, consulta.Name, vbTextCompare) > 0 Then
                        cargadaEnTabla = True
                        ' Preparar datos en memoria para escritura en bloque (OPTIMIZADO)
                        Dim datosTabla(1 To 1, 1 To 5) As Variant
                        datosTabla(1, 1) = consulta.Name ' Nombre de la consulta
                        datosTabla(1, 2) = rutaOrigen ' Ruta de origen
                        datosTabla(1, 3) = pasoNavegacion ' Paso de navegaci�n
                        datosTabla(1, 4) = ws.Name ' Nombre de la hoja
                        datosTabla(1, 5) = lo.Name ' Nombre de la tabla
                        ' Escribir toda la fila de una vez
                        hojaResultados.Range("A" & fila & ":E" & fila).Value = datosTabla
                        fila = fila + 1
                        Exit For
                    End If
                End If
            Next lo
        Next ws

        ' Si no est� cargada en ninguna tabla, registrarla como "Solo Conexi�n"
        If Not cargadaEnTabla Then
            ' Preparar datos en memoria para escritura en bloque (OPTIMIZADO)
            Dim datosTabla2(1 To 1, 1 To 5) As Variant
            datosTabla2(1, 1) = consulta.Name ' Nombre de la consulta
            datosTabla2(1, 2) = rutaOrigen ' Ruta de origen
            datosTabla2(1, 3) = pasoNavegacion ' Paso de navegaci�n
            datosTabla2(1, 4) = "" ' Sin hoja, porque no est� cargada en ninguna
            datosTabla2(1, 5) = "No (Solo Conexi�n)"
            ' Escribir toda la fila de una vez
            hojaResultados.Range("A" & fila & ":E" & fila).Value = datosTabla2
            fila = fila + 1
        End If
    Next consulta

    ' Seleccionar el rango de datos y aplicar formato de tabla
    Set rangoDatos = hojaResultados.Range(hojaResultados.Cells(1, 1), hojaResultados.Cells(fila - 1, 5))
    
    ' Aplicar formato de tabla
    Set lo = hojaResultados.ListObjects.Add(xlSrcRange, rangoDatos, , xlYes)
    
    ' Nombrar la tabla como "ficha de consulta"
    lo.Name = "ficha_de_consulta"

    ' Mostrar mensaje de finalizaci�n del proceso
    MsgBox "Proceso completado. Ver hoja 'TablasYConsultasPowerQuery'.", vbInformation
End Sub

