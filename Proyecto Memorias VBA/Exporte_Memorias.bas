' === Macro: ExportarHojasConNumero ===
' Copia todas las hojas cuyo nombre contiene al menos un número a un nuevo libro y pide el nombre del archivo.
' NO se eliminan hojas del libro original, solo se copian.
Sub ExportarHojasConNumero()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, hojasExportar As Collection, nombreLibro As String
    Dim nuevoLibro As Workbook, wsNueva As Worksheet, i As Long
    Set hojasExportar = New Collection
    ' Buscar hojas con número en el nombre
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*[0-9]*" Then
            hojasExportar.Add ws
            Debug.Print "[ExportarHojasConNumero] Hoja seleccionada: " & ws.Name
        Else
            Debug.Print "[ExportarHojasConNumero] Hoja ignorada: " & ws.Name
        End If
    Next ws
    Debug.Print "[ExportarHojasConNumero] Total hojas seleccionadas: " & hojasExportar.Count
    If hojasExportar.Count = 0 Then
        MsgBox "No se encontraron hojas con números en el nombre.", vbExclamation
        Exit Sub
    End If
    nombreLibro = InputBox("Ingrese el nombre para el nuevo libro:", "Nombre del archivo")
    Debug.Print "[ExportarHojasConNumero] Nombre ingresado: '" & nombreLibro & "'"
    If nombreLibro = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    Set nuevoLibro = Workbooks.Add(xlWBATWorksheet)
    Debug.Print "[ExportarHojasConNumero] Nuevo libro creado."
    ' Eliminar la hoja creada por defecto en el nuevo libro
    Application.DisplayAlerts = False
    For Each wsNueva In nuevoLibro.Worksheets
        Debug.Print "[ExportarHojasConNumero] Eliminando hoja por defecto: " & wsNueva.Name
        wsNueva.Delete
    Next wsNueva
    Application.DisplayAlerts = True
    ' Copiar hojas seleccionadas al nuevo libro (NO se eliminan del original)
    For i = 1 To hojasExportar.Count
        Debug.Print "[ExportarHojasConNumero] Copiando hoja: " & hojasExportar(i).Name
        hojasExportar(i).Copy After:=nuevoLibro.Sheets(nuevoLibro.Sheets.Count)
    Next i
    ' Guardar el libro
    Application.DisplayAlerts = False
    Debug.Print "[ExportarHojasConNumero] Guardando libro como: " & ThisWorkbook.Path & "\" & nombreLibro & ".xlsx"
    nuevoLibro.SaveAs ThisWorkbook.Path & "\" & nombreLibro & ".xlsx"
    Application.DisplayAlerts = True
    MsgBox "Libro exportado correctamente como '" & nombreLibro & ".xlsx' en: " & ThisWorkbook.Path, vbInformation
    Debug.Print "[ExportarHojasConNumero] Proceso finalizado."
    Exit Sub
ErrHandler:
    MsgBox "Error en ExportarHojasConNumero: " & Err.Description, vbCritical
    Debug.Print "[ExportarHojasConNumero] Error: " & Err.Description
End Sub
