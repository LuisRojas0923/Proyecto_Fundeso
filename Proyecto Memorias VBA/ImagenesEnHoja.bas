Public Sub CargarImagenEnControl(controlName As String, hoja As Worksheet)
    Dim fd As FileDialog
    Dim imgPath As String
    Dim oleObj As OLEObject
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecciona una imagen"
        .Filters.Clear
        .Filters.Add "Imágenes", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"
        .AllowMultiSelect = False
        If .Show = -1 Then
            imgPath = .SelectedItems(1)
            Set oleObj = hoja.OLEObjects(controlName)
            oleObj.Object.Picture = LoadPicture(imgPath)
        End If
    End With
    Set fd = Nothing
End Sub

' === EJEMPLO: Pega esto en el módulo de la hoja donde están los controles ===
' Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'     CargarImagenEnControl "Image1", Me
' End Sub
'
' Private Sub Image2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'     CargarImagenEnControl "Image2", Me
' End Sub
