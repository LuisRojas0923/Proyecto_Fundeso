' =============================
' MODIFICACIONES RECIENTES
' =============================
' - Mejora de la búsqueda de registros en TexCtaAux_Change: ahora usa dos llaves (Item_3 y Descripción) para identificar la fila exacta.
' - Se agregaron mensajes Debug.Print para depuración, mostrando las llaves usadas y el resultado de la búsqueda.
' - Se robusteció la comparación ignorando mayúsculas/minúsculas y espacios.
' =============================


Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdModificar_Click()
'ComClasificación  txtPorcentaje

Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("CuentasMenores")
Dim tabla As ListObject: Set tabla = ws.ListObjects("CuentasAuxiliares")

If Trim(txtCantidad.Value) = "" Or Trim(txtUnidad.Value) = "" Or Trim(txtVR_Unitario.Value) = "" _
Or Trim(ComClasificación.Value) = "" Or Trim(txtPorcentaje.Value) = "" Then

            MsgBox "Debe completar Cantidad, Unidad, Clasificación,  VR_Unitario y Porcentaje para modificar.", vbExclamation
            Exit Sub
'            GoTo Fin
        End If
        
        modificar = True
        
       
        
        For Each fila In tabla.DataBodyRange.Rows
            If Trim(fila.Cells(1, tabla.ListColumns("Item_3").Index).Value) = Me.TexAuxPadre.Value And _
            Trim(fila.Cells(1, tabla.ListColumns("Descripción").Index).Value) = Me.TexCtaAux.Value Then
            
                fila.Cells(1, tabla.ListColumns("Unidad").Index).Value = Trim(txtUnidad.Value)
                fila.Cells(1, tabla.ListColumns("Cantidad").Index).Value = Trim(txtCantidad.Value)
                fila.Cells(1, tabla.ListColumns("Desperdicio").Index).Value = Trim(TexDesperdicio.Value)
                fila.Cells(1, tabla.ListColumns("Vr/Unitario").Index).Value = Trim(txtVR_Unitario.Value)
                fila.Cells(1, tabla.ListColumns("Vr/Parcial").Index).Value = Trim(TexVrParcial.Value)
                fila.Cells(1, tabla.ListColumns("Clasificación").Index).Value = Trim(ComClasificación.Value)
                fila.Cells(1, tabla.ListColumns("Porcentaje").Index).Value = IIf(Trim(Me.txtPorcentaje.Value) > 1, Trim(Me.txtPorcentaje.Value) / 100, Trim(Me.txtPorcentaje.Value))
                fila.Cells(1, tabla.ListColumns("Valor Contratistas").Index).Value = Trim(txtValor_Contratista.Value)
                
                reemplazos = reemplazos + 1
            End If
        Next fila
   
    
    If Not modificar Then
        MsgBox "Debe marcar al menos un campo a modificar.", vbInformation
        GoTo Fin
    End If

'Actualizar query

Call ActualizarTb

'------------Carga el Lisbox de Items-------------
Dim lbxPadre  As MSForms.ListBox
Dim lbxAux    As MSForms.ListBox

Set lbxPadre = Worksheets("Consultar").OLEObjects("AuxCuentaPadre").Object
Set lbxAux = Worksheets("Consultar").OLEObjects("CuentAuxiliares").Object

Call CargarListaAuxiliar(lbxPadre, lbxAux)

    MsgBox "Modificación completada. Se actualizaron " & reemplazos & " fila(s).", vbInformation

Fin:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

   Unload Me

End Sub


Private Sub TexAuxPadre_Change()

End Sub

Private Sub TexCtaAux_Change()
    ' --- Carga los datos de la fila seleccionada en los controles del formulario ---
    ' Busca la fila en la tabla "CuentasAuxiliares" usando dos llaves:
    '   - Item_3 (TexAuxPadre)
    '   - Descripción (TexCtaAux)
    ' Si encuentra la fila, carga los datos en los controles; si no, limpia los controles.
    
    Dim ws                As Worksheet
    Dim tbl               As ListObject
    Dim descClave         As String ' Llave: Descripción
    Dim padreClave        As String ' Llave: Item_3
    Dim cel               As Range  ' Celda encontrada
    Dim idxDesc           As Long, idxUni As Long, idxCant As Long
    Dim idxDesperdicio    As Long, idxVrUni As Long, idxVrParcial As Long
    Dim idxClasif         As Long, idxPorcentaje As Long, idxValorContra As Long
    Dim idxPadre          As Long
    
    '--- 0. Si no hay selección, limpia y sal ---
    LimpiarControles
    
    '--- 1. Referencias de hoja y tabla ---
    Set ws = ThisWorkbook.Worksheets("CuentasMenores")
    Set tbl = ws.ListObjects("CuentasAuxiliares")
    
    '--- 2. Valores de las llaves (se normalizan para evitar errores de formato) ---
    descClave = Trim(UCase(Me.TexCtaAux))
    padreClave = Trim(UCase(Me.TexAuxPadre))
    Debug.Print "[TexCtaAux_Change] Buscando con llaves: Item_3='" & padreClave & "', Descripción='" & descClave & "'"
    If descClave = "" Or padreClave = "" Then
        LimpiarControles
        Debug.Print "[TexCtaAux_Change] Alguna llave vacía, saliendo."
        Exit Sub
    End If
    
    '--- 3. Índices de columnas de la tabla ---
    With tbl
        idxPadre = .ListColumns("Item_3").Index
        idxDesc = .ListColumns("Descripción").Index
        idxUni = .ListColumns("Unidad").Index
        idxCant = .ListColumns("Cantidad").Index
        idxDesperdicio = .ListColumns("Desperdicio").Index
        idxVrUni = .ListColumns("Vr/Unitario").Index
        idxVrParcial = .ListColumns("Vr/Parcial").Index
        idxClasif = .ListColumns("Clasificación").Index
        idxPorcentaje = .ListColumns("Porcentaje").Index
        idxValorContra = .ListColumns("Valor Contratistas").Index
    End With
    
    '--- 4. Buscar la fila por ambas llaves (Item_3 y Descripción) ---
    Dim celRng As Range
    Set celRng = Nothing
    Dim r As Range
    For Each r In tbl.DataBodyRange.Rows
        ' Compara ambas llaves normalizadas
        If Trim(UCase(r.Cells(1, idxPadre).Value)) = padreClave And _
           Trim(UCase(r.Cells(1, idxDesc).Value)) = descClave Then
            Set celRng = r.Cells(1, idxDesc)
            Exit For
        End If
    Next r
    Set cel = celRng
    
    If cel Is Nothing Then
        LimpiarControles
        Debug.Print "[TexCtaAux_Change] No se encontró con llaves: Item_3='" & padreClave & "', Descripción='" & descClave & "'"
        MsgBox "No se encontró el registro con las llaves: '" & Me.TexAuxPadre & "' y '" & Me.TexCtaAux & "' en la tabla.", vbExclamation
        Exit Sub
    Else
        Debug.Print "[TexCtaAux_Change] Encontrado: '" & cel.Value & "' (fila: " & cel.Row & ")"
    End If
    
    '--- 5. Volcar datos de la fila encontrada en los controles del formulario ---
    With Me
        .txtUnidad.Value = cel.Offset(0, idxUni - idxDesc).Value
        .txtCantidad.Value = cel.Offset(0, idxCant - idxDesc).Value
        .TexDesperdicio.Value = cel.Offset(0, idxDesperdicio - idxDesc).Value
        .txtVR_Unitario.Value = cel.Offset(0, idxVrUni - idxDesc).Value
        .TexVrParcial.Value = cel.Offset(0, idxVrParcial - idxDesc).Value
        .ComClasificación.Value = cel.Offset(0, idxClasif - idxDesc).Value
        .txtPorcentaje.Value = cel.Offset(0, idxPorcentaje - idxDesc).Value
        .txtValor_Contratista.Value = cel.Offset(0, idxValorContra - idxDesc).Value
    End With
    
End Sub
Private Sub LimpiarControles()
    With Me
        .txtUnidad.Value = vbNullString
        .txtCantidad.Value = vbNullString
        .TexDesperdicio.Value = vbNullString
        .txtVR_Unitario.Value = vbNullString
        .TexVrParcial.Value = vbNullString
        .ComClasificación.Value = vbNullString
        .txtPorcentaje.Value = vbNullString
        .txtValor_Contratista.Value = vbNullString
    End With
End Sub


Private Sub TexVrParcial_Change()
On Error GoTo Error

Me.txtValor_Contratista = (Me.txtPorcentaje.Value * Me.TexVrParcial.Value) / 100

Exit Sub
Error:
Me.txtValor_Contratista = 0

End Sub

Private Sub txtCantidad_Change()
Me.TexDesperdicio = 0

On Error GoTo Error

Me.TexVrParcial = Me.txtCantidad.Value * Me.txtVR_Unitario.Value + ((Me.txtCantidad.Value * Me.txtVR_Unitario.Value) * (Me.TexDesperdicio.Value / 100))
Exit Sub
Error:
Me.TexVrParcial = 0

End Sub

Private Sub txtPorcentaje_Change()
On Error GoTo Error

Me.txtValor_Contratista = (Me.txtPorcentaje.Value * Me.TexVrParcial.Value) / 100

Exit Sub
Error:
Me.txtValor_Contratista = 0

End Sub

Private Sub txtVR_Unitario_Change()
On Error GoTo Error

Me.TexVrParcial = Me.txtCantidad.Value * Me.txtVR_Unitario.Value + ((Me.txtCantidad.Value * Me.txtVR_Unitario.Value) * (Me.TexDesperdicio.Value / 100))
Exit Sub
Error:
Me.TexVrParcial = 0


End Sub

Private Sub UserForm_Initialize()

Me.TexAuxPadre.Enabled = False
Me.TexCtaAux.Enabled = False

'limpia texto
Me.TexAuxPadre = ""
Me.TexCtaAux = ""


Dim hoja As Worksheet
Dim lisbox1, lisbox2 As MSForms.ListBox
Dim valorlista1, valorlista2 As String

Set hoja = ThisWorkbook.Worksheets("Consultar")
Set lisbox1 = hoja.OLEObjects("AuxCuentaPadre").Object
Set lisbox2 = hoja.OLEObjects("CuentAuxiliares").Object

valorlista1 = lisbox1.List(lisbox1.ListIndex, 0)
valorlista2 = lisbox2.List(lisbox2.ListIndex, 1)

Me.TexAuxPadre = valorlista1
Me.TexCtaAux = valorlista2


'cargar lista
With Me.ComClasificación
    .Clear
    .AddItem "Materiales"
    .AddItem "Equipo"
    .AddItem "Mano de Obra"
    .AddItem "Otros"
    
End With


End Sub
