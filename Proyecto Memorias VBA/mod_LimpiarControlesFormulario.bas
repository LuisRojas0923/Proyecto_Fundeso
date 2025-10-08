' Attribute VB_Name = "mod_LimpiarControlesFormulario"
Public Sub LimpiarControlesFormulario(frm As Object)
    ' Limpiar campos de entrada del formulario
    With frm
        .txt_OT.Value = ""
        .cmb_Especialidad.Value = ""
        .cmb_SubInd.Value = ""
        .cmb_CentroCosto.Value = ""
        .cmb_Subcentro.Value = ""
        .cmb_Area.Value = ""
        .cmb_Regional.Value = ""
        .cmb_Regional.Enabled = False
        
        .Ausentismo.Value = ""
        .Observaciones.Value = ""

        .F_Desde.Value = ""
        .F_Hasta.Value = ""
        .H_Entrada.Value = ""
        .H_Salida.Value = ""
        .T_Almuerzo.Value = ""

        ' Restablecer opci�n No Aplica como predeterminada
        .opt_OT.Value = False
        .opt_CC.Value = False
        .opt_NoAplica.Value = True

      
    End With

    MsgBox "Formulario limpiado correctamente.", vbInformation
End Sub

