' Attribute VB_Name = "mod_ValidacionCamposFormulario"
Public Function ValidarCamposFormulario(frm As Object) As Boolean
    ValidarCamposFormulario = False ' Asumimos inv�lido por defecto

    ' Validar seg�n la opci�n seleccionada
    If frm.opt_OT.Value Then
        If Trim(frm.txt_OT.Value) = "" Or Trim(frm.cmb_Especialidad.Value) = "" Or Trim(frm.cmb_SubInd.Value) = "" Then
            MsgBox "Debes completar OT, Especialidad y Sub�ndice para continuar.", vbExclamation
            Exit Function
        End If
    ElseIf frm.opt_CC.Value Then
        If Trim(frm.cmb_CentroCosto.Value) = "" Or Trim(frm.cmb_Subcentro.Value) = "" Then
            MsgBox "Debes completar Centro de Costo y Subcentro para continuar.", vbExclamation
            Exit Function
        End If
    End If

    ' Validar horas obligatorias solo si no hay ausentismo
    If Trim(frm.Ausentismo.Value) = "" Then
        If Trim(frm.H_Entrada.Value) = "" Or Trim(frm.H_Salida.Value) = "" Or Trim(frm.T_Almuerzo.Value) = "" Then
            MsgBox "Debes completar la hora de entrada, salida y tiempo de almuerzo.", vbExclamation
            Exit Function
        End If
    End If

    ' Si lleg� aqu�, es porque pas� todas las validaciones
    ValidarCamposFormulario = True
End Function

