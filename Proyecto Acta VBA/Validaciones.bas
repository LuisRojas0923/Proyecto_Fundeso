Attribute VB_Name = "Validaciones"

' === MÓDULO DE VALIDACIONES PARA FORMULARIO DE CREACIÓN DE MEMORIAS ===

' Función principal para validar todos los controles del formulario
Public Sub ValidarControlesFormulario(frm As Object)
    On Error GoTo ErrHandler
    
    Debug.Print "=== INICIANDO VALIDACIÓN DE CONTROLES ==="
    Debug.Print "Formulario: " & frm.Name
    
    Dim controlesFaltantes As String
    controlesFaltantes = ""
    
    ' === VALIDAR MULTIPAGE ===
    If Not ExisteControl(frm, "MultiPage1") Then
        controlesFaltantes = controlesFaltantes & "MultiPage1, "
        Debug.Print "❌ FALTANTE: MultiPage1"
    Else
        Debug.Print "✅ MultiPage1 encontrado"
        ' Validar páginas del MultiPage
        ValidarPaginasMultiPage frm
    End If
    
    ' === VALIDAR CONTROLES GLOBALES ===
    ValidarControlesGlobales frm, controlesFaltantes
    
    ' === VALIDAR CONTROLES DE PÁGINA 1 (Selección) ===
    ValidarControlesPagina1 frm, controlesFaltantes
    
    ' === VALIDAR CONTROLES DE PÁGINA 2 (Validación y Exportación) ===
    ValidarControlesPagina2 frm, controlesFaltantes
    
    ' === VALIDAR CONTROLES DE PÁGINA 3 (Revisión) ===
    ValidarControlesPagina3 frm, controlesFaltantes
    
    ' === RESUMEN FINAL ===
    Debug.Print "=== RESUMEN DE VALIDACIÓN ==="
    If controlesFaltantes = "" Then
        Debug.Print "✅ TODOS LOS CONTROLES ESTÁN CREADOS CORRECTAMENTE"
    Else
        Debug.Print "❌ CONTROLES FALTANTES: " & Left(controlesFaltantes, Len(controlesFaltantes) - 2)
    End If
    Debug.Print "================================"
    
    Exit Sub
ErrHandler:
    Debug.Print "ERROR en ValidarControlesFormulario: " & Err.Description
End Sub

' Función auxiliar para verificar si existe un control
Private Function ExisteControl(frm As Object, nombreControl As String) As Boolean
    On Error GoTo ErrHandler
    Dim ctrl As Object
    Set ctrl = frm.Controls(nombreControl)
    ExisteControl = True
    Exit Function
ErrHandler:
    ExisteControl = False
End Function

' Validar páginas del MultiPage
Private Sub ValidarPaginasMultiPage(frm As Object)
    On Error GoTo ErrHandler
    
    Dim mp As Object
    Set mp = frm.Controls("MultiPage1")
    
    Debug.Print "  📄 Páginas del MultiPage:"
    
    If mp.Pages.Count >= 1 Then
        Debug.Print "    ✅ Página 1: " & mp.Pages(0).Caption
    Else
        Debug.Print "    ❌ FALTANTE: Página 1 del MultiPage"
    End If
    
    If mp.Pages.Count >= 2 Then
        Debug.Print "    ✅ Página 2: " & mp.Pages(1).Caption
    Else
        Debug.Print "    ❌ FALTANTE: Página 2 del MultiPage"
    End If
    
    If mp.Pages.Count >= 3 Then
        Debug.Print "    ✅ Página 3: " & mp.Pages(2).Caption
    Else
        Debug.Print "    ❌ FALTANTE: Página 3 del MultiPage"
    End If
    
    Exit Sub
ErrHandler:
    Debug.Print "    ❌ ERROR validando páginas del MultiPage: " & Err.Description
End Sub

' Validar controles globales (fuera del MultiPage)
Private Sub ValidarControlesGlobales(frm As Object, ByRef controlesFaltantes As String)
    On Error GoTo ErrHandler
    
    Debug.Print "  🌐 Controles Globales:"
    
    Dim controlesGlobales() As String
    controlesGlobales = Split("btn_LimpiarCampos,btn_Marcar,btn_Desmarcar", ",")
    
    Dim i As Long
    For i = LBound(controlesGlobales) To UBound(controlesGlobales)
        If ExisteControl(frm, Trim(controlesGlobales(i))) Then
            Debug.Print "    ✅ " & Trim(controlesGlobales(i))
        Else
            Debug.Print "    ❌ FALTANTE: " & Trim(controlesGlobales(i))
            controlesFaltantes = controlesFaltantes & Trim(controlesGlobales(i)) & ", "
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Debug.Print "    ❌ ERROR validando controles globales: " & Err.Description
End Sub

' Validar controles de Página 1 (Selección)
Private Sub ValidarControlesPagina1(frm As Object, ByRef controlesFaltantes As String)
    On Error GoTo ErrHandler
    
    Debug.Print "  📄 Controles Página 1 (Selección):"
    
    Dim controlesPagina1() As String
    controlesPagina1 = Split("Palabra_Clave,cmb_Area,cmb_Capitulos,Listbox_Registros,btn_AgregarATrabajo", ",")
    
    Dim i As Long
    For i = LBound(controlesPagina1) To UBound(controlesPagina1)
        If ExisteControl(frm, Trim(controlesPagina1(i))) Then
            Debug.Print "    ✅ " & Trim(controlesPagina1(i))
        Else
            Debug.Print "    ❌ FALTANTE: " & Trim(controlesPagina1(i))
            controlesFaltantes = controlesFaltantes & Trim(controlesPagina1(i)) & ", "
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Debug.Print "    ❌ ERROR validando controles Página 1: " & Err.Description
End Sub

' Validar controles de Página 2 (Validación y Exportación)
Private Sub ValidarControlesPagina2(frm As Object, ByRef controlesFaltantes As String)
    On Error GoTo ErrHandler
    
    Debug.Print "  📄 Controles Página 2 (Validación y Exportación):"
    
    Dim controlesPagina2() As String
    controlesPagina2 = Split("Listbox_Trabajo,txt_Cantidad,btn_Exportar,btn_EliminarSeleccionado,btn_AsignarCantidad", ",")
    
    Dim i As Long
    For i = LBound(controlesPagina2) To UBound(controlesPagina2)
        If ExisteControl(frm, Trim(controlesPagina2(i))) Then
            Debug.Print "    ✅ " & Trim(controlesPagina2(i))
        Else
            Debug.Print "    ❌ FALTANTE: " & Trim(controlesPagina2(i))
            controlesFaltantes = controlesFaltantes & Trim(controlesPagina2(i)) & ", "
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Debug.Print "    ❌ ERROR validando controles Página 2: " & Err.Description
End Sub

' Validar controles de Página 3 (Revisión)
Private Sub ValidarControlesPagina3(frm As Object, ByRef controlesFaltantes As String)
    On Error GoTo ErrHandler
    
    Debug.Print "  📄 Controles Página 3 (Revisión):"
    
    Dim controlesPagina3() As String
    controlesPagina3 = Split("Listbox_Exportados,btn_ActualizarVista", ",")
    
    Dim i As Long
    For i = LBound(controlesPagina3) To UBound(controlesPagina3)
        If ExisteControl(frm, Trim(controlesPagina3(i))) Then
            Debug.Print "    ✅ " & Trim(controlesPagina3(i))
        Else
            Debug.Print "    ❌ FALTANTE: " & Trim(controlesPagina3(i))
            controlesFaltantes = controlesFaltantes & Trim(controlesPagina3(i)) & ", "
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Debug.Print "    ❌ ERROR validando controles Página 3: " & Err.Description
End Sub
