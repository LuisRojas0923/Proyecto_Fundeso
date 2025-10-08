Attribute VB_Name = "Modulo_Sincronizacion"

' Macro para actualizar el proyecto VBA desde los archivos fuente (.bas, .frm)
' ADVERTENCIA: Requiere activar "Confiar en el acceso al modelo de objetos de proyectos de VBA"
' en Excel: Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza > Configuración de macros.

Public Sub ActualizarProyectoDesdeArchivos()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object ' VBProject
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object, component As Object
    Dim componentName As String, fileExtension As String
    Dim i As Integer
    
    Set vbProj = ThisWorkbook.VBProject
    folderPath = "C:\Users\luisr\Desktop\Proyecto Acta VBA"
    
    Debug.Print "--- INICIANDO SINCRONIZACION DE PROYECTO ---"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "No se encontro la carpeta del proyecto: " & folderPath, vbCritical, "Error"
        Exit Sub
    End If
    Set folder = fso.GetFolder(folderPath)
    
    Application.ScreenUpdating = False
    Application.VBE.MainWindow.Visible = True
    
    ' --- PASO 1: Eliminar todos los MODULOS estándar existentes (excepto este) ---
    Debug.Print "--- FASE 1: ELIMINANDO MODULOS ANTIGUOS ---"
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set component = vbProj.VBComponents(i)
        ' Type 1 = vbext_ct_StdModule (módulo de código estándar)
        If component.Type = 1 Then
            If component.Name <> "Modulo_Sincronizacion" Then
                Debug.Print "  - Eliminando modulo: " & component.Name
                vbProj.VBComponents.Remove component
                DoEvents
            End If
        End If
    Next i
    Debug.Print "--- ELIMINACION DE MODULOS COMPLETADA ---"
    
    ' --- PASO 2: Importar módulos nuevos y actualizar formularios existentes ---
    Debug.Print "--- FASE 2: IMPORTANDO/ACTUALIZANDO COMPONENTES ---"
    For Each file In folder.Files
        fileExtension = LCase(fso.GetExtensionName(file.Name))
        componentName = fso.GetBaseName(file.Name)
        
        If componentName = "Modulo_Sincronizacion" Then
            ' Saltar este módulo para no re-importarlo
        ElseIf fileExtension = "bas" Then
            ' Los módulos ya fueron eliminados, así que solo se importan.
            Debug.Print "  - Importando modulo: " & file.Name
            vbProj.VBComponents.Import file.Path
            DoEvents
        ElseIf fileExtension = "frm" Then
            ' Los formularios no fueron eliminados, por lo que deben ser reemplazados.
            Debug.Print "  - Actualizando formulario: " & file.Name
            ' Eliminar el antiguo
            On Error Resume Next
            vbProj.VBComponents.Remove vbProj.VBComponents(componentName)
            On Error GoTo ErrHandler
            ' Importar el nuevo
            vbProj.VBComponents.Import file.Path
            DoEvents
        End If
    Next file
    Debug.Print "--- IMPORTACION COMPLETADA ---"
    
    Application.ScreenUpdating = True
    
    Debug.Print "--- SINCRONIZACION COMPLETADA ---"
    MsgBox "El proyecto de VBA ha sido actualizado con los archivos de la carpeta.", vbInformation, "Sincronizacion Exitosa"
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Debug.Print "Error en ActualizarProyectoDesdeArchivos: " & Err.Description
    MsgBox "Ocurrio un error durante la actualizacion: " & vbCrLf & Err.Description, vbCritical, "Error"
    Application.StatusBar = originalStatusBar
End Sub
