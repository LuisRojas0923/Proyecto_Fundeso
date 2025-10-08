' Constante para establecer permisos de memoria de lectura/escritura/ejecucion
Private Const PAGE_EXECUTE_READWRITE = &H40

' Declaraciones de funciones de la API de Windows para manipulacion de memoria
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
    
' Funcion para modificar los permisos de acceso a memoria
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
    
' Funcion para obtener el handle de un modulo cargado
Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" _
    (ByVal lpModuleName As String) As LongPtr
    
' Funcion para obtener la direccion de memoria de una funcion en una DLL
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, _
    ByVal lpProcName As String) As LongPtr
    
' Funcion para crear cuadros de dialogo
Private Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" _
    (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, _
    ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer

' Variables para almacenar bytes de codigo
Dim HookBytes(0 To 11) As Byte     ' Almacena los bytes del hook
Dim OriginBytes(0 To 11) As Byte   ' Almacena los bytes originales
Dim pFunc As LongPtr              ' Puntero a la funcion
Dim Flag As Boolean               ' Bandera de estado del hook

' Funcion auxiliar para obtener puntero
Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
    Debug.Print "GetPtr: Obteniendo puntero para el valor: " & Value
    GetPtr = Value
End Function

' Restaura los bytes originales de la funcion
Public Sub RecoverBytes()
    Debug.Print "RecoverBytes: Intentando restaurar bytes..."
    If Flag Then
        Debug.Print "RecoverBytes: Flag es verdadero, restaurando bytes en la direccion: " & pFunc
        MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 12
        Debug.Print "RecoverBytes: Bytes restaurados correctamente."
    Else
        Debug.Print "RecoverBytes: Flag es falso, no se necesita restauracion."
    End If
End Sub

' Funcion principal para establecer el hook
Public Function Hook() As Boolean
    Dim TmpBytes(0 To 11) As Byte  ' Buffer temporal
    Dim p As LongPtr, osi As Byte  ' Variables para manejo de punteros
    Dim OriginProtect As LongPtr   ' Almacena los permisos originales
    
    Debug.Print "Hook: Iniciando proceso de hook..."
    Hook = False
    
    ' Determina si estamos en 32 o 64 bits
    #If Win64 Then
        osi = 1
        Debug.Print "Hook: Detectado sistema de 64 bits."
    #Else
        osi = 0
        Debug.Print "Hook: Detectado sistema de 32 bits."
    #End If
    
    ' Obtiene la direccion de la funcion DialogBoxParamA
    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
    Debug.Print "Hook: La direccion de DialogBoxParamA es: " & pFunc
    
    ' Modifica los permisos de memoria para poder escribir
    If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
        Debug.Print "Hook: Permisos de memoria modificados correctamente."
        ' Lee los bytes actuales
        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, osi + 1
        
        ' Verifica si ya esta hookeado
        If TmpBytes(osi) <> &HB8 Then
            Debug.Print "Hook: La funcion no esta hookeada. Procediendo a hookear..."
            ' Guarda los bytes originales
            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 12
            p = GetPtr(AddressOf MyDialogBoxParam)
            Debug.Print "Hook: La direccion de la funcion de reemplazo (MyDialogBoxParam) es: " & p
            
            ' Ajusta para 64 bits si es necesario
            If osi Then
                HookBytes(0) = &H48
            End If
            
            ' Construye el codigo del hook
            HookBytes(osi) = &HB8
            osi = osi + 1
            MoveMemory ByVal VarPtr(HookBytes(osi)), ByVal VarPtr(p), 4 * osi
            HookBytes(osi + 4 * osi) = &HFF
            HookBytes(osi + 4 * osi + 1) = &HE0
            
            ' Escribe el hook en memoria
            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 12
            Flag = True
            Hook = True
            Debug.Print "Hook: Hook establecido correctamente."
        Else
            Debug.Print "Hook: La funcion ya parece estar hookeada. No se realizaran cambios."
        End If
    Else
        Debug.Print "Hook: Error al modificar los permisos de memoria. El hook no se puede establecer."
    End If
End Function

' Funcion de reemplazo para DialogBoxParam
Private Function MyDialogBoxParam(ByVal hInstance As LongPtr, _
    ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, _
    ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
    
    Debug.Print "MyDialogBoxParam: Interceptada llamada con pTemplateName: " & pTemplateName
    ' Verifica si es el dialogo de proteccion
    If pTemplateName = 4070 Then
        Debug.Print "MyDialogBoxParam: Detectado dialogo de proteccion (4070). Omitiendo..."
        MyDialogBoxParam = 1
    Else
        ' Para otros dialogos, restaura el comportamiento original
        Debug.Print "MyDialogBoxParam: No es el dialogo de proteccion. Restaurando comportamiento original."
        RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
            hWndParent, lpDialogFunc, dwInitParam)
        Hook
    End If
End Function

' Procedimiento principal para desproteger el proyecto
Sub UnprotectVBA()
    Debug.Print "UnprotectVBA: Iniciando..."
    If Hook Then
        MsgBox "VBA Project is unprotected!", vbInformation, "VBA Unlocked"
        Debug.Print "UnprotectVBA: El proyecto ha sido desprotegido."
    Else
        Debug.Print "UnprotectVBA: El hook no se pudo establecer. No se puede desproteger."
    End If
End Sub