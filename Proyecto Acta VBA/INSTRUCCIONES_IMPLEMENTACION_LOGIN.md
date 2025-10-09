# Sistema de Autenticacion para Excel - Guia de Implementacion

## Resumen del Sistema

Este sistema implementa autenticacion obligatoria al abrir un libro Excel. Las caracteristicas principales son:

- ✅ Login obligatorio al abrir el archivo
- ✅ Credenciales almacenadas en hoja oculta y protegida
- ✅ El libro se cierra automaticamente si falla la autenticacion
- ✅ Proteccion completa de todas las hojas hasta el login exitoso
- ✅ Acceso a configuracion de usuarios solo para administrador
- ✅ Sistema de logs integrado
- ✅ Manejo de dos contextos: apertura del libro y operaciones criticas

---

## PASO 1: Preparacion Inicial

### 1.1 Crear Copia de Respaldo

**IMPORTANTE:** Antes de hacer cualquier cambio, crea una copia de respaldo de tu libro Excel.

```
1. Cierra el libro de Excel si esta abierto
2. Haz clic derecho en el archivo
3. Selecciona "Copiar"
4. Pega en el mismo directorio
5. Renombra la copia como "BACKUP_[NombreArchivo]_[Fecha].xlsm"
```

### 1.2 Habilitar Edicion del Proyecto VBA

En Excel:
```
1. Archivo > Opciones > Centro de confianza
2. Clic en "Configuracion del Centro de confianza"
3. Selecciona "Configuracion de macros"
4. Marca "Confiar en el acceso al modelo de objetos de proyectos de VBA"
5. Clic en Aceptar
```

---

## PASO 2: Importar Modulos VBA

### 2.1 Abrir el Editor de VBA

En Excel, presiona `Alt + F11` para abrir el Editor de VBA.

### 2.2 Importar Archivos .bas

Para cada archivo .bas creado, sigue estos pasos:

```
1. En el Editor VBA, ve a Archivo > Importar archivo
2. Navega a la carpeta "Proyecto Acta VBA"
3. Selecciona el archivo correspondiente
4. Clic en Abrir
```

**Archivos a importar (en orden):**

✅ `Modulo_Seguridad.bas` (actualizado)
✅ `Inicializar_Hoja_Config.bas` (nuevo)
✅ `Test_Sistema_Login.bas` (nuevo - temporal)

### 2.3 Actualizar el Formulario InicioSesion

**El formulario ya existe, pero necesita actualizarse:**

1. En el Editor VBA, busca "InicioSesion" en el Explorador de proyectos
2. Elimina el formulario existente (clic derecho > Eliminar)
3. Importa el nuevo formulario:
   - Archivo > Importar archivo
   - Selecciona `InicioSesion.frm`
   - Clic en Abrir

**IMPORTANTE:** Ahora debes agregar un boton de Configuracion al formulario:

```
1. Haz doble clic en "InicioSesion" en el Explorador de proyectos
2. Asegurate de estar en vista de diseno (View > Object)
3. Agrega un CommandButton desde el Cuadro de herramientas
4. Propiedades del boton:
   - Name: cmdConfiguracion
   - Caption: Configuracion
   - Posicion: Coloca donde prefieras (sugerencia: esquina inferior)
5. Cierra el editor de formularios
```

### 2.4 Codigo para ThisWorkbook

**Este codigo NO se puede importar, debe copiarse manualmente:**

```
1. En el Editor VBA, busca "ThisWorkbook" en el Explorador de proyectos
2. Haz doble clic en "ThisWorkbook"
3. Abre el archivo "ThisWorkbook_Codigo.bas" con un editor de texto
4. COPIA todo el codigo (desde Private Sub Workbook_Open hasta End Sub)
5. PEGA el codigo en el editor de "ThisWorkbook"
6. Guarda (Ctrl+S)
```

---

## PASO 3: Crear la Hoja de Configuracion

### 3.1 Ejecutar el Inicializador

```
1. En el Editor VBA, asegurate de que el modulo "Inicializar_Hoja_Config" esta importado
2. Ve a Ver > Ventana Inmediato (o presiona Ctrl+G)
3. En el Editor VBA, busca el modulo "Inicializar_Hoja_Config"
4. Abre el modulo y ubica el procedimiento "CrearHojaConfiguracion"
5. Coloca el cursor dentro del procedimiento
6. Presiona F5 para ejecutar
```

**Resultado esperado:**
- Se crea una hoja llamada "Config_Sistema"
- La hoja contiene 3 usuarios iniciales:
  - admin / 1234
  - usuario1 / pass1
  - usuario2 / pass2
- La hoja queda oculta (xlSheetVeryHidden) y protegida

### 3.2 Verificar la Creacion

```
1. En Excel, verifica que NO veas una hoja llamada "Config_Sistema" 
   (esto es correcto, debe estar oculta)
2. En el Editor VBA, en la ventana Inmediato (Ctrl+G), escribe:
   ? ThisWorkbook.Worksheets("Config_Sistema").Name
3. Presiona Enter
4. Debe mostrar: Config_Sistema
```

---

## PASO 4: Ejecutar Pruebas

### 4.1 Ejecutar la Bateria de Pruebas

```
1. En el Editor VBA, abre el modulo "Test_Sistema_Login"
2. Ubica el procedimiento "EjecutarTodasLasPruebas"
3. Coloca el cursor dentro del procedimiento
4. Presiona F5 para ejecutar
```

### 4.2 Revisar Resultados

```
1. Presiona Ctrl+G para abrir la ventana Inmediato
2. Revisa los resultados de cada prueba
3. Todas deben mostrar "APROBADO"
```

**Ejemplo de salida esperada:**
```
==========================================
INICIANDO BATERIA DE PRUEBAS DEL SISTEMA DE LOGIN
==========================================

[TEST 1] APROBADO - La hoja Config_Sistema existe
[TEST 2] APROBADO - Credenciales 'admin/1234' validadas correctamente
[TEST 3] APROBADO - Credenciales invalidas rechazadas correctamente
[TEST 4] APROBADO - Usuarios cargados: 3
[TEST 5] APROBADO - Todas las hojas estan protegidas
[TEST 6] APROBADO - Usuario inactivo rechazado correctamente
[TEST 7] APROBADO - Todas las funciones auxiliares estan disponibles

==========================================
PRUEBAS COMPLETADAS
==========================================
```

### 4.3 Eliminar el Modulo de Pruebas

**Una vez verificado que todo funciona:**

```
1. En el Editor VBA, clic derecho en "Test_Sistema_Login"
2. Selecciona "Eliminar Test_Sistema_Login"
3. Clic en "No" cuando pregunte si deseas exportarlo
4. El modulo Inicializar_Hoja_Config puede permanecer por si necesitas recrear la hoja
```

---

## PASO 5: Guardar y Probar

### 5.1 Guardar el Libro

```
1. Presiona Ctrl+S en el Editor VBA
2. Cierra el Editor VBA
3. En Excel, guarda el libro (Ctrl+S)
4. Asegurate de que el archivo este en formato .xlsm (con macros)
```

### 5.2 Probar el Sistema

```
1. Cierra completamente el libro Excel
2. Vuelve a abrir el libro
3. Debe aparecer el formulario de inicio de sesion automaticamente
4. Prueba iniciar sesion con:
   - Usuario: admin
   - Contrasena: 1234
5. Si el login es exitoso, deberas ver el mensaje de bienvenida
6. El libro debe estar disponible para trabajar
```

### 5.3 Probar el Cierre por Fallo

```
1. Cierra el libro
2. Vuelve a abrirlo
3. En el formulario de login, haz clic en "Cancelar"
4. El libro debe cerrarse automaticamente
```

---

## PASO 6: Gestion de Usuarios

### 6.1 Acceder a la Hoja de Configuracion

**Solo el usuario "admin" puede hacer esto:**

```
1. Abre el libro y logueate como admin
2. Presiona Alt+F8 para abrir macros
3. Busca y ejecuta la macro "AbrirHojaConfiguracion"
   O
   Si agregaste el boton de Configuracion en el formulario, usalo
```

### 6.2 Modificar Usuarios

Una vez abierta la hoja Config_Sistema:

```
1. Puedes agregar nuevos usuarios en las filas siguientes
2. Puedes modificar contraseñas
3. Puedes desactivar usuarios cambiando "Activo" a "Inactivo"
4. NO elimines las columnas ni los encabezados
```

**Formato de la hoja:**
```
| Usuario    | Contrasena | Estado   |
|------------|------------|----------|
| admin      | 1234       | Activo   |
| usuario1   | pass1      | Activo   |
| usuario2   | pass2      | Inactivo |
```

### 6.3 Cerrar la Hoja de Configuracion

```
1. Cuando termines de editar, presiona Alt+F8
2. Ejecuta la macro "CerrarHojaConfiguracion"
3. La hoja se protegera y ocultara automaticamente
```

---

## Uso en Operaciones Criticas

El sistema ya esta integrado con operaciones criticas existentes como modificar o eliminar registros.

**En el formulario frm_Creacion_Memorias_Modular:**

Cuando el usuario hace doble clic en un registro exportado para modificarlo o eliminarlo:
```
1. El sistema pide autenticacion (AutenticarUsuario)
2. Si falla, solo cancela la operacion (NO cierra el libro)
3. Si tiene exito, permite continuar
```

**Diferencia clave:**
- `AutenticarUsuario(cerrarSiFalla:=True)` → Se usa en Workbook_Open
- `AutenticarUsuario()` o `AutenticarUsuario(cerrarSiFalla:=False)` → Se usa en operaciones criticas

---

## Soluciones a Problemas Comunes

### Problema 1: "Error: No se encontro la hoja de configuracion del sistema"

**Solucion:**
```
1. La hoja Config_Sistema no existe o fue eliminada
2. Ejecuta nuevamente "CrearHojaConfiguracion"
```

### Problema 2: "El libro no se cierra al cancelar el login"

**Solucion:**
```
1. Verifica que el codigo de ThisWorkbook este correctamente implementado
2. Asegurate de usar cerrarSiFalla:=True en Workbook_Open
3. Revisa que Application.EnableEvents este en True
```

### Problema 3: "No puedo ver el boton de Configuracion"

**Solucion:**
```
1. El boton debe agregarse manualmente en el diseño del formulario
2. Sigue las instrucciones del Paso 2.3
3. El codigo del evento ya esta en el formulario
```

### Problema 4: "Olvide la contrasena de admin"

**Solucion de emergencia:**
```
1. Abre el Editor VBA (Alt+F11)
2. Ve a Ver > Ventana Inmediato (Ctrl+G)
3. Ejecuta: Inicializar_Hoja_Config.MostrarHojaConfigManual
4. Ingresa cualquier contraseña que tengas o modifica el código temporalmente
5. Modifica la contraseña en la hoja
6. Ejecuta: Inicializar_Hoja_Config.OcultarHojaConfigManual
```

### Problema 5: "Los logs no se muestran"

**Solucion:**
```
1. Los logs se muestran en la ventana Inmediato del Editor VBA
2. Presiona Ctrl+G en el Editor VBA para verlos
3. Verifica que LOGS_ACTIVOS = True en Modulo_Logs
```

---

## Configuraciones de Seguridad

### Contraseña de Proteccion VBA

La contraseña predeterminada para proteger las hojas es:
```
SistemaSeguridadVBA2024
```

**Para cambiarla:**
```
1. Abre Modulo_Seguridad.bas
2. Busca: Private Const PASSWORD_PROTECCION As String = "SistemaSeguridadVBA2024"
3. Cambia el valor
4. Tambien cambiala en Inicializar_Hoja_Config.bas
5. Guarda y prueba
```

### Nivel de Visibilidad de la Hoja

La hoja Config_Sistema usa `xlSheetVeryHidden`:
- No se puede hacer visible desde el menu normal de Excel
- Solo se puede acceder via VBA

---

## Mantenimiento

### Respaldo de Usuarios

**Recomendacion:** Exporta periodicamente la hoja Config_Sistema:

```
1. Ejecuta AbrirHojaConfiguracion
2. Copia todo el contenido
3. Pega en un archivo seguro
4. Ejecuta CerrarHojaConfiguracion
```

### Actualizaciones Futuras

Si necesitas actualizar el sistema:
```
1. Crea backup del libro
2. Exporta la hoja Config_Sistema (usuarios)
3. Actualiza los modulos
4. Reimporta los usuarios si es necesario
```

---

## Contacto y Soporte

Para problemas o mejoras al sistema, revisa:
- La ventana Inmediato (Ctrl+G) para logs detallados
- El codigo fuente en los modulos VBA
- Este documento de instrucciones

---

**¡Sistema implementado exitosamente!**

El libro ahora requiere autenticacion para ser usado y todas las hojas estan protegidas hasta el login exitoso.

