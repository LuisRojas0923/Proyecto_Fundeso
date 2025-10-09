# ☑️ Checklist de Implementacion - Sistema de Autenticacion

Marca cada item conforme lo vayas completando.

---

## 📦 PREPARACION

- [ ] **Crear copia de respaldo del libro Excel**
  - Archivo copiado: `_________________`
  - Fecha: `_________________`

- [ ] **Habilitar acceso al modelo de objetos VBA**
  - Excel > Opciones > Centro de confianza > Configuracion del Centro de confianza
  - Configuracion de macros > Marcar "Confiar en el acceso..."

- [ ] **Abrir Editor VBA**
  - Presionar Alt+F11

---

## 📥 IMPORTACION DE MODULOS

- [ ] **Importar Modulo_Seguridad.bas**
  - Archivo > Importar archivo
  - Seleccionar: `Modulo_Seguridad.bas`
  - Estado: ☐ Pendiente | ☐ Completado | ☐ Error: `_________________`

- [ ] **Importar Inicializar_Hoja_Config.bas**
  - Archivo > Importar archivo
  - Seleccionar: `Inicializar_Hoja_Config.bas`
  - Estado: ☐ Pendiente | ☐ Completado | ☐ Error: `_________________`

- [ ] **Importar Test_Sistema_Login.bas**
  - Archivo > Importar archivo
  - Seleccionar: `Test_Sistema_Login.bas`
  - Estado: ☐ Pendiente | ☐ Completado | ☐ Error: `_________________`

---

## 📋 FORMULARIO INICIO SESION

- [ ] **Eliminar formulario InicioSesion existente**
  - Clic derecho en "InicioSesion" > Eliminar

- [ ] **Importar nuevo formulario InicioSesion.frm**
  - Archivo > Importar archivo
  - Seleccionar: `InicioSesion.frm`

- [ ] **Agregar boton de Configuracion al formulario**
  - Abrir InicioSesion en modo diseño (View > Object)
  - Agregar CommandButton desde Cuadro de herramientas
  - Propiedades del boton:
    - Name: `cmdConfiguracion`
    - Caption: `Configuracion`
  - Posicion del boton: `_________________`

- [ ] **Verificar que el formulario se ve correcto**
  - Controles visibles: cmbUsuario, txtPassword, cmdLogin, cmdCancelar, chkMostrar, cmdConfiguracion

---

## 📄 CODIGO THISWORKBOOK

- [ ] **Abrir archivo ThisWorkbook_Codigo.bas con editor de texto**
  - Notepad, VSCode, etc.

- [ ] **Copiar todo el codigo del archivo**
  - Desde `Option Explicit` hasta el ultimo `End Sub`

- [ ] **Pegar en el modulo ThisWorkbook del Editor VBA**
  - En Explorador de proyectos, doble clic en "ThisWorkbook"
  - Pegar el codigo copiado

- [ ] **Guardar**
  - Ctrl+S en el Editor VBA

- [ ] **Verificar eventos creados**
  - [ ] Workbook_Open
  - [ ] Workbook_BeforeClose
  - [ ] Workbook_BeforeSave

---

## 🔧 CREAR HOJA CONFIG_SISTEMA

- [ ] **Ubicar el procedimiento CrearHojaConfiguracion**
  - Modulo: Inicializar_Hoja_Config
  - Procedimiento: CrearHojaConfiguracion

- [ ] **Ejecutar el procedimiento**
  - Cursor dentro del procedimiento
  - Presionar F5

- [ ] **Verificar mensaje de confirmacion**
  - Mensaje mostrado: ☐ Si | ☐ No
  - Contenido del mensaje: `_________________`

- [ ] **Verificar que la hoja NO sea visible**
  - En Excel, no debe aparecer pestaña "Config_Sistema"
  - Esto es CORRECTO (hoja muy oculta)

- [ ] **Verificar via VBA que existe**
  - Ventana Inmediato (Ctrl+G)
  - Escribir: `? ThisWorkbook.Worksheets("Config_Sistema").Name`
  - Resultado: `Config_Sistema`

---

## 🧪 PRUEBAS DEL SISTEMA

- [ ] **Ejecutar bateria de pruebas**
  - Modulo: Test_Sistema_Login
  - Procedimiento: EjecutarTodasLasPruebas
  - Presionar F5

- [ ] **Abrir ventana Inmediato**
  - Ctrl+G

- [ ] **Verificar resultados de pruebas**

  **Test 1 - Existencia de hoja:**
  - [ ] APROBADO | [ ] FALLIDO
  - Detalles: `_________________`

  **Test 2 - Credenciales correctas:**
  - [ ] APROBADO | [ ] FALLIDO
  - Detalles: `_________________`

  **Test 3 - Credenciales incorrectas:**
  - [ ] APROBADO | [ ] FALLIDO
  - Detalles: `_________________`

  **Test 4 - Carga de usuarios:**
  - [ ] APROBADO | [ ] FALLIDO
  - Usuarios cargados: `_________________`

  **Test 5 - Proteccion de hojas:**
  - [ ] APROBADO | [ ] FALLIDO
  - Hojas protegidas: `_________________`

  **Test 6 - Usuario inactivo:**
  - [ ] APROBADO | [ ] FALLIDO
  - Detalles: `_________________`

  **Test 7 - Funciones auxiliares:**
  - [ ] APROBADO | [ ] FALLIDO
  - Detalles: `_________________`

- [ ] **TODAS LAS PRUEBAS APROBADAS**
  - ☐ Si, continuar | ☐ No, revisar errores

---

## 🗑️ LIMPIEZA

- [ ] **Eliminar modulo Test_Sistema_Login**
  - Clic derecho en "Test_Sistema_Login"
  - Eliminar Test_Sistema_Login
  - Al preguntar si exportar: NO

- [ ] **Opcional: Mantener Inicializar_Hoja_Config**
  - Este modulo puede ser util en el futuro
  - Decide: ☐ Mantener | ☐ Eliminar

---

## 💾 GUARDAR Y PROBAR

- [ ] **Guardar el libro**
  - Ctrl+S en Excel
  - Verificar formato: .xlsm (con macros)

- [ ] **Cerrar el libro completamente**
  - Archivo > Cerrar

- [ ] **Reabrir el libro**
  - Doble clic en el archivo

- [ ] **Verificar que aparece formulario de login**
  - ☐ Si | ☐ No (revisar ThisWorkbook)

- [ ] **Prueba 1: Login exitoso**
  - Usuario: `admin`
  - Contraseña: `1234`
  - Resultado: ☐ Acceso concedido | ☐ Error: `_________________`

- [ ] **Cerrar y reabrir libro**

- [ ] **Prueba 2: Login cancelado**
  - Hacer clic en "Cancelar"
  - Resultado esperado: Libro se cierra automaticamente
  - Resultado real: ☐ Se cerro | ☐ No se cerro (revisar codigo)

- [ ] **Cerrar y reabrir libro**

- [ ] **Prueba 3: Credenciales incorrectas**
  - Usuario: `admin`
  - Contraseña: `incorrecta`
  - Resultado esperado: Mensaje de error, NO se cierra el libro
  - Resultado real: ☐ Correcto | ☐ Error: `_________________`

- [ ] **Login correcto para continuar**

---

## 🔐 PRUEBAS DE CONFIGURACION

- [ ] **Loguearse como admin**
  - Usuario: `admin`
  - Contraseña: `1234`

- [ ] **Acceder a configuracion**
  - Metodo 1: Boton en formulario (si lo agregaste)
  - Metodo 2: Alt+F8 > AbrirHojaConfiguracion

- [ ] **Verificar que se abre la hoja Config_Sistema**
  - Debe ser visible ahora
  - Columnas: Usuario | Contrasena | Estado

- [ ] **Verificar usuarios iniciales**
  - [ ] admin / 1234 / Activo
  - [ ] usuario1 / pass1 / Activo
  - [ ] usuario2 / pass2 / Activo

- [ ] **Agregar un usuario de prueba**
  - Usuario: `test`
  - Contraseña: `test123`
  - Estado: `Activo`

- [ ] **Cerrar hoja de configuracion**
  - Alt+F8 > CerrarHojaConfiguracion
  - Resultado: ☐ Hoja oculta | ☐ Error: `_________________`

- [ ] **Cerrar y reabrir libro**

- [ ] **Probar login con nuevo usuario**
  - Usuario: `test`
  - Contraseña: `test123`
  - Resultado: ☐ Funciona | ☐ Error: `_________________`

---

## 🔄 PRUEBAS DE INTEGRACION

- [ ] **Abrir formulario frm_Creacion_Memorias**
  - (Si aplica a tu sistema)

- [ ] **Intentar modificar/eliminar un registro**
  - Hacer doble clic en lista de exportados

- [ ] **Verificar que solicita autenticacion**
  - ☐ Si solicita | ☐ No solicita (revisar codigo)

- [ ] **Cancelar la autenticacion**
  - Resultado esperado: Operacion cancelada, libro permanece abierto
  - Resultado real: ☐ Correcto | ☐ Error: `_________________`

- [ ] **Intentar nuevamente y autenticarse**
  - Resultado esperado: Permite la operacion
  - Resultado real: ☐ Correcto | ☐ Error: `_________________`

---

## 📊 VERIFICACION FINAL

- [ ] **Sistema de logs funciona**
  - Alt+F11 > Ctrl+G (Ventana Inmediato)
  - Verificar que hay mensajes de log
  - Formato: [TIMESTAMP] [NIVEL] [PROCEDIMIENTO] - [MENSAJE]

- [ ] **Proteccion de hojas funciona**
  - Antes de login: ☐ Hojas protegidas
  - Despues de login: ☐ Hojas desprotegidas

- [ ] **Usuarios activos/inactivos funciona**
  - Crear usuario inactivo en Config_Sistema
  - Intentar login: ☐ Rechazado correctamente

- [ ] **Cierre automatico funciona**
  - Cancelar login: ☐ Libro se cierra
  - Credenciales incorrectas + cerrar: ☐ Libro permanece abierto

---

## 🎉 COMPLETADO

- [ ] **Todas las pruebas exitosas**

- [ ] **Sistema funcionando correctamente**

- [ ] **Documentacion revisada**
  - [ ] INSTRUCCIONES_IMPLEMENTACION_LOGIN.md
  - [ ] README_SISTEMA_LOGIN.md
  - [ ] RESUMEN_IMPLEMENTACION.md

- [ ] **Contraseñas documentadas en lugar seguro**
  - Ubicacion: `_________________`

- [ ] **Copia de respaldo creada**
  - Ubicacion: `_________________`

---

## 📝 NOTAS Y OBSERVACIONES

```
Anota aqui cualquier problema, duda o mejora:

_________________________________________________________________

_________________________________________________________________

_________________________________________________________________

_________________________________________________________________

_________________________________________________________________
```

---

## ✅ FIRMA DE COMPLETADO

- **Implementado por:** `_________________`
- **Fecha:** `_________________`
- **Tiempo total:** `_________________`
- **Estado final:** ☐ Exitoso | ☐ Exitoso con observaciones | ☐ Fallido

---

**¡Felicidades! Sistema de autenticacion implementado correctamente.**

