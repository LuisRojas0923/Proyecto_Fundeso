# Manual de Usuario - Sistema de Gestión de Actas
## Fundeso - Sistema de Memorias y Control de Ausentismo

---

**Versión del Sistema:** 5.00  
**Fecha de Actualización:** Agosto 2025  
**Desarrollado para:** Fundeso  
**Plataforma:** Microsoft Excel VBA  

---

## 📋 Índice

1. [Introducción al Sistema](#1-introducción-al-sistema)
2. [Requisitos del Sistema](#2-requisitos-del-sistema)
3. [Inicio de Sesión](#3-inicio-de-sesión)
4. [Interfaz Principal](#4-interfaz-principal---creación-de-memoriasactas)
5. [Proceso de Creación de Acta](#5-proceso-paso-a-paso---creación-de-acta)
6. [Funcionalidades Avanzadas](#6-funcionalidades-avanzadas)
7. [Solución de Problemas](#7-solución-de-problemas-comunes)
8. [Preguntas Frecuentes](#8-preguntas-frecuentes-faq)
9. [Glosario de Términos](#9-glosario-de-términos)
10. [Contacto y Soporte](#10-información-de-contacto-y-soporte)

---

## 1. Introducción al Sistema

### ¿Qué es el Sistema de Gestión de Actas?

El **Sistema de Gestión de Actas** es una herramienta desarrollada específicamente para **Fundeso** que permite:

- ✅ **Crear memorias de trabajo** de manera automatizada
- ✅ **Gestionar el control de ausentismo** del personal
- ✅ **Exportar reportes** en formato Excel
- ✅ **Filtrar y organizar** información de presupuestos
- ✅ **Mantener un registro** centralizado de actividades

### Características Principales

- **🔐 Sistema de autenticación** con usuarios y contraseñas
- **📊 Interfaz intuitiva** con filtros avanzados
- **📅 Integración con calendario** para selección de fechas
- **📈 Numeración automática** de actividades
- **💾 Exportación automática** a hojas Excel
- **🔍 Búsqueda inteligente** por palabra clave

---

## 2. Requisitos del Sistema

### Software Requerido

| Componente | Requisito Mínimo | Recomendado |
|------------|------------------|--------------|
| **Microsoft Excel** | 2016 o superior | 2019/365 |
| **Sistema Operativo** | Windows 10 | Windows 11 |
| **Memoria RAM** | 4 GB | 8 GB o más |
| **Espacio en Disco** | 100 MB libres | 500 MB libres |

### Configuración Necesaria

- ✅ **Macros habilitadas** en Excel
- ✅ **Permisos de VBA** activados
- ✅ **Acceso a hojas de trabajo** del sistema
- ✅ **Permisos de escritura** en la carpeta del archivo

### ⚠️ Importante

> **Nota de Seguridad:** El sistema requiere que las macros estén habilitadas para funcionar correctamente. Si Excel muestra una advertencia de seguridad, seleccione **"Habilitar contenido"** para continuar.

---

## 3. Inicio de Sesión

### Acceso al Sistema

1. **Abrir el archivo Excel** del sistema
2. **Habilitar macros** cuando Excel lo solicite
3. **Aparecerá automáticamente** el formulario de inicio de sesión

![Pantalla de Inicio de Sesión](capturas/01_inicio_sesion.png)
*Figura 1: Formulario de inicio de sesión del sistema*

### Proceso de Autenticación

#### Paso 1: Seleccionar Usuario
- Haga clic en el **ComboBox de usuarios**
- Seleccione su nombre de usuario de la lista desplegable

![Selección de Usuario](capturas/02_seleccion_usuario.png)
*Figura 2: Selección de usuario del sistema*

#### Paso 2: Ingresar Contraseña
- Haga clic en el campo **"Contraseña"**
- Escriba su contraseña (aparecerá oculta con asteriscos)
- Use el checkbox **"Mostrar contraseña"** si necesita verificar lo que escribió

![Ingreso de Contraseña](capturas/03_ingreso_password.png)
*Figura 3: Campo de contraseña con opción de mostrar*

#### Paso 3: Iniciar Sesión
- Haga clic en el botón **"Login"** para acceder al sistema
- O haga clic en **"Cancelar"** para salir

### Opciones de Usuario

| Usuario | Contraseña | Permisos |
|---------|------------|----------|
| **admin** | 1234 | Acceso completo + configuración |
| **usuario1** | pass1 | Acceso estándar |
| **usuario2** | pass2 | Acceso estándar |

### 🔧 Acceso a Configuración (Solo Administradores)

Si es administrador, puede acceder a la configuración de usuarios:

1. **Ingrese sus credenciales** de administrador
2. **Haga clic en "Configuración"** (botón adicional)
3. **Se abrirá la hoja** de gestión de usuarios

![Botón de Configuración](capturas/04_boton_configuracion.png)
*Figura 4: Botón de configuración para administradores*

---

## 4. Interfaz Principal - Creación de Memorias/Actas

### Vista General del Formulario

Una vez autenticado, accederá al **formulario principal** del sistema:

![Interfaz Principal](capturas/05_interfaz_principal.png)
*Figura 5: Vista general del formulario principal*

### Componentes de la Interfaz

#### 🔍 **Sección de Filtros**

| Campo | Descripción | Uso |
|-------|-------------|-----|
| **Palabra Clave** | Búsqueda por texto libre | Escriba palabras para filtrar registros |
| **Área** | Filtro por área de trabajo | Seleccione el área específica |
| **Capítulos** | Filtro por capítulo | Dependiente del área seleccionada |

![Sección de Filtros](capturas/06_seccion_filtros.png)
*Figura 6: Sección de filtros del sistema*

#### 📋 **ListBox de Registros**

El **ListBox principal** muestra los registros disponibles con las siguientes columnas:

| Columna | Descripción | Formato |
|---------|-------------|---------|
| **1** | Código del Item | Texto (ej: "1.2.14") |
| **2** | Numeración automática | Número consecutivo |
| **3** | Datos de la tabla (Col.1) | Texto descriptivo |
| **4** | Datos de la tabla (Col.2) | Área de trabajo |
| **5** | Datos de la tabla (Col.3) | Capítulo específico |
| **6** | Precios | Formato monetario ($) |

![ListBox de Registros](capturas/07_listbox_registros.png)
*Figura 7: ListBox con registros y columnas*

#### 🎛️ **Botones de Selección**

| Botón | Función | Descripción |
|-------|---------|-------------|
| **Marcar** | Seleccionar todos | Marca todos los registros visibles |
| **Desmarcar** | Deseleccionar todos | Quita la selección de todos los registros |

#### 📅 **Campos de Fecha**

| Campo | Descripción | Integración |
|-------|-------------|-------------|
| **Fecha Desde** | Fecha de inicio del período | Calendario visual |
| **Fecha Hasta** | Fecha de fin del período | Calendario visual |

![Campos de Fecha](capturas/08_campos_fecha.png)
*Figura 8: Campos de fecha con integración de calendario*

#### ⚡ **Botones de Acción**

| Botón | Función | Descripción |
|-------|---------|-------------|
| **Registrar Datos** | Procesar selección | Registra los datos seleccionados |
| **Exportar** | Generar hoja Excel | Crea una nueva hoja con los datos |
| **Limpiar Campos** | Reiniciar formulario | Limpia todos los campos |

---

## 5. Proceso Paso a Paso - Creación de Acta

### Flujo de Trabajo Completo

#### **Paso 1: Configurar Filtros**

1. **Escriba una palabra clave** en el campo correspondiente (opcional)
2. **Seleccione un área** del ComboBox
3. **Elija un capítulo** (se cargará automáticamente según el área)

![Configuración de Filtros](capturas/09_configuracion_filtros.png)
*Figura 9: Configuración de filtros paso a paso*

#### **Paso 2: Seleccionar Registros**

1. **Revise los registros** mostrados en el ListBox
2. **Seleccione los registros** que desea incluir:
   - Haga clic individual en cada registro
   - O use **"Marcar"** para seleccionar todos
   - Use **"Desmarcar"** para quitar selecciones

![Selección de Registros](capturas/10_seleccion_registros.png)
*Figura 10: Proceso de selección de registros*

#### **Paso 3: Definir Fechas**

1. **Haga clic en "Fecha Desde"**
2. **Seleccione la fecha** en el calendario que aparece
3. **Repita el proceso** para "Fecha Hasta"

![Selección de Fechas](capturas/11_seleccion_fechas.png)
*Figura 11: Calendario para selección de fechas*

#### **Paso 4: Registrar Datos**

1. **Verifique** que todos los campos estén correctos
2. **Haga clic en "Registrar Datos"**
3. **Confirme** la operación en el mensaje que aparece

![Registro de Datos](capturas/12_registro_datos.png)
*Figura 12: Proceso de registro de datos*

#### **Paso 5: Exportar (Opcional)**

1. **Haga clic en "Exportar"** para generar una hoja Excel
2. **El sistema creará** una nueva hoja con los datos
3. **La hoja tendrá** un nombre descriptivo basado en los datos

![Proceso de Exportación](capturas/13_proceso_exportacion.png)
*Figura 13: Exportación a hoja Excel*

---

## 6. Funcionalidades Avanzadas

### 🔄 **Filtros Dependientes**

El sistema implementa **filtros inteligentes** que se actualizan automáticamente:

1. **Al seleccionar un Área** → Se cargan los capítulos correspondientes
2. **Al seleccionar un Capítulo** → Se filtran los registros relevantes
3. **Búsqueda por palabra clave** → Filtra en tiempo real

![Filtros Dependientes](capturas/14_filtros_dependientes.png)
*Figura 14: Funcionamiento de filtros dependientes*

### 🔢 **Numeración Automática**

- **Consecutivos automáticos** para cada área
- **Numeración secuencial** por capítulo
- **Códigos únicos** para cada actividad

### 💰 **Formato de Precios**

- **Símbolo de moneda** ($) automático
- **Formato numérico** estándar
- **Cálculos automáticos** de totales

### 🧹 **Limpieza de Campos**

El botón **"Limpiar Campos"** realiza:

- ✅ Limpia todos los ListBox
- ✅ Resetea los filtros
- ✅ Borra las fechas
- ✅ Vuelve a la página inicial
- ✅ Recarga los datos

### 📄 **Sistema MultiPage**

El formulario incluye **múltiples páginas** para organizar las funciones:

- **Página 1**: Filtros y selección principal
- **Página 2**: Gestión de trabajo
- **Página 3**: Exportaciones y reportes

---

## 7. Solución de Problemas Comunes

### ❌ **Error: "Usuario o contraseña incorrectos"**

**Causa:** Credenciales incorrectas o usuario inactivo

**Solución:**
1. Verifique que el usuario esté en la lista
2. Confirme que la contraseña sea correcta
3. Contacte al administrador si el problema persiste

### ❌ **Error: "No se encontró la hoja de configuración"**

**Causa:** La hoja `Config_Sistema` no existe o fue eliminada

**Solución:**
1. Contacte al administrador del sistema
2. Ejecute la macro de inicialización
3. Verifique que el archivo esté completo

### ❌ **ListBox vacío o sin datos**

**Causa:** Problema con la tabla de datos origen

**Solución:**
1. Verifique que la tabla `EXPORTE_PRESUPUESTO` exista
2. Actualice los datos de origen
3. Use el botón "Limpiar Campos" y reintente

### ❌ **Error al exportar**

**Causa:** Permisos insuficientes o archivo bloqueado

**Solución:**
1. Cierre otras instancias de Excel
2. Verifique permisos de escritura
3. Guarde el archivo antes de exportar

### ❌ **Macros deshabilitadas**

**Causa:** Configuración de seguridad de Excel

**Solución:**
1. Vaya a **Archivo → Opciones → Centro de confianza**
2. Haga clic en **"Configuración del Centro de confianza"**
3. Seleccione **"Configuración de macros"**
4. Marque **"Habilitar todas las macros"**

### ❌ **Problemas de permisos**

**Causa:** Restricciones de seguridad del sistema

**Solución:**
1. Ejecute Excel como administrador
2. Verifique permisos de la carpeta
3. Contacte al administrador de TI

---

## 8. Preguntas Frecuentes (FAQ)

### 🔐 **Autenticación y Usuarios**

**P: ¿Cómo recupero mi contraseña?**
R: Contacte al administrador del sistema. Solo él puede restablecer contraseñas.

**P: ¿Puedo cambiar mi contraseña?**
R: No directamente. El administrador debe hacerlo desde la configuración del sistema.

**P: ¿Qué hago si mi usuario no aparece en la lista?**
R: Contacte al administrador para que agregue su usuario al sistema.

### 📊 **Uso del Sistema**

**P: ¿Qué hago si el sistema no me deja seleccionar registros?**
R: Verifique que haya registros cargados y que los filtros estén configurados correctamente.

**P: ¿Cómo actualizo los datos de origen?**
R: Los datos se actualizan automáticamente desde la tabla `EXPORTE_PRESUPUESTO`. Contacte al administrador si necesita actualizar esta tabla.

**P: ¿Puedo modificar una memoria ya creada?**
R: Sí, puede editar las hojas exportadas directamente en Excel.

**P: ¿Dónde se guardan las memorias exportadas?**
R: Se crean como nuevas hojas dentro del mismo archivo Excel del sistema.

### 🔧 **Problemas Técnicos**

**P: ¿Por qué el calendario no aparece?**
R: Verifique que el formulario de calendario esté instalado correctamente.

**P: ¿Qué hago si el sistema se cuelga?**
R: Cierre Excel completamente y vuelva a abrir el archivo del sistema.

**P: ¿Puedo usar el sistema en otra computadora?**
R: Sí, pero debe tener Excel con macros habilitadas y acceso al archivo del sistema.

---

## 9. Glosario de Términos

### **Términos del Sistema**

| Término | Definición |
|---------|------------|
| **Acta/Memoria** | Documento generado que registra actividades y fechas específicas |
| **ListBox** | Lista desplegable que muestra registros para selección |
| **ComboBox** | Campo desplegable para seleccionar opciones predefinidas |
| **Filtros Dependientes** | Sistema donde la selección de un filtro afecta las opciones de otros |
| **Exportación** | Proceso de crear una nueva hoja Excel con los datos seleccionados |
| **Power Query** | Herramienta de Excel para conectar y transformar datos |
| **VBA/Macros** | Lenguaje de programación que automatiza tareas en Excel |
| **Config_Sistema** | Hoja oculta que contiene la configuración de usuarios del sistema |

### **Términos Técnicos**

| Término | Definición |
|---------|------------|
| **MultiPage** | Control que permite tener múltiples páginas en un formulario |
| **Consecutivo** | Número secuencial automático asignado a cada actividad |
| **Validación** | Proceso que verifica que los datos ingresados sean correctos |
| **Logging** | Sistema de registro de actividades para auditoría |

---

## 10. Información de Contacto y Soporte

### 👨‍💼 **Administrador del Sistema**

**Nombre:** [Nombre del Administrador]  
**Email:** [email@fundeso.com]  
**Teléfono:** [Número de contacto]  
**Horario de Atención:** Lunes a Viernes, 8:00 AM - 5:00 PM  

### 🆘 **Procedimiento para Reportar Errores**

1. **Documente el error** con capturas de pantalla
2. **Anote los pasos** que llevaron al error
3. **Contacte al administrador** con la información
4. **Espere confirmación** de recepción

### 💡 **Solicitud de Nuevas Funcionalidades**

1. **Describa la necesidad** específica
2. **Explique el beneficio** esperado
3. **Proporcione ejemplos** de uso
4. **Envíe la solicitud** al administrador

### 📞 **Horarios de Soporte**

| Día | Horario | Disponibilidad |
|-----|---------|----------------|
| **Lunes - Jueves** | 8:00 AM - 5:00 PM | Soporte completo |
| **Viernes** | 8:00 AM - 3:00 PM | Soporte limitado |
| **Fines de semana** | No disponible | Solo emergencias |

### 📧 **Canales de Comunicación**

- **Email:** [soporte@fundeso.com]
- **Teléfono:** [Número de soporte]
- **Chat interno:** [Sistema de mensajería corporativa]
- **Tickets:** [Sistema de tickets de soporte]

---

## 📝 **Notas Finales**

### ✅ **Mejores Prácticas**

- **Guarde su trabajo** regularmente
- **Cierre sesión** cuando termine
- **Reporte errores** inmediatamente
- **Mantenga actualizado** el archivo del sistema

### 🔄 **Actualizaciones del Sistema**

El sistema se actualiza regularmente. Las nuevas versiones incluyen:
- Corrección de errores
- Nuevas funcionalidades
- Mejoras de rendimiento
- Actualizaciones de seguridad

### 📚 **Recursos Adicionales**

- **Manual técnico** para administradores
- **Videos tutoriales** en la intranet corporativa
- **Base de conocimientos** con artículos detallados
- **Foro de usuarios** para compartir experiencias

---

**© 2025 Fundeso - Sistema de Gestión de Actas v5.00**  
*Este manual está diseñado para usuarios finales del sistema de gestión de actas de Fundeso.*

---

*Para preguntas sobre este manual o sugerencias de mejora, contacte al administrador del sistema.*
