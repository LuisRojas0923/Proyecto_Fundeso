# Proyecto Acta VBA - Sistema de Gestión de Memorias

## 📋 Última Actualización: 16 de agosto de 2025

### 🆕 Cambios Recientes
- Simplificación del sistema de filtros (solo palabra clave)
- Nuevo formato de precios con símbolo $
- Implementación de numeración automática
- Sistema de exportación a hojas

## 🏗️ Estructura Actual del Sistema

### 📊 ListBox Principal (6 columnas)
1. Valor de Item_1 (editable)
2. Numeración automática
3. Datos tabla col.1
4. Datos tabla col.2
5. Datos tabla col.3
6. Precios (formato $)

### � Sistema de Filtrado
- Filtro único por palabra clave
- Búsqueda en columna 2 de la tabla origen
- Actualización dinámica de resultados

### 🔘 Botones Principales
- **btn_RegistrarDatos**: 
  - Asigna Item_1 a columna 1
  - Genera numeración consecutiva
  - Opera sobre selección múltiple
- **btn_Marcar/Desmarcar**: Gestión de selección
- **Crea_Presupuesto**: Exportación a hoja

### 📐 Características Técnicas
- ListBox con ancho fijo (750 pts)
- Formato monetario en precios
- Selección múltiple habilitada
- Monitoreo de dimensiones vía Debug

### 📁 Archivos del Sistema
- `frm_Creacion_Memorias.frm`: UI principal
- `ExportarSeleccionados.bas`: Módulo exportación
- `README.md`: Documentación

### 📊 Origen de Datos
Tabla: ListaPrecios_PreciosClientes
Columnas utilizadas: 4

### ⏳ Próximas Actualizaciones

#### 🔄 Visualización de Exportaciones
- Nuevo formulario para visualizar datos exportados
  - ListBox con datos exportados
  - Botón para llamar desde formulario principal
  - Capacidad de navegación y revisión
  - Actualización en tiempo real

#### 🔍 Sistema de Validación
- Validador de duplicados con 3 llaves
  - Prevención de duplicados antes de exportar
  - Mensaje de advertencia al usuario
  - Opción de sobrescribir o cancelar

#### 📋 Nuevos Controles de Filtrado
- Lista desplegable de Área
  - Integración con datos existentes
  - Filtrado dinámico de registros
  - Actualización del ListBox principal

- Lista desplegable de Capítulo
  - Sincronización con selección de Área
  - Filtrado contextual
  - Validación de selecciones

#### 🎯 Prioridades de Implementación
1. Sistema de validación de duplicados
2. Formulario de visualización
3. Nuevas listas desplegables
4. Integración y pruebas

## 🔧 Características Técnicas

### Manejo de Errores
- Error handling en cada procedimiento
- Debug logging para monitoreo
- Validaciones de datos

### Interfaz de Usuario
- Diseño optimizado
- Controles responsivos
- Formato consistente
- **Centrado automático**: Los formularios se centran automáticamente en la pantalla
- **Responsive design**: Ajuste dinámico del ancho del ListBox según el contenido
- **Validación visual**: Mensajes informativos para el usuario

## 📊 Flujo de Trabajo

### 1. Inicio de Sesión
```
Usuario ingresa credenciales → Validación → Acceso al sistema principal
```

### 2. Creación de Memorias
```
Selección de Item → Carga de registros → Selección de fechas → 
Registro de datos → Exportación (opcional)
```

### 3. Gestión de Ausentismo
```
Selección de registros → Definición de período → Validación → 
Registro en sistema → Confirmación
```

## ⚙️ Configuración y Uso

### Prerrequisitos
- Microsoft Excel con habilitación de macros
- Acceso a las hojas de trabajo: "Consolidado Memorias" y "ListaPrecios_PreciosClientes"
- Formulario de calendario (`frmCalendario_`) para selección de fechas

### Instalación
1. Importar los archivos .frm y .frx al proyecto VBA
2. Verificar que las hojas de trabajo requeridas existen
3. Configurar los usuarios y contraseñas según necesidades
4. Habilitar macros en Excel

### Uso Básico
1. **Iniciar sesión** con credenciales válidas
2. **Seleccionar item** del ComboBox para cargar registros
3. **Definir fechas** usando los campos de fecha (integración con calendario)
4. **Seleccionar registros** específicos del ListBox
5. **Registrar datos** o **exportar** según necesidades

## 🛠️ Mantenimiento

### Funciones de Utilidad
- `LimpiarControlesFormulario()`: Limpia todos los controles del formulario
- `ActualizarControlesOpciones()`: Actualiza el estado de los controles
- `GuardarFDesde()` / `GuardarFHasta()`: Funciones estáticas para almacenar fechas

### Personalización
- Modificar usuarios en `UserForm_Initialize()` del formulario de login
- Ajustar anchos de columnas en la configuración del ListBox
- Personalizar validaciones de fecha según requerimientos empresariales

## 📈 Características Avanzadas

- **Selección múltiple inteligente**: Control granular de selección en ListBox
- **Validación de coherencia**: Verificación automática de rangos de fechas
- **Integración con calendario visual**: Selección intuitiva de fechas
- **Exportación personalizada**: Generación de reportes en formato Excel
- **Logging y debugging**: Sistema completo de trazabilidad de errores

## 🏢 Información Corporativa

**Desarrollado para**: Fundeso  
**Tipo de sistema**: Gestión de memorias y control de ausentismo  
**Plataforma**: Microsoft Excel VBA  
**Versión**: 5.00

---

*Este sistema ha sido diseñado específicamente para las necesidades operativas de Fundeso, proporcionando una solución integral para la gestión de memorias de trabajo y control de ausentismo del personal.*
