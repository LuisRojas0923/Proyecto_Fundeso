# Automatización Fundeso

## 📋 Descripción del Proyecto

**Automatización Fundeso** es un sistema integral de automatización desarrollado en VBA para la gestión y procesamiento de memorias presupuestarias. El sistema está diseñado para optimizar los procesos de creación, actualización y exportación de datos presupuestarios en entornos corporativos.

## 🏗️ Arquitectura del Sistema

El proyecto está organizado en dos módulos principales:

### 📁 Proyecto Acta VBA
Sistema principal de gestión de actas y presupuestos con funcionalidades avanzadas de:
- Actualización automática de tablas con consultas Power Query
- Sistema de logs centralizado y robusto
- Gestión modular de formularios y controles
- Exportación y sincronización de datos

### 📁 Proyecto Memorias VBA
Sistema complementario para la gestión de memorias con:
- Creación automatizada de memorias presupuestarias
- Consolidación de datos mediante Power Query
- Exportación especializada de información
- Navegación y selección avanzada de registros

## 🚀 Características Principales

### ✨ Funcionalidades Core
- **Actualización Automática**: Sistema robusto para actualizar tablas de presupuesto con manejo de errores
- **Sistema de Logs**: Registro centralizado con niveles (ERROR, WARNING, INFO, DEBUG) y timestamps
- **Navegación por Tabs**: Orden lógico de navegación con teclado optimizado para UX
- **Exportación Modular**: Procesos de exportación con validación y confirmación
- **Gestión de Errores**: Manejo robusto de errores con logging detallado

### 🔧 Módulos Especializados
- **Modulo_Logs**: Sistema centralizado de logging con niveles configurables
- **Modulo_Actualizacion_Tablas**: Actualización automática de consultas Power Query
- **Modulo_Exportacion**: Procesos de exportación con validación
- **Modulo_Trabajo**: Gestión de área de trabajo y selección de registros
- **Modulo_Sincronizacion**: Sincronización de datos entre sistemas

## 📖 Documentación

### 📚 Manuales Disponibles
- **MANUAL_USUARIO_COMPLETO.md**: Guía completa del usuario con todas las funcionalidades
- **README - Memorias.md**: Documentación específica del módulo de memorias
- **README_MIGRACION_WEB.md**: Guía para migración a sistemas web

### 🔍 Guías de Uso
1. **Configuración Inicial**: Verificar dependencias y configurar entorno
2. **Procesos de Actualización**: Uso del sistema de actualización automática
3. **Exportación de Datos**: Procedimientos para exportar información
4. **Gestión de Logs**: Configuración y monitoreo del sistema de logging

## 🛠️ Requisitos del Sistema

### 📋 Software Requerido
- Microsoft Excel 2016 o superior
- Microsoft Office con soporte para VBA
- Acceso a Power Query (opcional para funcionalidades avanzadas)

### 🔧 Configuración Necesaria
- Habilitar macros en Excel
- Configurar referencias de objetos necesarias
- Establecer rutas de archivos según el entorno

## 🚀 Instalación y Configuración

### 1. Clonar el Repositorio
```bash
git clone https://github.com/[usuario]/automatizacion-fundeso.git
cd automatizacion-fundeso
```

### 2. Configurar Excel
1. Abrir Excel y habilitar macros
2. Importar los módulos VBA desde los archivos .bas
3. Configurar referencias necesarias
4. Ejecutar configuración inicial

### 3. Configuración de Logs
1. Verificar configuración en `Modulo_Logs.bas`
2. Ajustar niveles de logging según necesidades
3. Configurar rutas de archivos de log

## 📊 Estructura de Archivos

```
automatizacion-fundeso/
├── Proyecto Acta VBA/
│   ├── Modulo_Logs.bas                    # Sistema centralizado de logs
│   ├── Modulo_Actualizacion_Tablas.bas    # Actualización automática
│   ├── Modulo_Exportacion.bas             # Procesos de exportación
│   ├── Modulo_Trabajo.bas                 # Gestión de área de trabajo
│   ├── frm_Creacion_Memorias_Modular.frm  # Formulario principal
│   └── ...                                # Otros módulos especializados
├── Proyecto Memorias VBA/
│   ├── Macro Principal.bas                # Macro principal del sistema
│   ├── mod_CrearMemorias.bas              # Creación de memorias
│   ├── Exporte_Memorias.bas               # Exportación especializada
│   └── ...                                # Otros módulos
├── MANUAL_USUARIO_COMPLETO.md             # Documentación principal
└── README.md                              # Este archivo
```

## 🔧 Configuración Avanzada

### Sistema de Logs
```vba
' Configuración de niveles de log
Public Const LOGS_ACTIVOS As Boolean = True
Public Const NIVEL_LOG_MAXIMO As Integer = 3 ' LOG_INFO

' Uso en el código
RegistrarInfo "NombreProcedimiento", "Mensaje informativo"
RegistrarError "NombreProcedimiento", "Mensaje de error"
```

### Navegación por Tabs
El sistema incluye navegación optimizada por teclado:
- **Página 1**: Selección de registros (TabIndex 1-8)
- **Página 2**: Área de trabajo (TabIndex 10-14)
- **Página 3**: Revisión (TabIndex 20+)

## 🤝 Contribución

### 📝 Guías de Desarrollo
1. **Estilo de Código**: Seguir las mejores prácticas de VBA establecidas
2. **Documentación**: Documentar todos los procedimientos públicos
3. **Logging**: Usar el sistema centralizado de logs
4. **Manejo de Errores**: Implementar manejo robusto en todos los procedimientos

### 🔄 Flujo de Trabajo
1. Fork del repositorio
2. Crear rama para nueva funcionalidad
3. Implementar cambios siguiendo estándares
4. Crear pull request con documentación

## 📞 Soporte y Contacto

Para soporte técnico o consultas sobre el proyecto:
- Revisar documentación en `MANUAL_USUARIO_COMPLETO.md`
- Verificar logs del sistema para diagnóstico
- Consultar issues en el repositorio

## 📄 Licencia

Este proyecto está desarrollado para uso interno de Fundeso. Todos los derechos reservados.

## 🏷️ Versión

**Versión Actual**: 1.0.0  
**Última Actualización**: Enero 2024

---

*Desarrollado con ❤️ para optimizar los procesos de Fundeso*
