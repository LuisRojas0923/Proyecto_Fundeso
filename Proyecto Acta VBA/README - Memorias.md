# Proyecto de Automatización y Consolidación en Excel

Este proyecto contiene un conjunto de macros y módulos VBA diseñados para automatizar tareas de consolidación, exportación y gestión de datos en libros de Excel, especialmente orientado a la manipulación de tablas Power Query y la generación de reportes personalizados.

## Archivos principales

| Archivo                        | Descripción breve                                                                 |
|------------------------------- |---------------------------------------------------------------------------------|
| `Macro Principal.bas`          | Lógica principal del formulario: filtrado, selección y registro de fechas en ListBox. |
| `mod_CrearMemorias.bas`        | Exporta registros seleccionados del ListBox a nuevas hojas usando una plantilla. |
| `mod_RegistrarDatos.bas`       | Registra fechas en las columnas del ListBox para los ítems seleccionados.        |
| `mod_LimpiarControlesFormulario.bas` | Limpia todos los campos del formulario.                                    |
| `mod_MostrarFormulario.bas`    | Centra y muestra el formulario principal en pantalla.                            |
| `ACTUALIZAR.bas`               | Macro para actualizar todas las tablas de Power Query del libro.                 |
| `Exporte_Memorias.bas`         | Exporta hojas con números en el nombre a un nuevo libro.                        |
| `ListarConsultas.bas`          | Lista todas las consultas y conexiones de Power Query en una hoja resumen.       |
| `Consolidado.pq`               | Script Power Query para consolidar varias tablas en una sola.                    |
| `clsEtiquetaFecha.cls`         | Clase para manejo de etiquetas de fecha en formularios.                         |
| `FuncionFinal.vb`              | Exporta registros del ListBox a hojas nuevas, validando nombres y controlando errores. |

## Origen de los Datos

La tabla `EXPORTE_PRESUPUESTO` proviene de otro archivo llamado `acta presupuesto`. Es fundamental mantener actualizado este archivo de origen para asegurar la integridad de los datos procesados.

## Requisitos
- Microsoft Excel (recomendado 2016 o superior)
- Habilitar macros (VBA)
- Permitir acceso a objetos de proyecto VBA si se requiere depuración

## Uso básico
1. **Abrir el archivo Excel** que contiene estos módulos y macros.
2. **Habilitar macros** cuando Excel lo solicite.
3. **Usar el formulario principal** para filtrar, seleccionar y registrar fechas en los registros.
4. **Exportar registros** seleccionados a nuevas hojas o consolidar datos según las opciones del formulario.
5. **Actualizar tablas Power Query** usando la macro correspondiente si es necesario.

## Personalización
- Puedes modificar los anchos de columna del ListBox en `Macro Principal.bas` y `mod_RegistrarDatos.bas`.
- Los nombres de las hojas y celdas de destino pueden adaptarse según tu estructura de datos.

## Soporte
Para dudas, mejoras o reportes de errores, contacta al desarrollador o abre un issue en tu sistema de control de versiones si aplica.

## Actualizaciones Recientes (Resumen de la Sesión)

Durante la última sesión de desarrollo, se implementaron las siguientes mejoras significativas:

1.  **Módulo de Logging Centralizado (`mod_Logger.bas`)**:
    *   Se creó un módulo dedicado para gestionar todos los mensajes de depuración (`Debug.Print`).
    *   Permite activar o desactivar todos los logs desde una única constante (`LOGGING_ACTIVO`), mejorando la mantenibilidad y facilitando el paso a producción.
    *   Se refactorizó todo el código existente para utilizar las nuevas funciones de logging (`LogInfo`, `LogErrorVBA`, `LogDebug`, etc.).

2.  **Sistema de Filtros Dependientes de Dos Niveles**:
    *   Se implementó un segundo ComboBox (`cmb_Capitulo`) que funciona como un filtro dependiente del primero (`cmb_ITEMS`).
    *   Al seleccionar un "ITEM", se carga el `ListBox` y se puebla el filtro "Capítulo" con las opciones relevantes (basado en columnas 3 y 4).
    *   Al seleccionar un "Capítulo", se vuelve a filtrar el `ListBox` para mostrar resultados que cumplen ambas condiciones.
    *   Se corrigieron errores de duplicados y bucles de eventos mediante la desactivación temporal de eventos (`Application.EnableEvents`).

3.  **Mejoras en la Interfaz de Usuario (`ListBox`)**:
    *   **Reordenamiento y Rediseño**: Se modificó el orden y los anchos de las columnas del `ListBox` para una mejor visualización de los datos.
    *   **Ancho Fijo y Estable**: Se solucionó un problema donde el `ListBox` cambiaba de tamaño al aplicar filtros. Ahora tiene un ancho fijo (`ANCHO_LISTBOX`) que se establece al iniciar el formulario, garantizando una interfaz estable.
    *   **Centralización de Estilos**: Se crearon constantes (`ANCHO_LISTBOX`, `ANCHOS_COLUMNAS`) para gestionar el diseño del `ListBox` desde un único lugar, siguiendo el principio DRY.

4.  **Personalización de Hojas Creadas**:
    *   **Nombre de Hoja Dinámico**: La lógica para nombrar las nuevas hojas se ajustó para concatenar los valores de las columnas 1, 3 y 5, usando un punto (`.`) como separador.
    *   **Dato Adicional en Celda `B4`**: Se añadió la lógica para que el valor de la columna 2 ("AREA") de la hoja `EXPORTE_PRESUPUESTO` se escriba automáticamente en la celda `B4` de cada nueva memoria.

5.  **Optimización del Procesamiento de Datos**:
    *   Se ajustó el código para que la carga de datos en los filtros y en el `ListBox` comience desde la **segunda fila**, ignorando los encabezados de la tabla `EXPORTE_PRESUPUESTO`.

---
¡Gracias por usar este proyecto de automatización en Excel! 