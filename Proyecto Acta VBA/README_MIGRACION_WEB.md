# Hoja de Ruta: Migracion de Proyecto VBA a Aplicacion Web

## 📜 **1. Vision General y Justificacion**

Este documento describe la hoja de ruta estrategica para migrar el sistema actual, basado en macros de VBA para Excel, a una arquitectura de aplicacion web moderna.

El objetivo principal es superar las limitaciones inherentes de una solucion de escritorio y ofrecer una plataforma mas **escalable, accesible, colaborativa y robusta**. La logica de negocio, la estructura de datos y las funcionalidades ya validadas en el proyecto de VBA (como la gestion de actas, presupuestos, filtros dependientes y exportaciones) constituyen la base fundamental para esta evolucion.

**Beneficios Clave de la Migracion:**
- **Acceso Universal:** Los usuarios podran acceder al sistema desde cualquier dispositivo con un navegador web, sin depender de una version especifica de Excel.
- **Colaboracion en Tiempo Real:** Varios usuarios podran trabajar simultaneamente sobre el mismo conjunto de datos, eliminando la necesidad de compartir y consolidar archivos de Excel.
- **Centralizacion de Datos:** La informacion residira en una base de datos centralizada, garantizando la consistencia, integridad y seguridad de los datos.
- **Escalabilidad:** Una arquitectura web permitira manejar un volumen mucho mayor de datos y usuarios sin degradar el rendimiento.
- **Mantenibilidad y Despliegue Simplificados:** Las actualizaciones se despliegan en el servidor una sola vez y estan disponibles para todos los usuarios instantaneamente.

---

## 🏗️ **2. Arquitectura Tecnologica Propuesta**

Se propone una arquitectura de tres capas, separando la interfaz de usuario (Frontend), la logica de negocio (Backend) y el almacenamiento de datos (Base de Datos).

```mermaid
graph TD
    subgraph "Capa de Presentacion"
        A[Usuario en Navegador Web] --> B{Frontend (React / Vue.js)};
    end
    subgraph "Capa de Negocio"
        B --> C{API RESTful (Python/Django o Node.js/Express)};
    end
    subgraph "Capa de Datos"
        C --> D[Base de Datos (PostgreSQL)];
    end
    subgraph "Servicios de Soporte"
        C --> E[Sistema de Autenticacion (JWT)];
    end
```

### **2.1. Backend (Logica del Servidor)**
- **Proposito:** Reemplazara todas las macros y modulos de VBA. Sera el cerebro de la aplicacion, gestionando la logica de negocio, las validaciones y la comunicacion segura con la base de datos.
- **Tecnologias Sugeridas:**
  - **Python con Django:** Ideal por su enfoque "baterias incluidas", que ofrece un panel de administracion automatico, un ORM potente y un ecosistema maduro para un desarrollo rapido y seguro.
  - **Node.js con Express:** Excelente para aplicaciones que requieren alta concurrencia y operaciones en tiempo real.
- **Funcionalidades a Migrar:**
  - `mod_CrearMemorias.bas`: Se convertira en un endpoint de la API (ej: `POST /api/memorias`) que genere un PDF o un archivo CSV en lugar de una hoja de Excel.
  - `Modulo_Consecutivos.bas`: La logica de generacion de consecutivos se implementara en el backend para garantizar unicidad a nivel de base de datos.
  - `Validaciones.bas`: Las reglas de validacion se ejecutaran en el servidor antes de persistir cualquier dato.

### **2.2. Frontend (Interfaz de Usuario)**
- **Proposito:** Reemplazara los `UserForms` de VBA por una interfaz web moderna, interactiva y responsiva.
- **Tecnologias Sugeridas:**
  - **React o Vue.js:** Ambos son frameworks lideres para construir interfaces de usuario basadas en componentes. Permiten crear elementos reutilizables como tablas de datos, modales y formularios.
- **Funcionalidades a Migrar:**
  - **Filtros Dependientes:** La logica de los `ComboBox` dependientes se replicara utilizando el estado del framework. Al seleccionar una opcion en el primer filtro, se hara una llamada a la API para obtener las opciones relevantes para el segundo.
  - **Tablas de Datos (`ListBox`)**: Se utilizaran librerias como `TanStack Table` (para React) o `Vuetify Data Table` (para Vue) para crear tablas interactivas con paginacion, ordenamiento y busqueda, reemplazando la funcionalidad de los `ListBox`.
  - **Formularios:** Se crearan formularios web para la entrada de datos, con validacion en tiempo real.

### **2.3. Base de Datos**
- **Proposito:** Reemplazara las hojas de Excel (`Acta-Presupuesto`, `EXPORTE_PRESUPUESTO`) como la fuente unica y centralizada de la verdad.
- **Tecnologias Sugeridas:**
  - **PostgreSQL:** Una base de datos relacional de codigo abierto, conocida por su robustez, escalabilidad y estricto cumplimiento del estandar SQL.
- **Diseño de Tablas (Ejemplo Simplificado):**
  - `proyectos`
  - `areas` (con relacion a `proyectos`)
  - `capitulos` (con relacion a `areas`)
  - `actividades` (con catalogo maestro de precios)
  - `actas` (que agrupa un conjunto de `actividades_acta`)
  - `usuarios` y `roles`

---

## 🗺️ **3. Fases de la Migracion**

Se recomienda un enfoque por fases para mitigar riesgos y entregar valor de forma incremental.

### **Fase 1: Nucleo del Backend y Base de Datos**
1.  **Diseñar y crear el esquema** de la base de datos en PostgreSQL.
2.  **Migrar los datos existentes** de las hojas de Excel a la nueva base de datos mediante scripts.
3.  **Desarrollar la API RESTful** con los endpoints basicos para el CRUD (Crear, Leer, Actualizar, Eliminar) de las entidades principales (Areas, Capitulos, Actividades).
4.  **Implementar la autenticacion** de usuarios con JWT.

### **Fase 2: Frontend Inicial y Funcionalidad Basica**
1.  **Desarrollar la interfaz de usuario** para la gestion de actividades y el catalogo.
2.  **Implementar la tabla de datos principal** con filtros y busqueda, replicando la funcionalidad del `ListBox` principal.
3.  **Crear el formulario** para agregar nuevas actividades al presupuesto de un acta.
4.  **Conectar el Frontend** con la API del Backend.

### **Fase 3: Funcionalidades Avanzadas y Exportacion**
1.  **Desarrollar la logica de creacion de "Memorias" o "Actas"**, agrupando actividades.
2.  **Implementar la funcionalidad de exportacion**, generando archivos PDF o CSV en lugar de hojas de Excel.
3.  **Desarrollar el panel de control** con estadisticas y resumenes, similar a `MostrarEstadisticasExportados`.
4.  **Refinar la gestion de usuarios** y permisos.

### **Fase 4: Pruebas, Despliegue y Retirada del Sistema VBA**
1.  **Realizar pruebas exhaustivas** (unitarias, de integracion, de usuario final).
2.  **Desplegar la aplicacion** en un servidor (ej: AWS, Heroku, DigitalOcean).
3.  **Capacitar a los usuarios** en la nueva plataforma.
4.  **Archivar el proyecto de Excel** y operar exclusivamente en la aplicacion web.

---

## 🖥️ **4. Flujo de Trabajo y Vistas de la Aplicacion**

Para traducir la funcionalidad del sistema VBA a una experiencia web, se proponen las siguientes vistas (paginas) y un flujo de trabajo centrado en el usuario.

### **4.1. Vistas Principales de la Aplicacion**

La aplicacion se estructuraria en torno a las siguientes paginas clave:

1.  **Pagina de Inicio de Sesion (`/login`)**:
    - **Proposito:** Reemplaza al formulario `InicioSesion`. Interfaz limpia para que los usuarios ingresen sus credenciales.
    - **Componentes:** Campos para email/usuario, contraseña y boton de "Ingresar". Enlace a "recuperar contraseña".

2.  **Panel de Control / Mis Proyectos (`/dashboard`)**:
    - **Proposito:** Pantalla principal despues de iniciar sesion. Muestra una lista de todas las actas o proyectos en los que el usuario esta trabajando.
    - **Componentes:**
        - Un boton prominente para "Crear Nueva Acta".
        - Una tabla o cuadricula con las actas existentes, mostrando nombre, fecha de creacion, estado (ej: "En Progreso", "Finalizada") y un resumen del valor total.
        - Cada fila tendria acciones rapidas como "Editar", "Ver Resumen" o "Eliminar".

3.  **Espacio de Trabajo del Acta (`/acta/:id`)**:
    - **Proposito:** Esta es la vista principal de trabajo, el equivalente al formulario `frm_Creacion_Memorias_Modular` pero mucho mas potente.
    - **Componentes:**
        - **Seccion de Seleccion de Actividades:**
            - Filtros dependientes para "Area" y "Capitulo" (reemplazan a los `ComboBox`).
            - Una tabla con el "Catalogo de Actividades" que se actualiza dinamicamente segun los filtros. Incluira un boton "Agregar" en cada fila.
        - **Seccion del Acta Actual:**
            - Una tabla principal que muestra las actividades ya agregadas al acta (reemplaza al `ListBox_Trabajo` y `ListBox_Exportados`).
            - Columnas para descripcion, unidad, cantidad (editable), valor unitario y valor parcial.
            - Acciones por fila para "Modificar Cantidad" o "Eliminar del Acta".
        - **Resumen y Acciones Finales:**
            - Un cuadro de resumen que muestra el costo total del acta, actualizado en tiempo real.
            - Botones para "Guardar Progreso", "Generar PDF" o "Finalizar Acta".

4.  **Panel de Administracion (`/admin`)** (para usuarios con permisos):
    - **Proposito:** Un area restringida para gestionar los datos maestros de la aplicacion.
    - **Sub-paginas:**
        - `/admin/actividades`: CRUD completo para el catalogo maestro de actividades.
        - `/admin/usuarios`: Para crear, editar y asignar roles a los usuarios.
        - `/admin/proyectos`: Gestion de proyectos globales.

### **4.2. Diagrama del Flujo de Usuario**

El siguiente diagrama ilustra el recorrido tipico de un usuario al crear una nueva acta de obra.

```mermaid
graph TD
    A(Inicio) --> B[Usuario accede a la URL];
    B --> C{¿Esta autenticado?};
    C -- No --> D[Pagina de Inicio de Sesion];
    D --> E[Ingresa credenciales];
    E -- Exitoso --> F[Panel de Control];
    C -- Si --> F;
    F --> G[Clic en "Crear Nueva Acta"];
    G --> H[Espacio de Trabajo del Acta];
    subgraph "Ciclo de Edicion del Acta"
        H --> I[1. Selecciona Area/Capitulo];
        I --> J[2. Catalogo se filtra];
        J --> K[3. Agrega actividad al acta];
        K --> L[4. Modifica cantidad];
        L --> M[5. Resumen se actualiza];
        M --> I;
    end
    M --> N[Clic en "Generar PDF"];
    N --> O[Descarga el archivo del acta];
    O --> F;
```
