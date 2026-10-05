# Documentación Técnica: Dashboard de Aulas Virtuales 2026
**Plataforma eSalud — Ministerio de Salud y Deportes de Mendoza**  
**Dirección de Modernización**  
**Versión:** 2.0.0  
**Fecha de corte:** Octubre 2026  
**Tecnología:** HTML5 / CSS3 / JavaScript (ES6+) / Chart.js v4.5.1 / SheetJS v0.18.5 / Python 3  

---

## 1. Introducción y Alcance del Sistema

El **Dashboard de Aulas Virtuales 2026** es una aplicación analítica web autocontenida (**Single-File Component / Zero-Backend**) diseñada para el monitoreo, auditoría, actualización dinámica y toma de decisiones operativas sobre los espacios formativos (cursos y cápsulas) alojados en la plataforma Moodle del Campus Virtual de Salud de la Provincia de Mendoza.

### 1.1 Objetivos de la Solución
* **Visibilidad Unificada y Analítica Avanzada:** Centralizar en una única interfaz interactiva el estado operativo, el tipo de recurso formativo, la modalidad de matriculación, el volumen de cursantes, promedio de cohorte y tasa de completitud de datos.
* **Portabilidad Absoluta (Zero-Backend):** Funcionar de forma 100% autónoma sin requerir servidores Node.js, PHP ni bases de datos relacionales activas; puede ejecutarse abriendo el archivo localmente (`file://`), en buckets estáticos o en GitHub Pages / servidores web institucionales.
* **Ingesta Dinámica de Datos (Drag & Drop en Navegador):** Capacidad de procesar y recalcular el dashboard instantáneamente al arrastrar o seleccionar una nueva planilla Excel (`.xlsx`) en caliente gracias al motor cliente de SheetJS.
* **Suite Completa de Exportación:** Descarga local de la versión autónoma actualizada en HTML, exportación de datos a Excel (`.xlsx`), CSV estructurado UTF-8 con BOM, resumen ejecutivo en Markdown (`.md`) e impresión/PDF estilizada.
* **Actualización Asistida por CLI:** Script local `update_dashboard.py` que permite procesar planillas Excel por terminal y compilar el HTML oficial de forma desatendida.

---

## 2. Modelo y Estructura de Datos

El conjunto de datos reside en el propio documento HTML encapsulado dentro de una etiqueta `<script id="data-holder" type="application/json">`. Al cargar la página, el motor JavaScript instancia y parsea este objeto JSON en memoria. Además, puede ser reemplazado dinámicamente si el usuario arrastra un nuevo archivo Excel sobre la interfaz.

### 2.1 Esquema de Cada Registro (`AulaRecord`)

| Campo | Tipo | Nullable | Descripción y Reglas de Negocio |
| :--- | :--- | :---: | :--- |
| `id` | `Integer` | No | Identificador numérico único y secuencial ($1 \dots N$). |
| `aula` | `String` | No | Denominación formal del espacio formativo en la plataforma Moodle. |
| `tipo` | `String` | No | Clasificación del recurso: `'Curso'` o `'Capsula'`. |
| `categoria` | `String` | No | Área o Dirección ministerial responsable (15 valores normalizados). |
| `acceso` | `String` | No | Política de matriculación: `'Solo cursantes'` (cerrada/matriculada) o `'Libre'` (autoinscripción). |
| `participantes` | `Float` / `null` | Sí | Valor numérico de cursantes/inscriptos. `null` si aún no hay cohorte abierta o está en diseño. |
| `participantes_raw`| `String` | No | Texto crudo de origen: `""` (próxima apertura) o `"SIN"` (en diseño didáctico). |
| `fecha` | `String` (ISO) | Sí | Formato `YYYY-MM-DD`. `null` para aulas permanentes o sin fecha de inicio definida. |
| `mes` | `String` (ISO) | Sí | Formato `YYYY-MM`. Utilizado para la agregación de series cronológicas mensuales. |
| `fecha_raw` | `String` | Sí | Registro textual original de la fecha (e.g., `'2026-10-13'`, `'NO TIENE'`, `'oct 14, 2026'`). |
| `estado` | `String` | No | Estado del ciclo de vida: `'Iniciado'`, `'Permanente'`, `'Terminado'`, `'No Iniciado'`, `'Sin Fecha Inicio'`. |

### 2.2 Normalización de Entidades (`ORG_CANONICAL`)
Las categorías de origen son estandarizadas respecto al registro crudo de la planilla para mantener coherencia institucional:
* `DIR. MODERNIZACION` $\rightarrow$ `Dirección de Modernización`
* `DIR. EPIDEMIOLOGIA` $\rightarrow$ `Dirección de Epidemiología`
* `DIR. MATERNIDAD` $\rightarrow$ `Dirección de Maternidad e Infancia`
* `SALUD MENTAL` $\rightarrow$ `Dirección de Salud Mental`
* `EDUCACION PARA LA SALUD` $\rightarrow$ `Depto. Educación para la Salud`
* `LAB.NUCLEO` $\rightarrow$ `Laboratorio Núcleo (LABNU)`
* `CIENCIA Y TECNICA` $\rightarrow$ `DICYT (Investigación, Ciencia y Técnica)`
* `BIOESTADISTICA` $\rightarrow$ `Departamento de Bioestadística`
* `DIR. H.ALIMENTOS` $\rightarrow$ `Depto. Higiene e Inocuidad Alimentaria`
* `CAEP- CURSO ANUAL DE ENFERMERIA PRACTICA` $\rightarrow$ `CAEP (Enfermería Práctica)`
* `SUBSECRETARIA` $\rightarrow$ `Subsecretaría de Salud`
* `INMUNIZACIONES` / `Inmunizaciones` $\rightarrow$ `Depto. Inmunizaciones`
* `PLANIFICACION` $\rightarrow$ `Dirección de Planificación`
* `DIR. DISCAPACIDAD` $\rightarrow$ `Dirección de Discapacidad`
* `EXTERNO` $\rightarrow$ `Organismos Externos`

---

## 3. Lógica de Negocio y Reglas Operativas

### 3.1 Semántica de los Estados Operativos
1. **Permanente:** Espacios formativos y repositorios autoadministrados vigentes de forma continua durante todo el año. No poseen fechas de inicio ni de finalización de cohorte obligatorias.
2. **Iniciado:** Cursos con cohorte activa en plataforma, con fecha confirmada y estudiantes matriculados.
3. **Terminado:** Cursos cuya cohorte culminó sus actividades lectivas durante el ciclo 2026 (incluye aulas puente que iniciaron a fines del período anterior).
4. **No Iniciado:** Cursos planificados y estructurados didácticamente que cuentan con fecha de inicio confirmada, listos para apertura de inscripciones. Su matrícula numérica es `null` y figura como `Próx. inscr.`
5. **Sin Fecha Inicio:** Propuestas formativas previstas para el año 2026 pero aún en etapa de diseño curricular, sin fecha de apertura. Su campo figura como `En diseño`.

### 3.2 Indicadores Clave de Rendimiento (KPIs)
* **Total Participantes Acumulado:**
  $$\text{Matrícula} = \sum_{i=1}^{N} \text{participantes}_i \quad \forall \; \text{participantes}_i \neq \text{null}$$
  Base actual: **6.101 participantes**.
* **Promedio de Cohorte:**
  $$\overline{X}_{\text{cohorte}} = \frac{\sum \text{participantes}_i}{\text{Aulas con participantes registrados}}$$
* **Completitud de Matrícula (%):**
  $$\text{Completitud} = \left( \frac{\text{Aulas con valor numérico}}{\text{Total de aulas activas}} \right) \times 100$$
  Permite monitorear la madurez de la carga operativa y la calidad del dato institucional.

---

## 4. Arquitectura del Frontend y Componentes

La interfaz adopta el sistema de diseño institucional moderno:

```
+---------------------------------------------------------------------------------+
|  HEADER: Marca institucional, título, resumen y barra de filtros multidimensional |
+---------------------------------------------------------------------------------+
|  DATA TOOLS: Suite de exportación (HTML, Excel .xlsx, CSV, Markdown, Print/PDF) |
+---------------------------------------------------------------------------------+
|  DROPZONE: Ingesta Drag & Drop de planillas Excel (.xlsx) con SheetJS en caliente|
+---------------------------------------------------------------------------------+
|  QUICK FILTERS: Chips de acceso rápido por estado, tipo o modalidad             |
+---------------------------------------------------------------------------------+
|  ACTIVE FILTER BANNER: Etiquetas descartables y botón de limpieza general       |
+---------------------------------------------------------------------------------+
|  KPI ROWS: Primarios (Estados de Aulas) + Secundarios (Matrícula, Promedio, %)   |
+---------------------------------------------------------------------------------+
|  CHART ROW:                                                                     |
|    - [Ancho Completo] Gráfico 1: Aulas por Dirección desglosadas por Estado     |
|    - [Mitad Izq.]     Gráfico 2: Participantes por Dirección / Área             |
|    - [Mitad Der.]     Gráfico 3: Tipo de Recurso vs. Modo de Acceso             |
|    - [Ancho Completo] Gráfico 4: Cronograma Mensual de Cohortes 2026            |
+---------------------------------------------------------------------------------+
|  TABLE SECTION: Detalle paginado con ordenación visual (▲/▼) y badges           |
+---------------------------------------------------------------------------------+
|  TOAST NOTIFICATIONS: Avisos flotantes de confirmación o alerta de proceso      |
+---------------------------------------------------------------------------------+
|  FOOTER: Metadatos de auditoría y procedencia de datos                          |
+---------------------------------------------------------------------------------+
```

### 4.1 Modo Alto Contraste (`.hc`)
Al conmutar el botón `◐ Contraste`, se eliminan las sombras difusas y se activan bordes reforzados de alto contraste conformes a las pautas de accesibilidad WCAG AAA.

### 4.2 Soporte de Impresión Ejecutiva (`@media print`)
La hoja de estilos incluye directivas de impresión que ocultan controles interactivos (filtros, botones, dropzone, paginadores), ajustan la escala de canvas a hojas tamaño A4 y preservan la fidelidad cromática institucional (`print-color-adjust: exact`).

---

## 5. Ingesta Dinámica y Flujo de Actualización

### 5.1 En el Navegador (Vía SheetJS)
1. El usuario arrastra una planilla `.xlsx` a la zona de dropzone (o hace clic en `Seleccionar archivo .xlsx`).
2. El lector `FileReader` convierte el binario en `ArrayBuffer`.
3. `XLSX.read()` procesa el libro y extrae la matriz de celdas.
4. El método `parseRawExcelRows()` detecta la cabecera, normaliza las dependencias, repara corrimientos de columnas en origen, parsea fechas y recalcula el estado.
5. Se invoca `dash.refresh()`, actualizando en tiempo real KPIs, gráficos y tablas sin recargar la página.
6. El usuario puede hacer clic en **`⬇️ Descargar HTML Autónomo`** para guardar una copia local del dashboard con los datos recién inyectados.

### 5.2 Vía Consola (Script Python `update_dashboard.py`)
Para flujos de integración continua o mantenimiento periódico:
```bash
python update_dashboard.py --excel course-report-table.xlsx --html dashboard_de_aulas_virtuales_2026.html
```
El script lee la estructura XML del archivo `.xlsx`, aplica la misma taxonomía de limpieza e inyecta el JSON resultante en el bloque `<script id="data-holder">`.

---

## 6. Canales de Exportación Disponibles

1. **`⬇️ Descargar HTML Autónomo`:** Clona el DOM actualizando el script de datos embebido para crear un nuevo archivo `.html` 100% independiente y listo para distribución.
2. **`📊 Exportar a Excel (.xlsx)`:** Genera un archivo `.xlsx` estructurado con los datos filtrados en pantalla usando `XLSX.writeFile()`.
3. **`📄 Exportar CSV`:** Descarga un archivo delimitado por punto y coma con encoding UTF-8 y marca BOM (`\uFEFF`) para compatibilidad directa con Microsoft Excel en español.
4. **`📝 Resumen Markdown (.md)`:** Produce un reporte ejecutivo con métricas consolidadas y tabla en formato Markdown estándar.
5. **`🖨️ Imprimir / Guardar PDF`:** Lanza el diálogo nativo del navegador con maquetación optimizada para impresión o guardado como documento PDF de alta calidad.