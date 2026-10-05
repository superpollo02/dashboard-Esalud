# Dashboard de Aulas Virtuales 2026 — Plataforma eSalud

**Ministerio de Salud y Deportes — Gobierno de Mendoza**  
**Dirección de Modernización**

Aplicación analítica interactiva y autocontenida (**Zero-Backend / Single-File**) para el monitoreo, auditoría, visualización y actualización operativa de los espacios formativos (cursos y cápsulas) alojados en el Campus Virtual de Salud.

---

## 🚀 Características Principales

* **Portabilidad Absoluta:** Funciona como un único archivo HTML autónomo (`dashboard_de_aulas_virtuales_2026.html`) sin necesidad de base de datos ni servidores de aplicaciones. Compatible con apertura local (`file://`) y GitHub Pages.
* **Actualización Dinámica con Excel (SheetJS):** Zona de *Drag & Drop* en el navegador para arrastrar un archivo `.xlsx` y recalcular métricas y gráficos en caliente.
* **Script CLI en Python (`update_dashboard.py`):** Utilidad para procesar planillas Excel desde consola e inyectar automáticamente los datos consolidados en el HTML.
* **Suite Completa de Exportaciones:**
  * ⬇️ Descargar HTML autónomo con datos actualizados
  * 📊 Exportar a Excel (`.xlsx`)
  * 📄 Exportar a CSV (UTF-8 con BOM)
  * 📝 Resumen ejecutivo en Markdown (`.md`)
  * 🖨️ Impresión optimizada / Guardar PDF (A4)
* **Analítica y Filtros Cruzados:** Gráficos interactivos en Chart.js v4 vinculados bidireccionalmente con KPIs (promedio de cohorte, completitud de carga, estado operativo) y tabla paginada con ordenamiento.

---

## 📂 Estructura del Repositorio

* `dashboard_de_aulas_virtuales_2026.html`: Tablero web interactivo principal.
* `update_dashboard.py`: Script CLI para normalización e inyección de datos desde planillas Excel.
* `course-report-table.xlsx`: Planilla de reporte de origen.
* `documento_t_cnico_configuraci_n_y_funcionamiento_del_dashboard_de_aulas_virtuales_2026.md`: Documentación técnica detallada (v2.0.0).

---

## 🛠️ Uso Rápido

### Visualización
Basta con abrir `dashboard_de_aulas_virtuales_2026.html` en cualquier navegador web moderno (Google Chrome, Microsoft Edge, Firefox, etc.).

### Actualización por Terminal
```bash
python update_dashboard.py --excel course-report-table.xlsx --html dashboard_de_aulas_virtuales_2026.html
```
