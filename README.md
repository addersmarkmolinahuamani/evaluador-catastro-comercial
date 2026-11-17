# Evaluador de Meta de Gestión: Catastro Comercial

Aplicativo web interactivo de alta precisión para la auditoría, cruce y evaluación de la meta de gestión **"Catastro Comercial"** en Empresas Prestadoras de Servicios de Saneamiento (EPS).

---

## 🎯 Objetivo y Criterios de Evaluación

Las Empresas Prestadoras remiten tres (3) bases de datos independientes en formato Excel (.xlsx / .xls / .csv):

1. **Base de Datos 1: Conexiones Totales (Agua potable y alcantarillado)**
   - Representa el universo del padrón comercial de conexiones existentes.
   - **Criterio de Fecha:** No requiere coincidir con el año de evaluación. Constituye el inventario general base.
2. **Base de Datos 2: Fichas Catastrales en Campo**
   - Conexiones que fueron efectivamente catastradas y censadas físicamente en campo.
   - **Criterio de Fecha:** **Obligatorio dentro del año regulatorio en evaluación**.
3. **Base de Datos 3: Tabla Atributiva de la Plataforma Gráfica GIS**
   - Información cartográfica y georreferenciada del sistema GIS de la empresa.
   - **Criterio de Fecha:** **Obligatorio dentro del año regulatorio en evaluación**.

### 🧩 Lógica de Validación y Cruce Tripartito

| Estado de Validación | En Padrón Totales | Ficha Catastral en Año | Registro GIS en Año | Cumple Meta | Diagnóstico Regulatorio |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **CATASTRADO VÁLIDO** | ✅ SÍ | ✅ SÍ | ✅ SÍ | **CUMPLE** | Cumplimiento pleno: Conexión catastrada en campo e incorporada al sistema GIS dentro del año regulatorio evaluado. |
| **OBSERVADO: SOLO FICHA** | ✅ SÍ | ✅ SÍ | ❌ NO | **NO CUMPLE** | Se realizó levantamiento en campo, pero la EPS no digitalizó ni incorporó la conexión a su plataforma gráfica GIS. |
| **OBSERVADO: SOLO GIS** | ✅ SÍ | ❌ NO | ✅ SÍ | **NO CUMPLE** | Figura en cartografía GIS, pero no cuenta con ficha catastral de campo de respaldo en el periodo evaluado. |
| **OBSERVADO: FUERA DE AÑO** | ✅ SÍ | ⚠️ Fecha anterior/posterior | ⚠️ Fecha anterior/posterior | **NO CUMPLE** | Las fechas de levantamiento o digitalización no corresponden al año regulatorio evaluado. |
| **NO CATASTRADO** | ✅ SÍ | ❌ NO | ❌ NO | **NO CUMPLE** | Conexión activa en padrón pero sin actividades de catastro comercial registradas en el periodo. |
| **NO EN PADRÓN** | ❌ NO | ✅ SÍ | ✅ SÍ | **OBSERVADO** | Conexión registrada en fichas o GIS que no existe en el padrón comercial oficial remitido por la empresa. |

---

## 🚀 Características del Aplicativo

- **Diferenciación y Diagnóstico por Localidad / Sector:**
  - Auto-detección y mapeo del campo `Localidad` (o `Distrito`, `Sector`, `Zona`, `Anexo`).
  - **Tabla Desglose de Meta por Localidad:** Indicadores individuales de Total Conexiones, Validadas (Cumple Meta), Observadas, No Catastradas, porcentaje de avance con barra de progreso, y botón de acción directa *"Ver"* para filtrar la localidad deseada.
  - **Gráfico Comparativo por Localidad:** Comparación visual en barras del total de conexiones del padrón frente a las catastradas válidas para cada localidad o sector.
  - **Filtro Rápido por Localidad:** Menú desplegable dinámico en la barra de herramientas del explorador que recalcula las cantidades por localidad en tiempo real.
  - **Columna Localidad en la Tabla:** Columna ordenable dedicada para identificar geográficamente cada suministro de forma independiente a la dirección.
  - **Exportación en Excel con Desglose por Localidad:** La hoja de *Resumen Ejecutivo* incorpora la tabla comparativa por localidad, y la hoja de *Cruce Consolidado* incluye la columna *Localidad / Sector*.
- **100% Autónomo y Offline:** Todas las librerías (`SheetJS/XLSX`, `Chart.js`, `Lucide Icons`, `Canvas-Confetti`) están integradas localmente en `assets/vendor/`. No requiere conexión a internet ni envía datos fuera de su equipo.
- **Detección Inteligente de Columnas (Auto-Mapping):** Reconoce automáticamente variaciones de nombres de columnas como:
  - `Número de suministro`, `Numero de suministro`, `NIS`, `Suministro`
  - `Tipo de servicio`, `Servicio`
  - `Fecha de ficha catastral`, `Fecha de registro en el sistema grafico`, `Fecha`
  - `Usuario`, `Titular`, `Dirección`, `Código catastral`
- **Selector de Año Regulatorio Previo:** Define el año evaluado (ej. 2025) o rangos personalizados de fechas como primer paso fundamental.
- **Tablero de Control Ejecutivo (Dashboard):**
  - **KPIs:** Tasa de Cumplimiento de Meta (% y barra de avance), Total Conexiones, Validadas, Observadas, No Catastradas.
  - **Gráficos Interactivos:**
    1. Gráfico circular (Donut) con la distribución porcentual de estados de validación.
    2. Gráfico comparativo de barras por Tipo de Servicio (Agua y Alcantarillado, Solo Agua, Solo Alcantarillado).
    3. Cronograma mensual (Ene - Dic) de avance de fichas en campo vs incorporaciones en GIS.
- **Explorador Dinámico de Conexiones:**
  - Búsqueda en tiempo real por número de suministro, usuario, dirección o código.
  - Filtros rápidos por estado (`Cumple Meta`, `Solo Ficha`, `Solo GIS`, `Fuera de Año`, `No Catastrado`).
  - Filtro por tipo de servicio.
  - Paginación y ordenamiento por cualquier columna.
  - **Modal de Auditoría Individual:** Inspecciona lado a lado los datos exactos que figuran en cada una de las 3 bases de datos para un suministro específico.
- **Exportación en Excel (.xlsx):**
  - **Base Consolidada:** Libro Excel con pestaña de *Resumen Ejecutivo* e indicadores, pestaña con el *Cruce Consolidado* completo y columna de diagnóstico.
  - **Reporte de Observaciones:** Descarga directa de todas las conexiones no conformes para notificación oficial a la EPS.
- **Botón "Cargar Datos de Ejemplo":** Permite probar y familiarizarse con toda la funcionalidad del aplicativo en 1 solo clic.

---

## 💻 Formas de Ejecución

### Opción 1: Ejecución Directa (Sin dependencias)
Haga doble clic en el archivo [`index.html`](file:///c:/Users/acer/Desktop/Antigravity/Indicadores%20MG/Catastro%20comercial/index.html) para abrirlo en su navegador preferido (Google Chrome, Microsoft Edge, Firefox, etc.).

### Opción 2: Servidor Local Node.js
Si desea ejecutarlo con servidor web local:
```bash
npm start
# o
npm run dev
```
Abra en su navegador: `http://localhost:3000`
