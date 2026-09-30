# Control de Asistencias

Sistema integral y ligero para el procesamiento, limpieza, deduplicacion y analisis de registros biometricos y de asistencia laboral a partir de archivos Excel (`.xlsx`, `.xls`) y CSV.

Disenado para ejecutarse **100% en el navegador del cliente (Client-Side)** sin necesidad de servidores o bases de datos, garantizando la **privacidad total de los datos** y una respuesta instantanea.

---

## Caracteristicas Principales

* **Deteccion Automatica de Formatos**:
  * Formato tradicional de 5 columnas (`Nombre, Fecha, Hora, Periodo, Observaciones`).
  * Formato unificado de 4 columnas (`Nombre, Fecha, Hora y Periodo, Observaciones`).
  * Formato de 24 horas / militar (`07:45:00`, `18:30:00`).
* **Algoritmo Inteligente de Deduplicacion**:
  * Detecta y filtra marcajes repetidos o accidentales con segundos de diferencia.
  * Asigna automaticamente el primer marcaje del dia como **Entrada (Llegada)** y el ultimo como **Salida (Ida)**.
  * Clasifica casos especiales: *Completo*, *Solo Entrada (Falta Salida)*, *Solo Salida (Falta Entrada)*.
* **Calculo de Ratios y Faltas de Marca**:
  * Parametros configurables en vivo para Entradas y Salidas totales esperadas del periodo.
  * Calculo exacto de faltas netas contra la base esperada.
  * Opcion de exclusion automatica de fines de semana (sabados y domingos).
* **Interfaz Web Moderna y Profesional**:
  * Paleta de colores descansada para la vista (Gris Perla y Azul Ejecutivo).
  * Zona interactiva de carga con soporte **Drag & Drop**.
  * Cabeceras de tabla fijas con efecto esmerilado translucido (`backdrop-blur`).
  * Transicion suave al pasar el cursor sobre las filas (`row-smooth`).
  * Filtros en tiempo real por empleado, fecha y estado de marcaje.
* **Vista Previa de Impresion Integrada**:
  * Modal interactivo estilo documento institucional.
  * Adaptado con reglas `@media print` para imprimir en hojas limpias con paginacion continua.
* **Exportacion a Excel Limpio (`.xlsx`)**:
  * Genera un libro de calculo con dos hojas formateadas: **Resumen General** y **Detalle Diario**.
* **Script CLI de PowerShell**:
  * Automatizacion por linea de comandos para procesar lotes de archivos de forma desatendida.

---

## Estructura del Proyecto

```text
├── index.html                           # Aplicacion Web completa (Single-Page App)
├── app_logo.png                         # Logo oficial de la aplicacion
├── app_icon.ico                         # Icono oficial embebido en el ejecutable
├── procesar_asistencias.ps1             # Script de automatizacion en PowerShell
├── ejemplo_formato_5_columnas.xlsx      # Archivo de prueba (5 columnas)
├── ejemplo_formato_4_columnas_ampm.xlsx # Archivo de prueba (4 columnas AM/PM)
├── ejemplo_formato_4_columnas_24h.xlsx  # Archivo de prueba (4 columnas 24h)
├── .gitignore                           # Exclusion de datos y reportes locales
└── README.md                            # Documentacion del proyecto
```

---

## Descarga y Ejecucion (.exe para Windows)

Puedes utilizar la aplicacion de dos formas:

### 1. Aplicacion de Escritorio Independiente (`.exe`) *(Recomendado para Windows)*
No necesitas instalar navegadores adicionales ni dependencias. Descarga el ejecutable autonomo y abrelo directamente:

* **[Descargar ControlDeAsistencias.exe (Ultima version)](https://github.com/pontro/control-de-asistencias/releases/latest/download/ControlDeAsistencias.exe)**
* **[Ver todas las versiones y notas de la version (GitHub Releases)](https://github.com/pontro/control-de-asistencias/releases)**

> **Nota para Windows:** Al abrir por primera vez un ejecutable nuevo descargado de internet, es posible que Windows SmartScreen muestre una advertencia informativa. Haz clic en *"Mas informacion"* y luego en *"Ejecutar de todas formas"*.
> La aplicacion incluye un comprobador automatico de versiones que te notificara dentro del programa cuando exista una nueva actualizacion.

---

## Inicio Rapido (Version Web)

1. Abre el archivo [`index.html`](file:///c:/Users/Nando%20Pontro/Desktop/Projects/isaacnewton/index.html) en cualquier navegador web moderno (Google Chrome, Microsoft Edge, Firefox, Safari).
2. Arrastra tu archivo de asistencias (`.xlsx`, `.xls` o `.csv`) a la zona de carga o haz clic para seleccionarlo.
3. Ajusta las **Entradas y Salidas Esperadas** del periodo en el panel lateral segun el calendario laboral.
4. Explora las pestanas:
   * **Resumen por Empleado**: Cumplimiento general, entradas, salidas, dias asistidos y faltas de marca.
   * **Detalle Diario**: Registro historico desglosado dia a dia con filtros avanzados.
5. Haz clic en **Vista Previa del Reporte** para inspeccionar o imprimir, o en **Descargar Excel Limpio (.xlsx)** para exportar.

---

## Uso desde Terminal (PowerShell)

Para procesar archivos masivos por consola sin abrir el navegador:

```powershell
# Ejecucion basica
.\procesar_asistencias.ps1 -RutaArchivo "asistencias.csv"

# Con parametros personalizados y exclusion de fines de semana
.\procesar_asistencias.ps1 -RutaArchivo "asistencias.csv" -TotalEntradasEsperadas 18 -TotalSalidasEsperadas 17 -ExcluirFinesDeSemana
```

---

## Formatos de Entrada Admitidos

### Formato 1: 5 Columnas
| Columna A | Columna B | Columna C | Columna D | Columna E |
| :--- | :--- | :--- | :--- | :--- |
| **Nombre** | **Fecha** | **Hora** | **Periodo** | **Observaciones** |
| Juan Perez | 01/09/2026 | 07:45:00 | a. m. | Entrada puntual |
| Juan Perez | 01/09/2026 | 06:10:00 | p. m. | Salida regular |

### Formato 2: 4 Columnas (Hora y Periodo juntos)
| Columna A | Columna B | Columna C | Columna D |
| :--- | :--- | :--- | :--- |
| **Nombre** | **Fecha** | **Hora y Periodo** | **Observaciones** |
| Juan Perez | 01/09/2026 | 07:45:00 a. m. | Entrada puntual |
| Juan Perez | 01/09/2026 | 06:10:00 p. m. | Salida regular |

### Formato 3: 4 Columnas (Hora en formato 24h)
| Columna A | Columna B | Columna C | Columna D |
| :--- | :--- | :--- | :--- |
| **Nombre** | **Fecha** | **Hora (24h)** | **Observaciones** |
| Juan Perez | 01/09/2026 | 07:45:00 | Entrada |
| Juan Perez | 01/09/2026 | 18:10:00 | Salida |

---

## Privacidad y Seguridad

* **Sin almacenamiento en la nube**: Ningun dato o archivo se envia a servidores externos.
* **Procesamiento Local**: Todo el analisis y la generacion de archivos Excel ocurren en la memoria del navegador del usuario via SheetJS.

---

## Licencia

Este proyecto esta bajo la Licencia [MIT](https://opensource.org/licenses/MIT).
