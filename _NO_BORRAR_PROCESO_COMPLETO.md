# PROCESO COMPLETO — BITÁCORA GGR / NOTARÍA 31

> **IMPORTANTE:** No eliminar ni mover este archivo. Es la guía maestra del proceso.
> Última actualización: Mayo 2026

---

## ARCHIVOS DEL PROYECTO

### Carpeta: `/Users/ke7/Documents/Claude/Projects/Bitácoras Notaría 31/`

| Archivo | Rol | ¿Tocar? |
|---|---|---|
| `Bitacora.xlsx` | Fuente de datos principal. Se actualiza con cada nueva escritura. | SÍ (es la bitácora activa) |
| `_NO_BORRAR_label_streaming.py` | Script de etiquetado. Lee el xlsx y asigna Estatus + Trámite a filas sin etiquetar. | NO modificar sin revisar con Claude |
| `_NO_BORRAR_build_report_mayo2026.py` | Script de construcción. Lee el xlsx etiquetado + template HTML y genera el reporte HTML final. | NO modificar sin revisar con Claude |
| `_NO_BORRAR_template.html` | Plantilla HTML base del reporte. Contiene toda la lógica JS y CSS del dashboard. | NO modificar directamente — usar Claude |
| `TENDENCIA.xlsx` | Histórico 2015–2025 de escrituras por abogado y mes. Referencia visual, datos ya hardcodeados en el script. | Solo consulta |
| `Reporte_Ejecutivo_Mayo2026.html` | Último reporte generado (output del script). Se abre en navegador. | Output, no editar |
| `Bitacora.xlsx.bak` | Respaldo automático que crea `label_streaming.py` antes de modificar. | Solo emergencias |

### Carpeta: `/Users/ke7/Documents/Claude/Projects/Bitácora GGR/`

| Archivo | Rol |
|---|---|
| `Bitacora GGR.html` | Workspace principal del reporte GGR. Claude edita aquí y copia a Downloads. |
| `Bitacora GGR_respaldo_13may2026.html` | Respaldo de versión anterior del reporte. |
| `TENDENCIA.xlsx` | Copia del histórico (para referencia). |

### Carpeta: `/Users/ke7/Downloads/`
| Archivo | Rol |
|---|---|
| `Bitacora GGR.html` | Versión que el usuario abre en el navegador (copia del workspace). |
| `BITACORA GGR*.xlsx` | Xlsx a etiquetar que el usuario descarga antes de cada ciclo. |

---

## DEPENDENCIAS PYTHON

```bash
pip install openpyxl --break-system-packages
```

Solo se requiere `openpyxl`. No se necesita pandas ni otras librerías.
Python 3.8+ recomendado.

---

## FLUJO PASO A PASO

### PASO 1 — Preparar el xlsx
1. Exportar/descargar el xlsx actualizado de Integranot.
2. Colocarlo en `/Users/ke7/Downloads/` con nombre `BITACORA GGR6.xlsx` (o el número que corresponda).
3. El archivo debe tener la hoja con las columnas estándar (ver sección COLUMNAS más abajo).

### PASO 2 — Etiquetar (label_streaming)
```bash
python3 "/Users/ke7/Documents/Claude/Projects/Bitácoras Notaría 31/_NO_BORRAR_label_streaming.py" "/Users/ke7/Downloads/BITACORA GGR6.xlsx"
```
- Crea respaldo `.bak` automáticamente.
- Solo etiqueta filas con Estatus **vacío** (no sobreescribe).
- Aplica reglas de Estatus y Trámite (traslativas, no traslativas, bancos, INFONAVIT, FOVISSSTE, etc.).
- Alias: COO → LBR / J31 → ignorar.

### PASO 3 — Generar reporte (build_report)
```bash
python3 "/Users/ke7/Documents/Claude/Projects/Bitácoras Notaría 31/_NO_BORRAR_build_report_mayo2026.py"
```
- Lee `Bitacora.xlsx` (o el path configurado en `SRC`).
- Lee `_NO_BORRAR_template.html` como plantilla base.
- Inyecta los JSON de datos: `DATA`, `DATA_ME`, `DATA_YR`, `DATA_AB_ME`, `DATA_CO`.
- Calcula `REF_DATE` = hoy - 1 día en CDT (UTC-5).
- Genera `Reporte_Ejecutivo_Mayo2026.html` (o el nombre configurado en `OUT`).

### PASO 4 — Abrir en navegador
Abrir directamente el HTML generado en Safari/Chrome. No requiere servidor.

---

## COLUMNAS CLAVE DEL XLSX

El script lee estas columnas por nombre exacto:

```
Escritura, Expediente, Estatus, Tramite, Abogado, Operación, Municipio,
Enajenante/Vendedor, Acreedor, Fech. Esc., Lleva TD,
TD. Digitalización Pago TD,
RPP. Entrada a RPP Testimonio,
RPP. Digitalización Entrada RPP,
RPP. Regreso de RPP Testimonio,
RPP. Captura Dato Inscripción RPP,
ESC. Entrega Firmas Completas,
CAT. Solicitud,
CAT. Dig Notificacion Catastral,
BANCOS. Dig. Acuse de Recibo Banco,
INFONAVIT. Dig. Acuse de Recibo INFONAVIT,
FOVISSSTE. Dig. Acuse de Recibo FOVISSSTE
```

---

## ESTATUS VÁLIDOS

```
CONCLUIDO | FALTA FINALIZAR | FALTA RPP | FALTA SIGER | EN RPP | FALTA TD | FALTA CIERRE
```

---

## ABOGADOS REGISTRADOS

```
ACM, AGB, ATM, EAC, GGR, GMB, JCL, JLGP, KST, LBR, LSR, MSR
Alias: COO → LBR   |   J31 → ignorar
```

---

## UMBRALES DE ALERTA (días naturales desde Fech. Esc.)

| Alerta | Umbral |
|---|---|
| TD Vencido | > 15 días hábiles sin Pago TD (traslativas) |
| TD Por Vencer | 9–14 días hábiles sin Pago TD |
| Folios >60 | > 60 días sin Entrega Firmas Completas |
| Entrada RPP >45 | > 45 días sin entrada a RPP |
| Inscripción RPP >90 | > 90 días traslativas / >30 días no traslativas sin retorno RPP |
| Bancos >150 | > 150 días con banco sin retorno RPP |
| Infonavit >150 | > 150 días con Infonavit sin retorno RPP |
| Fovissste >150 | > 150 días con Fovissste sin retorno RPP |

---

## NOTAS IMPORTANTES

- El reporte es un **HTML autocontenido** — no requiere internet ni servidor para funcionar (excepto la sección de comentarios/expedientes que usa GitHub API como backend).
- Los datos históricos de TENDENCIA (2015–2025) están **hardcodeados** en el script de build (`TEND_HIST` dict). No dependen de `TENDENCIA.xlsx` en runtime.
- El template `_NO_BORRAR_template.html` es la base para **todos los abogados**. El reporte GGR (`Bitacora GGR.html`) es una versión especializada con ajustes visuales adicionales que se edita directamente con Claude.
- Siempre que Claude edite `Bitacora GGR.html`, copia automáticamente a `/Users/ke7/Downloads/Bitacora GGR.html`. Recargar con **⌘+R** en el navegador para ver cambios.
- Para una nueva versión mensual: actualizar las rutas `SRC`, `TPL` y `OUT` al inicio del script de build.

---

## PALETA DE COLORES DEL REPORTE

```
--bg:       #0f172a   (fondo principal)
--surface:  #1e293b   (tarjetas)
--accent:   #6c8ef5   (azul acento)
--green:    #16a34a   (concluido / export CSV)
--yellow:   #d97706   (warning)
--red:      #dc2626   (urgente / export PDF)
--orange:   #ea580c
Fuente: Inter (Google Fonts, pesos 400–900)
```

### Colores por Estatus (tablas y gráficas)
```
CONCLUIDO:       #16a34a   verde
FALTA FINALIZAR: #f59e0b   ámbar
FALTA RPP:       #ca8a04   ámbar oscuro (tabla) / #92400e (gráfica)
FALTA SIGER:     #2dd4bf   teal
FALTA CIERRE:    #a78bfa   morado
EN RPP:          #38bdf8   azul cielo
FALTA TD:        #ef4444   rojo
```

---

## RESPALDOS

Antes de cualquier cambio mayor al template o script, crear respaldo con fecha:
```bash
cp "_NO_BORRAR_template.html" "_NO_BORRAR_template_BACKUP_DD-MM-YYYY.html"
```
