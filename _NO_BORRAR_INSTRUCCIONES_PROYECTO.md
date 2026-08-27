# INSTRUCCIONES DEL PROYECTO — BITÁCORA GGR / NOTARÍA 31
# (Copiar y pegar íntegramente en "Project Instructions" de Cowork)
# Última actualización: 2026-05-18 — PUNTO DE RESTAURACIÓN APROBADO

## Contexto general
Reporte ejecutivo mensual en formato HTML de una sola página (sin servidor, abre directo en navegador). Lee datos de Bitacora GGR.xlsx y genera el HTML mediante un script Python. Reporte exclusivo del abogado GGR.
Usuario: notariapublica31qro@gmail.com | Zona horaria: CDT (UTC-5)

## Archivos del proyecto
Ruta base: /Users/ke7/Documents/Claude/Projects/Bitácora GGR/

- index.html                            → pantalla de bienvenida GitHub Pages (usuario sube xlsx aquí)
- reporte.html                          → visor del reporte GitHub Pages (lee sessionStorage del index)
- _NO_BORRAR_template_ggr.html          → plantilla base HTML del reporte local (sincronizar al editar)
- _NO_BORRAR_build_report_mayo2026.py   → genera el reporte HTML local desde xlsx + template GGR
- _NO_BORRAR_label_streaming.py         → etiqueta Estatus/Tramite en xlsx (no modificar sin revisar)
- _NO_BORRAR_INSTRUCCIONES_PROYECTO.md  → este archivo (instrucciones para Cowork)
- _NO_BORRAR_PROCESO_COMPLETO.md        → guía maestra del proceso
- _upload_github.py                     → script Python de subida a GitHub (respaldo, ya no es el método principal)

Fuente de datos (se actualiza periódicamente): /Users/ke7/Downloads/Bitacora GGR.xlsx
Reporte local (generado por build script):     /Users/ke7/Downloads/Bitacora GGR.html
GitHub Pages (URL pública):                    https://notariapublica31qro.github.io/bitacora-ggr/

## REGLA CRÍTICA DE WORKFLOW

### Para editar reporte.html o index.html y subirlos a GitHub:
MÉTODO PREFERIDO — Chrome JS (no requiere Terminal):
Navegar a https://notariapublica31qro.github.io/bitacora-ggr/ y ejecutar en consola:
```js
(async()=>{
  const token=localStorage.getItem('gh_bitacora_token');
  const raw=await fetch('https://raw.githubusercontent.com/notariapublica31qro/bitacora-ggr/main/ARCHIVO.html?t='+Date.now());
  const html=await raw.text();
  const htmlNew=html.replace('CADENA_VIEJA','CADENA_NUEVA');
  const bytes=new TextEncoder().encode(htmlNew);
  let bin='';const C=8192;
  for(let i=0;i<bytes.length;i+=C)bin+=String.fromCharCode(...bytes.subarray(i,i+C));
  const b64=btoa(bin);
  const shaR=await fetch('https://api.github.com/repos/notariapublica31qro/bitacora-ggr/contents/ARCHIVO.html',{headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json'}});
  const{sha}=await shaR.json();
  const r=await fetch('https://api.github.com/repos/notariapublica31qro/bitacora-ggr/contents/ARCHIVO.html',{method:'PUT',headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json','Content-Type':'application/json'},body:JSON.stringify({message:'fix',content:b64,sha})});
  const j=await r.json();return j.content?'OK '+j.content.sha:'ERR '+JSON.stringify(j).slice(0,200);
})()
```
Token debe estar en localStorage('gh_bitacora_token'). Usar mcp__Claude_in_Chrome__javascript_tool.

### Para editar el reporte local (Bitacora GGR.html):
- Usar Edit (nunca Write completo) — ~3000 líneas
- Al terminar: cp workspace → Downloads Y cp Downloads → template
- NO modificar archivos con prefijo _NO_BORRAR_ sin confirmar primero con el usuario.
- NO tocar la carpeta Bitácoras Notaría 31 — ya no aplica a este proyecto GGR.
- Usuario recarga en navegador con ⌘+R (hard reload: ⌘+Shift+R).
- GitHub Pages tarda ~30 seg en actualizar — sugerir ⌘+Shift+R después de subir.

## Proceso completo paso a paso

### PASO 1 — Etiquetar el xlsx (cuando hay filas nuevas sin Estatus)
python3 "/Users/ke7/Documents/Claude/Projects/Bitácora GGR/_NO_BORRAR_label_streaming.py" "/Users/ke7/Downloads/Bitacora GGR.xlsx"
- Crea respaldo .bak automáticamente antes de modificar.
- Solo etiqueta filas con Estatus vacío (no sobreescribe).

### PASO 2 — Generar el reporte
python3 "/Users/ke7/Documents/Claude/Projects/Bitácora GGR/_NO_BORRAR_build_report_mayo2026.py"
- Lee /Users/ke7/Downloads/Bitacora GGR.xlsx (SRC al inicio del script).
- Usa _NO_BORRAR_template_ggr.html como plantilla base (TPL).
- Inyecta JSON de datos: DATA, DATA_ME, DATA_YR, DATA_AB_ME, DATA_CO.
- Calcula REF_DATE = hoy − 1 día en CDT (UTC-5).
- Genera /Users/ke7/Downloads/Bitacora GGR.html (OUT al inicio del script).
- Histórico GGR 2015-2025 hardcodeado en el script (dict GGR_HIST). TENDENCIA.xlsx ya no existe.

### PASO 3 — Abrir en navegador
Abrir /Users/ke7/Downloads/Bitacora GGR.html directamente. No requiere servidor.

## Columnas clave del xlsx (nombres exactos requeridos)
Escritura, Expediente, Estatus, Tramite, Abogado, Operación, Municipio,
Enajenante/Vendedor, Acreedor, Fech. Esc., Lleva TD,
TD. Digitalización Pago TD,
RPP. Entrada a RPP Testimonio, RPP. Digitalización Entrada RPP,
RPP. Regreso de RPP Testimonio, RPP. Captura Dato Inscripción RPP,
ESC. Entrega Firmas Completas,
CAT. Solicitud, CAT. Dig Notificacion Catastral,
BANCOS. Dig. Acuse de Recibo Banco,
INFONAVIT. Dig. Acuse de Recibo INFONAVIT,
FOVISSSTE. Dig. Acuse de Recibo FOVISSSTE

## Estatus válidos
CONCLUIDO | FALTA FINALIZAR | FALTA RPP | FALTA SIGER | EN RPP | FALTA TD | FALTA CIERRE

## Abogados registrados
ACM, AGB, ATM, EAC, GGR, GMB, JCL, JLGP, KST, LBR, LSR, MSR
Alias: COO → LBR  |  J31 → ignorar

## Umbrales de alerta (días naturales desde Fech. Esc.)
- TD Vencido:          > 15 días hábiles sin Pago TD (traslativas)
- TD Por Vencer:       9–14 días hábiles sin Pago TD
- Folios >60:          > 60 días sin Entrega Firmas Completas
- Entrada RPP >45:     > 45 días sin entrada a RPP (no EN RPP, sin acuse banco/info/fovi)
- Inscripción RPP >90: > 90 días traslativas / >30 días no traslativas sin retorno RPP
- Bancos >150:         > 150 días con banco sin retorno RPP
- Infonavit >150:      > 150 días con Infonavit sin retorno RPP
- Fovissste >150:      > 150 días con Fovissste sin retorno RPP

## Paleta de colores activa (mayo 2026) — MODO ÚNICO (sin toggle claro/oscuro)
Fuente: Inter (Google Fonts, pesos 400–900)
--bg: #3a7090 | --surface: #255570 | --card: #255570 | --border: #4e8aaa
--accent: #6c8ef5 | --text: #e2e8f0 | --muted: #94a3b8
--green: #16a34a | --yellow: #d97706 | --red: #dc2626 | --orange: #ea580c
NOTA: NO existe body.light — azul acero es el único modo para este usuario.

### Colores por Estatus (pie chart y tablas):
CONCLUIDO: #16a34a | FALTA FINALIZAR: #f59e0b | FALTA RPP: #a16207
FALTA SIGER: #2dd4bf | FALTA CIERRE: #a78bfa | EN RPP: #38bdf8 | FALTA TD: #ef4444

### SEG_COLORS (gráfica Tubería):
['#ef4444','#a78bfa','#2dd4bf','#a16207','#38bdf8','#f59e0b']

### Colores de botones de filtro:
Bancos (ambos): #db2777 (rosa)
Fovissste (ambos): #b45309 (café ámbar)
Infonavit (ambos): #059669 (verde esmeralda)
Foráneas: #2dd4bf (teal — mismo FALTA SIGER)
En RPP / Inscripción RPP: #0369a1 (azul marino)
Falta Cierre / Entrada RPP: #a78bfa (morado)
Falta RPP: #a16207 (café)
Traslativas / Falta TD: #ef4444 (rojo)
Finalizar / No Traslativas / Folios / Por Vencer: #f59e0b/#d97706 (ámbar)
Todas (.al-v-btn-all): accent por defecto; .on → #f5e49a bg + accent text

### Estado .on de botones:
background:#f5e49a; border-color:transparent; color: (color propio del botón)

## Layout del reporte (charts-row)
- grid-template-columns: 1fr 1fr (50-50)
- Columna izquierda: #box-pie (pie chart + leyenda; botones de abogado OCULTOS — siempre GGR)
- Columna derecha: .charts-right flex-col 2 filas 50-50:
  - Fila superior: tarjeta URGENTES (190px fijo, rojo #ef4444) + gráfica TENDENCIA
  - Fila inferior: tarjeta PENDIENTES (190px fijo, ámbar #f59e0b, SiTD/NoTD) + gráfica TUBERÍA
- syncKpi2() al final de updateKPIs() y updateMetaRing()
- updateMetaRing() usa GGR como abogado default (no GLOBAL)
- Sección INDICADORES (kpi-grid top 4 cards): OCULTA
- Sección pie-ab-btns (filtro abogados): OCULTA

## Secciones colapsables
Los 4 títulos de sección son colapsables con clic:
- #sec-estatus (📈 Estatus), #sec-pe (📝 Escrituras), #sec-ex (📂 Expedientes), #sec-al (🚨 URGENTES)
- Colapsado: título en #94a3b8, peso 400, chevron ▾ gira a ▸
- CSS: .sec-chev, .sec-collapsed, .sec-hidden{display:none!important;}
- JS: función toggleSection(secId) itera nextElementSibling hasta el próximo .sec

## Botones export CSV/PDF
- ESCRITURAS, EXPEDIENTES y URGENTES: botones removidos del header (funciones JS intactas)
- exportExCSV() y exportExPDF() exportan DESDE EX_DATA (no desde DOM)

## Persistencia (GitHub API como backend)
- Repo: notariapublica31qro/bitacora-ggr
- URL pública: https://notariapublica31qro.github.io/bitacora-ggr/
- Token PAT (sin expiración): <TOKEN_REDACTADO>
- Token se guarda en localStorage('gh_bitacora_token') — se ingresa una vez por navegador vía ⚙
- Archivos en repo: comments.json, expedientes.json
- NO hay integración con Notion en este reporte

## Tabla EXPEDIENTES
- Campos: xp(EXPEDIENTE), cl(CLIENTE), op(OPERACIÓN), td(TD), ac(BANCOS), mu(MUNICIPIO), es(ESTATUS), tr(TRÁMITE)
- EX_DATA = [] — se carga desde GitHub al abrir
- exportExCSV/PDF exportan desde EX_DATA (no DOM) — crítico para paginación
- Columna MUNICIPIO: font-size .72rem

## Gráfica TENDENCIA
- Año actual: #fcd34d | Años anteriores: paleta TENUE_BASE (tonos pastel)
- Y-axis máximo: Math.ceil(maxV/10)*10, cap histórico 120
- Sin líneas horizontales de cuadrícula
- Título izquierda, leyenda años derecha, font-size:.52rem

## Gráfica TUBERÍA
- Canvas: boxH-96 | pad.t:2 | Título margin:4px 0 0 0 | Leyendas margin-bottom:6px
- SEG_COLORS: ['#ef4444','#a78bfa','#2dd4bf','#a16207','#38bdf8','#f59e0b']
- Sin líneas separadoras entre segmentos (Math.round en coordenadas)

## Botones de tamaño uniforme
.tab y .al-v-btn: padding:4px 7px; border-radius:14px; font-size:.75rem; font-weight:700;
display:inline-flex; align-items:center; justify-content:center; gap:3px; box-sizing:border-box;
Función syncAlVBtnSizes() empareja tamaños en cada render.

## Scroll de secciones
function scrollToSection(id){ const el=document.getElementById(id); if(!el)return; const top=el.getBoundingClientRect().top+window.scrollY-18; window.scrollTo({top:Math.max(0,top),behavior:'smooth'}); }
function scrollToAl(){ scrollToSection('sec-al'); }
function scrollToPe(){ scrollToSection('sec-pe'); }

## Dependencias Python
pip install openpyxl --break-system-packages
Solo se requiere openpyxl. Python 3.8+.

## Rol y estilo de respuesta de Claude en este proyecto
- Responder siempre con fundamento legal claro conforme a legislación mexicana (especialmente Querétaro) cuando aplique.
- Para ediciones al HTML: usar siempre Edit (nunca Write completo).
- Al finalizar edición: cp workspace → Downloads Y cp Downloads → template.
- Confirmar cambios con una línea concisa. No explicar extensamente lo que se hizo.
- Antes de modificar cualquier archivo _NO_BORRAR_, confirmar con el usuario.
- NO tocar la carpeta Bitácoras Notaría 31.

## ✅ PUNTO DE RESTAURACIÓN — 2026-05-18 (estado aprobado por usuario)
GitHub repo: notariapublica31qro/bitacora-ggr
| Archivo       | SHA                                      | Tamaño   |
|--------------|------------------------------------------|----------|
| index.html   | 3d0661a6cf5ed156f0f575b44ecf7947e4347c48 | 725,218 B|
| reporte.html | ed5833edd08ffe226bb297c04e1ff44b173c739e | 838,849 B|

Estado visual validado:
- Banner superior: búho | texto | logo 31y69 (justify-self:end, borde derecho alineado)
- Banner inferior: logo | texto centrado | búho — una sola línea
- Sync GitHub: token en localStorage funcional
- SyntaxErrors JS corregidos (3 template literals en funciones export)

## Inventario de archivos — qué conservar y qué se puede borrar

### Workspace /Documents/Claude/Projects/Bitácora GGR/ — CONSERVAR TODO:
- _NO_BORRAR_template_ggr.html          ← base del reporte local, NO tocar
- _NO_BORRAR_build_report_mayo2026.py   ← script generador principal
- _NO_BORRAR_label_streaming.py         ← etiquetado de xlsx
- _NO_BORRAR_INSTRUCCIONES_PROYECTO.md  ← este archivo
- _NO_BORRAR_PROCESO_COMPLETO.md        ← guía maestra
- index.html                            ← pantalla bienvenida GitHub Pages
- reporte.html                          ← visor GitHub Pages
- _upload_github.py                     ← script subida (respaldo)

### /Users/ke7/Downloads/ — CONSERVAR:
- Bitacora GGR.xlsx                     ← datos activos
- Bitacora GGR.html                     ← reporte local activo
- Carpeta NUEVO/                         ← históricos xlsx 2019-2026

### Downloads — SE PUEDEN ELIMINAR (limpieza 2026-05-18):
- Bitacora.xlsx, Bitacora.xlsx.bak, Bitacora sin etiquetar.xlsx  ← duplicados/obsoletos
- index.html (en Downloads)             ← duplicado del workspace
- upload_index.html, upload_now.html    ← herramientas temporales obsoletas
- ICONO.af                              ← ícono fuente ya embebido en HTML
- ~$100284.docx, ~$p_9021.25_...doc    ← archivos temporales de Word
- Carpeta BORRADORES/                    ← versiones de trabajo anteriores
