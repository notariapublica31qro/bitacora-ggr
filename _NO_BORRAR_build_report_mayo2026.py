"""
build_report_mayo2026.py — genera Reporte_Ejecutivo_Mayo2026.html a partir de Bitacora.xlsx
Notaría 31 Querétaro — script principal del Reporte Ejecutivo mensual.

Uso directo:
    python3 _NO_BORRAR_build_report_mayo2026.py

Para una bitácora diferente (patrón exec-with-patches):
    src_code = open('<ruta_a_este_script>').read()
    src_code = src_code.replace('SRC = f"{_BASE}/Bitacora.xlsx"', 'SRC = "<ruta_origen>"')
    src_code = src_code.replace('OUT = f"{_BASE}/Reporte_Ejecutivo_Mayo2026.html"', 'OUT = "<ruta_salida>"')
    exec(compile(src_code, '<build>', 'exec'))
"""
import sys, os, json, re, time, shutil, glob as _glob, subprocess
from datetime import datetime, date, timedelta, timezone
from collections import defaultdict
import openpyxl

_vm_ggr = _glob.glob("/sessions/*/mnt/Bitácora GGR")
_BASE_GGR = _vm_ggr[0] if _vm_ggr else "/sessions/UNKNOWN/mnt/Bitácora GGR"
_vm_dl = _glob.glob("/sessions/*/mnt/Downloads")
_BASE_DL = _vm_dl[0] if _vm_dl else "/sessions/UNKNOWN/mnt/Downloads"

def _find_src_xlsx(base_dl):
    """Detecta el xlsx GGR automáticamente; orden de prioridad:
    1. Argumento CLI explícito
    2. Nombre legacy 'Bitacora GGR.xlsx'
    3. Cualquier *GGR*.xlsx en Downloads (más reciente)
    4. Cualquier .xlsx en Downloads (más reciente) como último recurso
    """
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        return sys.argv[1]
    legacy = os.path.join(base_dl, 'Bitacora GGR.xlsx')
    if os.path.isfile(legacy):
        return legacy
    ggr_xlsx = [f for f in _glob.glob(os.path.join(base_dl, '*GGR*.xlsx'))
                if not f.endswith('.bak')]
    if ggr_xlsx:
        return max(ggr_xlsx, key=os.path.getmtime)
    any_xlsx = [f for f in _glob.glob(os.path.join(base_dl, '*.xlsx'))
                if not f.endswith('.bak')]
    if any_xlsx:
        return max(any_xlsx, key=os.path.getmtime)
    raise FileNotFoundError(f"No se encontró ningún archivo .xlsx en {base_dl}")

SRC = _find_src_xlsx(_BASE_DL)
_vm_n31 = _glob.glob("/sessions/*/mnt/Bitácoras Notaría 31")
_BASE_N31 = _vm_n31[0] if _vm_n31 else "/sessions/UNKNOWN/mnt/Bitácoras Notaría 31"
TPL = f"{_BASE_GGR}/_NO_BORRAR_template_ggr.html"
OUT = f"{_BASE_DL}/Bitacora GGR.html"

_CDT = timezone(timedelta(hours=-5))
REF = (datetime.now(_CDT) - timedelta(days=1)).replace(tzinfo=None, hour=0, minute=0, second=0, microsecond=0)

T0 = time.time()
def t(): return f"[{time.time()-T0:.1f}s]"

INFONAVIT_KW = "INSTITUTO DEL FONDO NACIONAL DE LA VIVIENDA PARA LOS TRABAJADORES"
FOVISSSTE_KW = "FONDO DE LA VIVIENDA DEL INSTITUTO DE SEGURIDAD Y SERVICIOS SOCIALES DE LOS TRABAJADORES DEL ESTADO"
NO_BANCO = ['TIERRA Y ARMONIA']

# ─── TENDENCIA histórica 2015-2025 hardcoded (extraída de TENDENCIA.xlsx) ───
# Solo abogados activos. El año en curso (2026+) se suma desde Bitacora en runtime.
TEND_HIST = {
  'TOTAL': {
    '2015':[132,170,240,191,210,288,240,242,356,319,219,287],
    '2016':[190,224,204,237,234,219,229,304,374,354,278,302],
    '2017':[190,240,278,243,259,260,216,272,374,320,243,318],
    '2018':[170,263,285,250,217,221,229,260,380,294,238,368],
    '2019':[207,227,244,217,236,233,250,276,407,290,275,266],
    '2020':[220,239,270,144,211,268,312,264,518,644,341,480],
    '2021':[195,275,312,291,281,284,319,305,574,471,390,664],
    '2022':[276,282,313,256,339,301,312,380,491,552,355,446],
    '2023':[257,334,461,239,407,391,344,505,500,628,448,479],
    '2024':[302,367,346,374,329,376,374,361,484,478,390,458],
    '2025':[305,308,345,324,349,382,330,310,494,435,355,450],
  },
  'ACM': {
    '2015':[60,57,79,63,59,96,77,35,49,54,55,79],
    '2016':[54,45,59,55,65,55,59,109,57,37,71,61],
    '2017':[35,58,74,49,66,68,53,60,68,65,52,78],
    '2018':[38,91,65,73,43,53,53,61,54,50,53,102],
    '2019':[30,44,56,46,43,53,39,55,48,50,44,43],
    '2020':[28,31,56,22,49,40,83,52,84,144,50,100],
    '2021':[26,44,45,42,42,47,62,57,102,80,103,165],
  },
  'AGB': {
    '2015':[10,19,22,12,20,22,17,22,122,112,22,18],
    '2016':[12,25,17,15,20,26,24,16,127,138,40,25],
    '2017':[22,18,24,24,28,21,13,20,157,82,31,18],
    '2018':[8,14,20,15,10,10,11,14,143,51,31,22],
    '2019':[17,20,15,8,12,6,9,11,36,15,17,12],
  },
  'ATM': {
    '2019':[0,0,9,10,20,19,19,16,19,21,30,16],
    '2020':[20,23,16,9,13,19,27,23,28,49,41,32],
    '2021':[11,25,33,30,25,28,31,26,98,54,39,72],
    '2022':[36,41,43,33,51,42,50,45,47,85,51,42],
    '2023':[31,39,39,20,38,49,38,38,64,70,38,55],
    '2024':[26,40,40,41,46,38,40,41,58,53,49,32],
    '2025':[26,43,46,41,41,49,48,41,59,48,43,41],
  },
  'GGR': {
    '2015':[47,75,94,69,67,109,76,72,64,77,80,120],
    '2016':[26,32,23,36,26,23,24,32,37,28,20,54],
    '2017':[21,31,38,29,34,31,31,30,31,31,36,36],
    '2018':[25,37,6,4,5,3,3,5,3,3,8,3],
    '2019':[2,6,7,2,0,3,8,15,24,19,18,16],
    '2020':[15,17,20,9,14,23,31,27,69,86,35,59],
    '2021':[18,19,23,24,23,20,7,0,0,0,0,0],
    '2023':[0,0,7,8,25,20,23,35,34,62,19,33],
    '2024':[30,30,27,23,17,17,26,25,36,21,16,7],
    '2025':[19,4,26,13,25,18,12,14,32,13,25,22],
  },
  'GMB': {
    '2015':[0,0,1,0,0,0,0,0,0,0,0,0],
    '2016':[30,44,53,69,64,55,56,70,77,66,83,81],
    '2017':[41,58,72,69,58,56,52,66,47,69,65,97],
    '2018':[49,65,83,79,77,53,81,86,84,102,87,141],
    '2019':[63,73,73,64,67,73,96,74,87,74,70,77],
    '2020':[64,65,70,53,65,64,62,67,142,119,78,129],
    '2021':[56,95,103,84,79,80,95,96,161,110,119,276],
    '2022':[81,54,67,56,66,62,79,98,118,85,76,105],
    '2023':[53,54,76,43,85,51,48,43,96,119,112,137],
    '2024':[46,61,49,68,67,72,61,74,91,100,72,105],
    '2025':[49,46,57,64,64,67,50,51,99,93,77,97],
  },
  'JCL': {
    '2015':[0,2,28,40,38,39,31,68,84,34,38,33],
    '2016':[34,48,29,35,31,36,32,30,48,49,40,60],
    '2017':[27,38,35,34,38,48,35,46,41,32,42,51],
    '2018':[25,31,57,38,39,66,36,44,43,43,31,65],
    '2019':[35,34,48,43,47,34,45,51,81,47,57,52],
    '2020':[27,30,32,21,23,52,38,39,72,92,70,48],
    '2021':[32,42,47,45,44,49,45,47,73,103,50,42],
    '2022':[36,50,37,31,43,37,33,59,65,78,54,49],
    '2023':[40,49,73,46,75,72,47,59,69,100,48,45],
    '2024':[47,60,40,69,45,69,68,62,81,100,56,53],
    '2025':[57,52,37,57,43,69,61,53,55,72,46,95],
  },
  'JLGP': {
    '2015':[14,15,15,7,21,10,14,23,19,15,8,15],
    '2016':[16,2,1,0,2,3,2,15,9,5,14,2],
    '2017':[21,24,17,11,14,13,9,15,13,9,2,8],
    '2018':[8,6,15,7,16,4,8,24,19,13,6,12],
    '2019':[25,14,10,16,7,16,10,17,57,32,20,11],
    '2020':[21,16,24,4,9,18,18,9,12,6,8,15],
    '2021':[3,8,11,10,12,12,13,25,28,23,9,17],
    '2022':[3,6,10,4,14,6,10,11,20,13,8,13],
    '2023':[6,11,20,0,14,5,6,12,7,9,11,7],
    '2024':[1,1,1,1,2,1,7,3,3,3,4,2],
    '2025':[2,8,1,0,4,1,7,1,4,0,3,0],
  },
  'KST': {
    '2020':[25,28,26,9,19,23,32,28,70,87,35,60],
    '2021':[18,19,24,25,23,20,35,23,58,54,40,54],
    '2022':[27,40,41,44,50,38,48,60,69,83,50,44],
    '2023':[36,47,102,34,52,57,34,195,53,81,56,50],
    '2024':[31,43,48,38,46,43,46,37,68,54,57,55],
    '2025':[43,46,53,45,49,60,60,45,77,58,40,53],
  },
  'LSR': {
    '2015':[1,2,1,0,5,12,25,22,18,27,16,22],
    '2016':[18,28,22,27,26,21,32,32,19,31,10,19],
    '2017':[23,13,18,27,21,23,23,35,17,32,15,30],
    '2018':[17,19,39,34,27,32,37,26,34,32,22,23],
    '2019':[35,36,26,28,40,29,24,37,55,32,19,39],
    '2020':[20,29,26,17,19,29,21,19,41,61,24,37],
    '2021':[31,23,26,31,33,28,31,31,54,47,30,38],
    '2022':[36,31,25,25,36,35,29,36,56,69,36,37],
    '2023':[28,34,53,28,38,32,53,38,64,70,53,25],
    '2024':[40,40,40,32,32,30,34,28,40,33,30,102],
    '2025':[29,28,43,30,24,28,25,20,39,40,35,27],
  },
  'MSR': {
    '2022':[57,60,90,63,79,81,63,71,116,139,80,156],
    '2023':[63,100,91,60,80,105,95,85,113,117,111,127],
    '2024':[81,92,101,102,74,106,92,91,107,114,106,102],
    '2025':[80,81,82,74,99,90,67,85,129,111,86,115],
  },
}

LOCAL_MUNI = {'QUERÉTARO','QUERETARO','CORREGIDORA','EL MARQUÉS','EL MARQUES'}

def _es_foranea(muni):
    if not isinstance(muni, str) or not muni.strip() or muni.strip() == '-':
        return False
    parts = [p.strip().upper() for p in muni.split('-') if p.strip()]
    return not all(p in LOCAL_MUNI for p in parts)

CONSTR_KW = ['CONSTRUCTORA','INMOBILIARIA','GRUPO','RESIDENCIAL','DESARROLLOS',
             'PROYECTOS','URBANIZADORA','PROMOTORA','ARQUITECTOS']
BANCOS_KEYS = [("BBVA","BBVA"),("SCOTIABANK","SCOTIABANK"),("HSBC","HSBC"),("SANTANDER","SANTANDER"),
               ("BANREGIO","BANREGIO"),("BANBAJÍO","BANBAJÍO"),("BANBAJIO","BANBAJÍO"),
               ("AFIRME","AFIRME"),("MIFEL","MIFEL"),("INBURSA","INBURSA"),("BANORTE","BANORTE"),
               ("CITIBANAMEX","CITIBANAMEX"),("BANAMEX","BANAMEX"),("AZTECA","AZTECA"),
               ("INVEX","BANCO INVEX"),("AXIONEX","AXIONEX FINANCIERA"),
               ("BANCO NACIONAL DEL EJERCITO","BANJERCITO"),
               ("INSTITUTO DE SEGURIDAD SOCIAL PARA LAS FUERZAS ARMADAS","ISSFAM")]

def _parse_str_date(s):
    if not isinstance(s, str): return None
    s = s.strip()
    if not s or s == '-' or s == '*': return None
    for fmt in ('%d/%m/%Y','%Y-%m-%d','%d-%m-%Y'):
        try: return datetime.strptime(s, fmt)
        except: pass
    return None
def is_dt(v):
    if isinstance(v,(datetime,date)): return True
    return _parse_str_date(v) is not None
def to_dt(v):
    if isinstance(v, datetime): return v
    if isinstance(v, date): return datetime.combine(v, datetime.min.time())
    return _parse_str_date(v)
def is_dash(v): return isinstance(v, str) and v.strip() == "-"

def lleva_bancos(a):
    if not isinstance(a,str) or a.strip() in ('','-'): return False
    if any(nb in a for nb in NO_BANCO): return False
    p = {x.strip() for x in a.split(' - ')}
    return not p.issubset({INFONAVIT_KW, FOVISSSTE_KW})

def es_constructora(enaj):
    if not isinstance(enaj, str): return False
    e = enaj.upper()
    return any(kw in e for kw in CONSTR_KW)

def fmt_xp(xp):
    if not isinstance(xp,str) or '.' not in xp: return xp
    parts = xp.split('.')
    if len(parts)!=2: return xp
    if len(parts[1])==1: return f"{parts[0]}.{parts[1]}0"
    return xp

def abrev_acreedor(ac):
    if not ac or not isinstance(ac,str): return ""
    a = ac.strip()
    if a in ("","-"): return ""
    partes = [p.strip() for p in a.split(" - ")]
    abrev = []
    for p in partes:
        u = p.upper()
        if any(nb in p for nb in NO_BANCO): continue
        if INFONAVIT_KW in p: abrev.append("INFONAVIT")
        elif FOVISSSTE_KW in p: abrev.append("FOVISSSTE")
        else:
            matched = None
            for kw, short in BANCOS_KEYS:
                if kw in u: matched = short; break
            if matched: abrev.append(matched)
            else: abrev.append((p[:18]).strip() or p)
    seen = set(); out = []
    for x in abrev:
        if x not in seen: seen.add(x); out.append(x)
    if 'INFONAVIT' in out:
        rest = [x for x in out if x != 'INFONAVIT']; out = ['INFONAVIT'] + rest
    elif 'FOVISSSTE' in out:
        rest = [x for x in out if x != 'FOVISSSTE']; out = ['FOVISSSTE'] + rest
    REN = [('CAJA POPULAR FLORE','CP FLORENCIO ROSAS'),
           ('CAJA POPULAR MEXIC','CP MEXICANA'),
           ('BANCO INVEX SA IBM','BANCO INVEX'),
           ('INSTITUTO MEXICANO','IMSS')]
    out2 = []
    for x in out:
        for old, new in REN: x = x.replace(old, new)
        out2.append(x)
    return '-'.join(out2)

def dias_habiles(start, end):
    if not is_dt(start) or not is_dt(end): return 0
    if isinstance(start, datetime): start = start.date()
    if isinstance(end, datetime): end = end.date()
    days = 0
    d = start + timedelta(days=1)
    while d <= end:
        if d.weekday() < 5: days += 1
        d += timedelta(days=1)
    return days

# ─── PASO 0: Etiquetar filas nuevas en el xlsx antes de leer ───────────────
_label_script = f"{_BASE_GGR}/_NO_BORRAR_label_streaming.py"
if os.path.isfile(_label_script):
    print(f"{t()} Etiquetando {SRC} con label_streaming...", flush=True)
    _result = subprocess.run([sys.executable, _label_script, SRC],
                             capture_output=True, text=True)
    if _result.returncode == 0:
        print(f"{t()} Etiquetado OK.", flush=True)
    else:
        print(f"{t()} ADVERTENCIA — label_streaming terminó con error:\n{_result.stderr}", flush=True)
else:
    print(f"{t()} ADVERTENCIA — no se encontró label_streaming.py en {_BASE_GGR}", flush=True)

print(f"{t()} Leyendo {SRC}...", flush=True)
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
ws = wb.active
rows_iter = ws.iter_rows(values_only=True)
header = list(next(rows_iter))
H = {h:i for i,h in enumerate(header) if h is not None}

total = 0
es_count = defaultdict(int)
ab_stats = {}
tr_count = defaultdict(int)
me_global = defaultdict(int)
year_month_global = defaultdict(lambda: [0]*12)
year_month_ab = defaultdict(lambda: defaultdict(lambda: [0]*12))

seg_se = {'foraneas':0,'constructoras':0,'con_bancos':0,'con_infonavit':0,'con_fovissste':0}
data_pe = []
data_co = []

al_td = []; al_td_pv = []
al_rpp45 = []
al_firmas = []
al_rpp90solo = []
al_bancos90 = []; al_info90 = []; al_fovi90 = []

for row in rows_iter:
    if all(v is None for v in row): continue
    total += 1
    r_dict = dict(zip(header, row))

    es_str = (r_dict.get('Estatus') or '')
    es_str = es_str.strip() if isinstance(es_str, str) else ''

    VALID_ES = {'CONCLUIDO','FALTA FINALIZAR','FALTA RPP','FALTA SIGER','EN RPP','FALTA TD','FALTA CIERRE'}
    if es_str and es_str not in VALID_ES: continue

    if es_str: es_count[es_str] += 1

    ab = (row[H['Abogado']] or '')
    ab = ab.strip().upper() if isinstance(ab, str) else ''
    if ab == 'COO': ab = 'LBR'
    if ab == 'J31': ab = ''

    op = row[H['Operación']] or ''
    muni = row[H['Municipio']] or ''
    enaj_idx = H.get('Enajenante/Vendedor', H.get('Enajenante / Vendedor'))
    enaj = row[enaj_idx] if enaj_idx is not None else None
    ac_raw = row[H['Acreedor']] if 'Acreedor' in H else None
    fe = row[H['Fech. Esc.']]
    td_lleva = row[H['Lleva TD']] if 'Lleva TD' in H else None
    is_tras = isinstance(td_lleva, str) and td_lleva.strip().upper() == 'SI'
    tr_str = (r_dict.get('Tramite') or '')
    tr_str = tr_str.strip() if isinstance(tr_str, str) else ''
    e_num = row[H['Escritura']]
    e = str(e_num) if e_num is not None else ''
    xp = fmt_xp(str(row[H['Expediente']]).strip() if row[H['Expediente']] is not None else '')

    if tr_str: tr_count[tr_str] += 1

    if ab:
        if ab not in ab_stats:
            ab_stats[ab] = {'nombre':ab,'total':0,'concluido':0,'pendiente':0,
                            'falta_rpp':0,'en_rpp':0,'falta_td':0,'falta_cierre':0,
                            'falta_finalizar':0,'falta_siger':0}
        s = ab_stats[ab]
        s['total'] += 1
        if es_str in ('CONCLUIDO','FALTA FINALIZAR'):
            s['concluido'] += 1
            if es_str == 'FALTA FINALIZAR': s['falta_finalizar'] += 1
        elif es_str:
            s['pendiente'] += 1
            mp = {'FALTA RPP':'falta_rpp','EN RPP':'en_rpp','FALTA TD':'falta_td',
                  'FALTA CIERRE':'falta_cierre','FALTA SIGER':'falta_siger'}
            if es_str in mp: s[mp[es_str]] += 1

    if is_dt(fe):
        fe_d = to_dt(fe)
        ym = fe_d.strftime('%Y-%m')
        me_global[ym] += 1
        year_month_global[str(fe_d.year)][fe_d.month-1] += 1
        if ab: year_month_ab[ab][str(fe_d.year)][fe_d.month-1] += 1

    is_foranea = _es_foranea(muni)
    is_co = es_constructora(enaj)
    is_ba_seg   = lleva_bancos(ac_raw)
    is_info_seg = isinstance(ac_raw,str) and INFONAVIT_KW in ac_raw
    is_fovi_seg = isinstance(ac_raw,str) and FOVISSSTE_KW in ac_raw
    if is_foranea:  seg_se['foraneas'] += 1
    if is_co:       seg_se['constructoras'] += 1
    if is_ba_seg:   seg_se['con_bancos'] += 1
    if is_info_seg: seg_se['con_infonavit'] += 1
    if is_fovi_seg: seg_se['con_fovissste'] += 1

    def _acuse_pend(v):
        if v is None: return True
        if isinstance(v, (datetime, date)): return False
        s = str(v).strip()
        return s in ('', '*', '-', 'nan', 'NaN', 'None')
    _ba_acuse   = row[H['BANCOS. Dig. Acuse de Recibo Banco']]       if 'BANCOS. Dig. Acuse de Recibo Banco'       in H else None
    _info_acuse = row[H['INFONAVIT. Dig. Acuse de Recibo INFONAVIT']] if 'INFONAVIT. Dig. Acuse de Recibo INFONAVIT' in H else None
    _fovi_acuse = row[H['FOVISSSTE. Dig. Acuse de Recibo FOVISSSTE']] if 'FOVISSSTE. Dig. Acuse de Recibo FOVISSSTE' in H else None
    is_ba   = is_ba_seg   and _acuse_pend(_ba_acuse)
    is_info = is_info_seg and _acuse_pend(_info_acuse)
    is_fovi = is_fovi_seg and _acuse_pend(_fovi_acuse) and (not is_ba_seg or _acuse_pend(_ba_acuse))

    fe_str = fe_d.strftime('%Y-%m-%d') if is_dt(fe) else ''
    ac_short = abrev_acreedor(ac_raw)

    _cat_sol = row[H.get('CAT. Solicitud')] if 'CAT. Solicitud' in H else None
    _cat_dig = row[H.get('CAT. Dig Notificacion Catastral')] if 'CAT. Dig Notificacion Catastral' in H else None
    _activo  = lambda v: v is not None and str(v).strip() not in ('', '-', '*', 'nan', 'NaN', 'None')
    cat_pend = bool(
        is_tras and
        _activo(_cat_sol) and
        not _activo(_cat_dig) and str(_cat_dig or '').strip() != '-' and
        es_str != 'FALTA TD'
    )

    item_base = {
        'e':e,'ab':ab,'op':op if isinstance(op,str) else '',
        'mu':(muni[:25] if isinstance(muni,str) else ''),
        'fe':fe_str,'es':es_str,'tr':tr_str,
        'td':is_tras,'ba':is_ba,'co':is_co,'fo':is_foranea,
        'infonavit':is_info,'fovissste':is_fovi,
        'rpp':False,'ac':ac_short,'xp':xp,'cat_pend':cat_pend
    }

    if es_str in ('CONCLUIDO','FALTA FINALIZAR'):
        data_co.append(item_base)
    elif es_str:
        data_pe.append(item_base)

    if not es_str or es_str == 'CONCLUIDO': continue
    if not is_dt(fe): continue
    fe_d = to_dt(fe)
    d_nat = (REF - fe_d).days
    if d_nat <= 0: continue

    item_al = {'e':e,'ab':ab,'op':item_base['op'],'mu':item_base['mu'],
               'fe':fe_str,'d':d_nat,'v':True,'xp':xp,'ac':ac_short}

    td_pago = row[H.get('TD. Digitalización Pago TD')] if 'TD. Digitalización Pago TD' in H else None
    cat_dig = row[H.get('CAT. Dig Notificacion Catastral')] if 'CAT. Dig Notificacion Catastral' in H else None
    op_str_upper = op.upper() if isinstance(op, str) else ''
    excepcion_aplic = is_tras and 'APLICACION DE BIENES' in op_str_upper and is_dt(cat_dig) and not is_dt(td_pago)
    if is_tras and not is_dt(td_pago) and not excepcion_aplic:
        dh = dias_habiles(fe_d, REF)
        if dh > 15:
            al_td.append({**item_al,'li':15,'v':True})
        elif dh >= 9 and dh < 15:
            al_td_pv.append({**item_al,'li':15,'v':False,'d':d_nat})

    rpp_ent = row[H.get('RPP. Entrada a RPP Testimonio')]
    rpp_dig_ent = row[H.get('RPP. Digitalización Entrada RPP')]
    rpp_entered = is_dt(rpp_ent) or is_dt(rpp_dig_ent)
    if es_str != 'EN RPP' and not rpp_entered and d_nat > 45:
        al_rpp45.append({**item_al,'li':45})

    firmas_dt = row[H.get('ESC. Entrega Firmas Completas')]
    if not is_dt(firmas_dt) and d_nat > 60:
        al_firmas.append({**item_al,'li':60})

    rpp_reg = row[H.get('RPP. Regreso de RPP Testimonio')]
    rpp_cap = row[H.get('RPP. Captura Dato Inscripción RPP')]
    rpp_term = is_dt(rpp_reg) or is_dt(rpp_cap)
    if not rpp_term and not is_ba and not is_info and not is_fovi:
        if is_tras and d_nat > 90:
            al_rpp90solo.append({**item_al,'li':90})
        elif not is_tras and d_nat > 30:
            al_rpp90solo.append({**item_al,'li':30})

    # Bancos/Infonavit/Fovissste: alertar cuando el acuse aún está pendiente (is_ba/is_info/is_fovi
    # ya incorporan el chequeo de acuse pendiente), independientemente de si RPP regresó o no.
    if is_ba and d_nat > 150:
        al_bancos90.append({**item_al,'li':150})
    if is_info and d_nat > 150:
        al_info90.append({**item_al,'li':150})
    if is_fovi and d_nat > 150:
        al_fovi90.append({**item_al,'li':150})

wb.close()
print(f"{t()} Filas procesadas: {total}")

traslativas = sum(1 for r in (data_pe + data_co) if r['td'])
concluido = es_count.get('CONCLUIDO',0) + es_count.get('FALTA FINALIZAR',0)
pendientes = total - concluido
pct_concluido = round(concluido/total*100, 1) if total else 0

data = {
    't': {'total': total,'traslativas': traslativas,'no_traslativas': total - traslativas,
          'concluido': concluido,'pendientes': pendientes,'pct_concluido': pct_concluido},
    'es': dict(es_count),
    'ab': ab_stats,
    'se': seg_se,
    'me': dict(me_global),
    'tr': dict(tr_count),
    'al': {'td': al_td,'rpp45': al_rpp45,'firmas': al_firmas,'rpp90': [],'pe': [],
           'bancos': [],'infonavit': [],'fovissste': [],'td_pv': al_td_pv,
           'rpp90solo': al_rpp90solo,'bancos90': al_bancos90,
           'infonavit90': al_info90,'fovissste90': al_fovi90},
    'pe': data_pe
}

# ── Tendencia: histórico GGR hardcodeado (2015-2025) + año en curso desde bitácora ──
GGR_HIST = {
    2015: [47,75,94,69,67,109,76,72,64,77,80,120],
    2016: [26,32,23,36,26,23,24,32,37,28,20,54],
    2017: [21,31,38,29,34,31,31,30,31,31,36,36],
    2018: [25,37,6,4,5,3,3,5,3,3,8,3],
    2019: [2,6,7,2,0,3,8,15,24,19,18,16],
    2020: [15,17,20,9,14,23,31,27,69,86,35,59],
    2021: [18,19,23,24,23,20,7,0,0,0,0,0],
    2023: [0,0,7,8,25,20,23,35,34,62,19,33],
    2024: [30,30,27,23,17,17,26,25,36,21,16,7],
    2025: [19,4,26,13,25,18,12,14,32,13,25,22],
}
data_me = {str(y): list(arr) for y, arr in sorted(GGR_HIST.items())}
for y, arr in sorted(year_month_ab.get('GGR', {}).items()):
    if str(y) not in data_me:          # solo años no cubiertos por GGR_HIST (ej. 2026+)
        data_me[str(y)] = list(arr)
data_yr = {y: sum(arr) for y, arr in data_me.items()}
data_ab_me = {}
for ab, yr_dict in year_month_ab.items():
    ab_n = 'LBR' if ab == 'COO' else ab
    if ab_n not in data_ab_me:
        data_ab_me[ab_n] = {}
    for y, arr in yr_dict.items():
        data_ab_me[ab_n][y] = list(arr)

print(f"{t()} DATA construido. Total: {total}, Traslativas: {traslativas}, Concluidas: {concluido}, Pendientes: {pendientes}")
print(f"  Alertas — td:{len(al_td)}, td_pv:{len(al_td_pv)}, rpp45:{len(al_rpp45)}, firmas:{len(al_firmas)}, rpp90solo:{len(al_rpp90solo)}")

print(f"{t()} Cargando template {TPL}...")
with open(TPL, encoding='utf-8') as f: html = f.read()

def replace_const(html, name, json_str):
    # Coincide con la línea completa sin importar el valor actual (localStorage, {}, [], etc.)
    pat = re.compile(rf'^const {re.escape(name)} = .+;\s*$', re.MULTILINE)
    new_line = f'const {name} = {json_str};'
    result = pat.sub(new_line, html, count=1)
    if result == html:
        raise ValueError(f"replace_const: no se encontró 'const {name} = ...;' en el template")
    return result

html = replace_const(html, 'DATA', json.dumps(data, ensure_ascii=False))
html = replace_const(html, 'DATA_ME', json.dumps(data_me, ensure_ascii=False))
html = replace_const(html, 'DATA_YR', json.dumps(data_yr, ensure_ascii=False))
html = replace_const(html, 'DATA_AB_ME', json.dumps(data_ab_me, ensure_ascii=False))
html = replace_const(html, 'DATA_CO', json.dumps(data_co, ensure_ascii=False))

mes_idx_to_name = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
ref_mes = mes_idx_to_name[REF.month - 1]
ref_dia = f'{REF.day:02d}'
ref_anio = str(REF.year)
html = re.sub(r'<span id="cal-mes">[^<]+</span>', f'<span id="cal-mes">{ref_mes}</span>', html)
html = re.sub(r'<span id="cal-dia">[^<]+</span>', f'<span id="cal-dia">{ref_dia}</span>', html)
html = re.sub(r'<span id="cal-anio">[^<]+</span>', f'<span id="cal-anio">{ref_anio}</span>', html)
html = re.sub(r'(<div class="top"><span>)[^<]+(</span></div>)', rf'\g<1>{ref_mes}\2', html, count=1)
html = re.sub(r'(<div class="day">)\s*\d+\s*(</div>)', rf'\g<1>{ref_dia}\2', html, count=1)
html = re.sub(r'(<div class="yr"><span>)[^<]+(</span></div>)', rf'\g<1>{ref_anio}\2', html, count=1)

_T = total or 1
for _id, _val in [
    ('seg-cons-sv', str(seg_se['constructoras'])),
    ('seg-cons-sp', f"{seg_se['constructoras']/_T*100:.1f}%"),
    ('seg-banco-sv', str(seg_se['con_bancos'])),
    ('seg-banco-sp', f"{seg_se['con_bancos']/_T*100:.1f}%"),
    ('seg-info-sv',  str(seg_se['con_infonavit'])),
    ('seg-info-sp',  f"{seg_se['con_infonavit']/_T*100:.1f}%"),
    ('seg-fovi-sv',  str(seg_se['con_fovissste'])),
    ('seg-fovi-sp',  f"{seg_se['con_fovissste']/_T*100:.1f}%"),
    ('seg-fora-sv',  str(seg_se['foraneas'])),
    ('seg-fora-sp',  f"{seg_se['foraneas']/_T*100:.1f}%"),
]:
    html = re.sub(rf'id="{_id}">[^<]+<', f'id="{_id}">{_val}<', html)

me_keys = sorted(me_global.keys())
if me_keys:
    minK = me_keys[0]; maxK = me_keys[-1]
    minY, minM = minK.split('-'); maxY, maxM = maxK.split('-')
    MES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    if minK == maxK:
        periodo = f'{MES[int(minM)-1]} {minY}'
    else:
        periodo = f'{MES[int(minM)-1]} {minY} - {MES[int(maxM)-1]} {maxY}'
    html = re.sub(r'<span class="lbl">Periodo:</span>\s*[^<]+</span>',
                  f'<span class="lbl">Periodo:</span> {periodo}</span>', html, count=1)
    html = re.sub(r'(<span style="color:#ffffff;">)[^<]+(</span></span>)',
                  rf'\g<1>{periodo}\2', html, count=1)

_ref_str = REF.strftime('%d-%m-%Y')
html = re.sub(r'<title>[^<]+</title>', f'<title>Bitácora GGR a {_ref_str}</title>', html, count=1)
html = re.sub(r'<span id="periodo-label" style="display:none">[^<]+</span>',
              '<span id="periodo-label" style="display:none">Mayo 2026</span>', html, count=1)

with open(OUT, 'w', encoding='utf-8') as f: f.write(html)
print(f"{t()} Reporte guardado: {OUT}")
