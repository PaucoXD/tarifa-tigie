"""Actualiza index.html con los archivos vigentes del SNICE y el Anexo 2.4.1 (NOM) del DOF.

Uso:  python tools/actualizar.py            (descarga lo nuevo, regenera la página y escribe el reporte)
      python tools/actualizar.py --offline  (usa solo lo que ya está en tools/cache)

Fuentes:
  - Fracciones, aranceles y NICO: Excel «FRACCIONESARANCELARIAS-LIGIE_<fecha>.xlsx» del SNICE.
  - Partidas, subpartidas y regulaciones: «BASEUNICA-LIGIE_<fecha>.xlsb» del SNICE.
  - NOM: Anexo 2.4.1 del Acuerdo de reglas y criterios de la SE (Word del DOF), base + modificaciones de nom_mods.json.
Sale con código 2 si algo no cuadra (la página no se toca).
"""
import json, re, sys, os, html, datetime, urllib.request
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')
TOOLS = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(TOOLS)
CACHE = os.path.join(TOOLS, 'cache')
PAGE = os.path.join(ROOT, 'index.html')
REPORT = os.path.join(TOOLS, 'ultimo_reporte.txt')
OFFLINE = '--offline' in sys.argv
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
SNICE = 'https://www.snice.gob.mx'
MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
rep_lines = []


def log(s=''):
    print(s)
    rep_lines.append(s)


def fail(msg):
    log('ERROR: ' + msg)
    log('La página NO se modificó.')
    open(REPORT, 'w', encoding='utf-8').write('\n'.join(rep_lines) + '\n')
    sys.exit(2)


def get(url, timeout=180):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def cached(name, url):
    path = os.path.join(CACHE, name)
    if not os.path.exists(path):
        if OFFLINE:
            fail('falta en caché: ' + name)
        print('  descargando', name, '...')
        data = get(url)
        open(path, 'wb').write(data)
    return path


def latest_snice(page, prefix, ext):
    """Busca en una página del SNICE el archivo <prefix>_<AAAAMMDD>...<ext> más reciente."""
    if OFFLINE:
        files = sorted(f for f in os.listdir(CACHE) if f.startswith(prefix))
        if not files:
            fail('no hay ' + prefix + ' en caché')
        return files[-1], None, files[-1][len(prefix) + 1:len(prefix) + 9]
    s = get(SNICE + page).decode('utf-8', 'ignore')
    found = re.findall(r'(?:https://www\.snice\.gob\.mx)?/~oracle/SNICE_DOCS/(' + re.escape(prefix) + r'_(\d{8})-\d{8}\.' + ext + ')', s)
    if not found:
        fail('no encontré ' + prefix + ' en ' + SNICE + page + ' (¿cambió la página del SNICE?)')
    name, fecha = max(found, key=lambda x: x[1])
    return name, SNICE + '/~oracle/SNICE_DOCS/' + name, fecha


def fecha_larga(aaaammdd):
    d = datetime.date(int(aaaammdd[:4]), int(aaaammdd[4:6]), int(aaaammdd[6:]))
    return f'{d.day} de {MESES[d.month - 1]} de {d.year}'


def fecha_corta(aaaammdd):
    return f'{aaaammdd[6:]}/{aaaammdd[4:6]}/{aaaammdd[:4]}'


def dmy(s):
    d, m, y = s.split('/')
    return (int(y), int(m), int(d))


def norm(x):
    if x is None:
        return ''
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return re.sub(r'\s+', ' ', str(x)).strip()


FRP = re.compile(r'^\d{4}\.\d{2}\.\d{2}$')

# ------------------------------------------------------------------ página actual
src = open(PAGE, encoding='utf-8').read()


def block(id_):
    m = re.search(r'<script id="' + id_ + r'" type="application/json">(.*?)</script>', src, re.S)
    if not m:
        fail('no encontré el bloque de datos "' + id_ + '" en index.html')
    return m, json.loads(m.group(1).replace('<\\/', '</'))


m_data, OLD = block('data')
m_nom, OLDNOM = block('nom')
m_rg, OLDRG = block('rg')
os.makedirs(CACHE, exist_ok=True)

# ------------------------------------------------------------------ 1. fracciones, aranceles, NICO
log('1) Fracciones, aranceles y NICO (SNICE)')
fa_name, fa_url, fa_fecha = latest_snice('/cs/avi/snice/ligie.info22.html', 'FRACCIONESARANCELARIAS-LIGIE', 'xlsx')
log(f'   archivo: {fa_name}')
import openpyxl  # noqa: E402
wb = openpyxl.load_workbook(cached(fa_name, fa_url), read_only=True)


def sheet_rows(ws, first_header):
    """Devuelve (índice de columnas por encabezado, filas de datos) a partir de la fila de encabezado."""
    rows = list(ws.iter_rows(values_only=True))
    for i, r in enumerate(rows):
        cells = [norm(c).lower() for c in r]
        if any(c.startswith(first_header) for c in cells):
            return cells, rows[i + 1:]
    fail('no encontré el encabezado "' + first_header + '" en la hoja ' + ws.title)


hdr, rows = sheet_rows(wb['FA'], 'fracci')
c_f = next(i for i, c in enumerate(hdr) if c.startswith('fracci'))
FA = {}
for r in rows:
    f = norm(r[c_f])
    if FRP.match(f):
        FA[f] = [norm(r[c_f + 1]), norm(r[c_f + 2]), norm(r[c_f + 3]), norm(r[c_f + 4])]
hdr, rows = sheet_rows(wb['NICO'], 'fracci')
c_f = next(i for i, c in enumerate(hdr) if c.startswith('fracci'))
NICO = {}
for r in rows:
    f = norm(r[c_f])
    if FRP.match(f):
        NICO.setdefault(f, []).append([norm(r[c_f + 1]).zfill(2), norm(r[c_f + 2])])
log(f'   {len(FA):,} fracciones · {sum(len(v) for v in NICO.values()):,} NICO')
if len(FA) < 7000:
    fail(f'solo {len(FA)} fracciones en el archivo del SNICE: parece incompleto')

# ------------------------------------------------------------------ 2. Base Única (jerarquía y regulaciones)
log('2) Base Única (SNICE)')
bu_name, bu_url, bu_fecha = latest_snice('/cs/avi/snice/home.html', 'BASEUNICA-LIGIE', 'xlsb')
log(f'   archivo: {bu_name}')
from pyxlsb import open_workbook  # noqa: E402
BU = []
with open_workbook(cached(bu_name, bu_url)) as wbb:
    with wbb.get_sheet(wbb.sheets[0]) as ws:
        for r in ws.rows():
            BU.append([c.v for c in r])
hi = next((i for i, r in enumerate(BU) if any(norm(c).lower() == 'fracción arancelaria' for c in r)), None)
if hi is None:
    fail('no encontré el encabezado de la Base Única')
H = [norm(c).lower() for c in BU[hi]]
W = len(H)


def col(prefix, start=0):
    for i in range(start, W):
        if H[i].startswith(prefix.lower()):
            return i
    fail('la Base Única ya no tiene la columna "' + prefix + '"')


# Ojo: «FA CORRELATIVA» y «NICO CORRELATIVA» son los códigos anteriores (correlación), no los vigentes;
# el código vigente es la columna «Fracción arancelaria» y las filas de NICO solo traen el NICO.
C = {k: col(p) for k, p in {'code': 'fracción arancelaria', 'nico': 'nico', 'desc': 'descripción'}.items()}
FLAGS = {'perm': 'permisos/ avisos previos', 'azu': 'acuerdo azúcar', 'cexp': 'acuerdo de control de exportaciones', 'nom': "acuerdo de nom's",
         'cul': 'acuerdo de cultura', 'cic': 'acuerdo de cicoplafest', 'emb': 'acuerdo de embargos', 'sener': 'acuerdo de sener', 'sader': 'acuerdo sader',
         'salud': 'acuerdo salud', 'sedena': 'acuerdo de sedena', 'semarnat': 'acuerdo de semarnat', 'cc': 'cuotas compensatorias', 'imxp': 'immex prohibidas',
         'imxs': 'immex sensibles', 'apecic': 'apecic'}
FLAGC = {k: col(p) for k, p in FLAGS.items()}
TXTC = {'fr1': col('decreto de frontera fracción i '), 'fr2': col('decreto de frontera fracción ii'), 'chet': col('decreto de chetumal'), 'pe': col('precios estimados')}
P5, P4 = col('prosec art. 5'), col('prosec art. 4')
P5C = range(TXTC['chet'] + 1, P5)
P4C = range(P5 + 1, P4)
PT, SUB5, SUB6 = {}, {}, {}
RGraw, BU_NOM = {}, set()
curf = None
for r in BU[hi + 1:]:
    r = r + [None] * (W - len(r))
    code, desc, ni = norm(r[C['code']]), norm(r[C['desc']]), norm(r[C['nico']])
    if re.fullmatch(r'\d{2}\.\d{2}', code):
        PT[code.replace('.', '')] = desc
    elif re.fullmatch(r'\d{4}\.\d', code):
        SUB5[code] = desc
    elif re.fullmatch(r'\d{4}\.\d{2}', code):
        SUB6[code] = desc
    if FRP.match(code):
        curf = f = code
        ni = ni.zfill(2) if re.fullmatch(r'\d{1,2}', ni) else ''
    elif not code and curf and re.fullmatch(r'\d{1,2}', ni):
        f, ni = curf, ni.zfill(2)
    else:
        if code:
            curf = None
        continue
    g = RGraw.setdefault(f, {})

    def put(k, val):
        if ni == '':
            g[k] = {'*': val}
        else:
            cur = g.setdefault(k, {})
            if '*' not in cur:
                cur[ni] = val
    for k, c in FLAGC.items():
        x = norm(r[c])
        if x and x != '*':
            if k == 'nom':
                BU_NOM.add(f)
            else:
                put(k, '' if x == 'Sí' else x)
    for k, c in TXTC.items():
        x = norm(r[c])
        if x and x != '*':
            put(k, x)
    p5 = [[norm(BU[hi][c]).replace('SECTOR ', ''), norm(r[c])] for c in P5C if norm(r[c]) and norm(r[c]) != '*']
    if p5:
        put('p5', p5)
    p4 = [[norm(BU[hi][c]).replace('SECTOR ', ''), re.sub(r'^Sí\s*', '', norm(r[c]))] for c in P4C if norm(r[c]) and norm(r[c]) != '*']
    if p4:
        put('p4', p4)
RG = {}
for f, g in RGraw.items():
    o = {}
    allni = [n[0] for n in NICO.get(f, [])]
    for k, m in g.items():
        if '*' in m:
            o[k] = [m['*']] if m['*'] != '' else 1
        else:
            vals = list(m.values())
            if len(set(json.dumps(x, ensure_ascii=False) for x in vals)) == 1 and set(m) == set(allni):
                o[k] = [vals[0]] if vals[0] != '' else 1
            else:
                o[k] = {n: (x if x != '' else 1) for n, x in m.items()}
    if o:
        RG[f] = o
log(f'   {len(PT):,} partidas · {len(SUB6):,} subpartidas · {len(RG):,} fracciones con regulaciones o programas')
if len(PT) < 1000:
    fail('la Base Única parece incompleta')

# ------------------------------------------------------------------ 3. armar la lista de fracciones
FR = []
for f in sorted(FA):
    d, u, imp, exp = FA[f]
    dash = SUB5.get(f[:6], '')
    sub = SUB6.get(f[:7], '')
    FR.append([f, d, u, imp, exp, dash, sub, NICO.get(f, [])])
CAP = OLD['cap']
caps = sorted({f[:2] for f in FA})
falt = [c for c in caps if c not in CAP]
if falt:
    log('   AVISO: capítulos sin nombre en la página: ' + ', '.join(falt))

# ------------------------------------------------------------------ 4. NOM (Anexo 2.4.1)
log('3) NOM (Anexo 2.4.1 del DOF)')
CFG = json.load(open(os.path.join(TOOLS, 'nom_mods.json'), encoding='utf-8'))
NUM = re.compile(r'^\s*"?\s*([1-5])\.-')
SEC = re.compile(r'^\s*([IVXL]+)\.(?:\s|$)(?!\s*(?:a|y)\s+[IVXL]+\.)')
NOMRE = re.compile(r'NOM-\d{3}(?:[/-]\d)?-[A-Z0-9]+(?:/[A-Z0-9]+)?-\d{4}')
NI2 = re.compile(r'^\d{2}$')


def clean(s):
    return re.sub(r'NOM- +', 'NOM-', re.sub(r'[ \t\xa0]+', ' ', s.replace('\x0b', '\n').replace('\r', '\n'))).strip()


def parse_anexo(t, amend=False):
    """Lee las tablas del Word del DOF (celdas separadas por \\x07). Devuelve entradas por fracción."""
    out, cur, num, sec, sh = [], None, None, None, ''
    toks = t.split('\x07')
    i = 0

    def head(p):
        nonlocal num, sec, sh, cur
        m = NUM.match(p)
        if m:
            num, sec, sh, cur = m.group(1), None, '', None
            return
        if re.match(r'^\s*[IVXL]+\.\s+(a|y)\s+[IVXL]+\.', p):
            sec, cur = None, None
            return
        m = SEC.match(p)
        if m and num == '3' and (amend or 'NOM-' in p):
            sec, sh, cur = m.group(1), (clean(p) if 'NOM-' in p else ''), None
    while i < len(toks):
        paras = toks[i].split('\r')
        for p in paras[:-1]:
            head(p)
        if len(paras) > 1:
            head(paras[-1])
        c0 = clean(paras[-1])
        if (FRP.match(c0) or (NI2.match(c0) and cur is not None)) and num:
            d = clean(toks[i + 1]) if i + 1 < len(toks) else ''
            a = clean(toks[i + 2]) if i + 2 < len(toks) else ''
            if FRP.match(a) or NI2.match(a):
                a, i = '', i + 2
            else:
                i += 3
            if FRP.match(c0):
                cur = {'n': num, 's': sec, 'sh': sh, 'f': c0, 'd': d, 'a': a, 'ni': []}
                out.append(cur)
            else:
                cur['ni'].append([c0, d, a])
            continue
        i += 1
    return out


def blocks(a):
    res = []
    for line in [x.strip() for x in a.split('\n') if x.strip()]:
        m = NOMRE.match(line)
        if m:
            res.append([m.group(0), line[m.end():].strip(' .,;')])
        elif res:
            res[-1][1] = (res[-1][1] + ' ' + line).strip()
        else:
            res.append(['', line])
    return res


def anexo_text(codigo):
    b = open(cached(f'dof_{codigo}.doc', f'https://sidof.segob.gob.mx/notas/getDoc/{codigo}'), 'rb').read()
    i = b.find(b'ANEXO 2.4.1')
    if i < 0:
        fail(f'el documento DOF {codigo} no contiene el ANEXO 2.4.1 (revisa nom_mods.json)')
    ends = [x for x in (b.find(b'TRANSITORIO', i), b.find(b'ANEXO 2.5', i), b.find(b'Se suprimen', i)) if x > 0]
    return b[i:min(ends) if ends else len(b)].decode('cp1252')


base = parse_anexo(anexo_text(CFG['base']['codigo']))
secnom, sechead = {}, {}
for e in base:
    if e['n'] == '3' and e['sh']:
        sechead[e['s']] = e['sh']
        mm = NOMRE.search(e['sh'])
        secnom[e['s']] = mm.group(0) if mm else ''
for s, v in CFG.get('encabezados', {}).items():
    secnom[s], sechead[s] = v['nom'], v['texto']
db = {}


def entry_blocks(e, dof):
    ni = [x[0] for x in e['ni']]
    out = []
    if e['n'] == '3':
        out.append({'nom': secnom.get(e['s'], ''), 'ac': e['a'] if e['a'] != '...' else '', 'ni': ni, 'dof': dof})
    else:
        for nom, ac in blocks(e['a']):
            out.append({'nom': nom, 'ac': ac, 'ni': ni, 'dof': dof})
        for c, d, a in e['ni']:
            if a and a != '...':
                for nom, ac in blocks(a):
                    out.append({'nom': nom, 'ac': ac, 'ni': [c], 'dof': dof})
    return out


for e in base:
    db.setdefault((e['n'], e['s'], e['f']), {'b': []})['b'] += entry_blocks(e, CFG['base']['dof'])
for mod in CFG['modificaciones']:
    t = anexo_text(mod['codigo'])
    dof = mod['dof']
    for e in parse_anexo(t, amend=True):
        k = (e['n'], e['s'], e['f'])
        old = db.get(k)
        newb = entry_blocks(e, dof) if e['n'] == '3' else [b for b in entry_blocks(e, dof) if b['nom'] or b['ac']]
        ni_new = [x[0] for x in e['ni']]
        if old is None:
            db[k] = {'b': newb}
            continue
        keep = e['a'].startswith('...') or (e['a'] == '' and e['n'] != '3')
        oldni = sorted(set(n for b in old['b'] for n in b['ni']))
        allni = sorted(set(oldni) | set(ni_new))
        if keep:
            for b in old['b']:
                b['ni'] = sorted(set(b['ni']) | set(ni_new))
            add = [b for b in newb if b['nom'] and b['nom'] not in [x['nom'] for x in old['b']]]
            for b in add:
                b['ni'] = allni
            old['b'] += add
        else:
            for b in newb:
                if not b['ni'] or '...' in t:
                    b['ni'] = sorted(set(b['ni']) | set(oldni))
            old['b'] = newb
for x in CFG.get('eliminar_fracciones', []):
    db.pop((x['numeral'], x['seccion'], x['fraccion']), None)
for x in CFG.get('suprimir_nicos', []):
    k = (x['numeral'], x['seccion'], x['fraccion'])
    for b in db.get(k, {'b': []})['b']:
        if x['nico'] in b['ni'] and len(b['ni']) > 1:
            b['ni'].remove(x['nico'])
NM = {}
for (n, s, f), v in sorted(db.items(), key=lambda kv: (kv[0][2], kv[0][0])):
    for b in v['b']:
        NM.setdefault(f, []).append([n, s or '', b['nom'], b['ac'], ' '.join(b['ni']), b['dof']])
fechas = [CFG['base']['dof']] + [m_['dof'] for m_ in CFG['modificaciones']] + [v['dof'] for v in CFG.get('encabezados', {}).values()]
fnom = max(fechas, key=dmy)
log(f'   {len(NM):,} fracciones con NOM · última modificación aplicada: DOF {fnom}')
fuera = sorted(f for f in NM if f not in FA)
if fuera:
    log(f'   AVISO: {len(fuera)} fracciones del Anexo 2.4.1 ya no existen en la TIGIE: ' + ', '.join(fuera[:15]))
solo_bu = sorted(BU_NOM - set(NM))
solo_nm = sorted(set(NM) - BU_NOM)
log(f'   Comparación con la Base Única: coinciden {len(BU_NOM & set(NM)):,}; solo en la Base Única {len(solo_bu)}; solo en el anexo {len(solo_nm)}')
if solo_bu:
    log('     solo en la Base Única: ' + ', '.join(solo_bu[:20]))

# ------------------------------------------------------------------ 5. acuerdos nuevos que revisar
log('4) Acuerdos nuevos de la SE que podrían cambiar NOM')
nuevos = []
if not OFFLINE:
    try:
        h = get(SNICE + '/cs/avi/snice/historicos.dof.html').decode('utf-8', 'ignore')
        t = html.unescape(re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' | ', h)))
        rev = [dmy(x) for x in CFG.get('revisados', [])]
        ultimo = max(rev) if rev else (2022, 5, 9)
        vistos = 0
        for mm in re.finditer(r'(\d{2})\.(\d{2})\.(\d{4})(?:\s*\|)+\s*(Acuerdo[^|]{0,160}reglas y criterios[^|]*)', t, re.I):
            vistos += 1
            fe = (int(mm.group(3)), int(mm.group(2)), int(mm.group(1)))
            if fe > ultimo:
                nuevos.append(f'{mm.group(1)}/{mm.group(2)}/{mm.group(3)} · {mm.group(4).strip()}')
        if not vistos:
            log('   AVISO: no encontré ningún acuerdo en el histórico del DOF del SNICE; revisa a mano si hay modificaciones al Anexo 2.4.1')
        else:
            log(f'   leídos {vistos} acuerdos de reglas y criterios en el histórico del SNICE')
    except Exception as ex:  # el aviso es de ayuda; no detiene la actualización
        log(f'   no pude revisar el histórico del DOF en el SNICE ({ex})')
if nuevos:
    log('   ATENCIÓN: hay acuerdos nuevos sin revisar. Si modifican el Anexo 2.4.1, agrégalos a tools/nom_mods.json:')
    for x in sorted(set(nuevos)):
        log('     - ' + x)
else:
    log('   ninguno nuevo desde el último revisado')

# ------------------------------------------------------------------ 6. reporte de cambios contra la página actual
log('5) Cambios contra la versión anterior de la página')
oldfr = {x[0]: x for x in OLD['fr']}
newfr = {x[0]: x for x in FR}
alta, baja = sorted(set(newfr) - set(oldfr)), sorted(set(oldfr) - set(newfr))
arancel = [(f, oldfr[f][3], newfr[f][3], oldfr[f][4], newfr[f][4]) for f in sorted(set(newfr) & set(oldfr)) if (oldfr[f][3], oldfr[f][4]) != (newfr[f][3], newfr[f][4])]
desc = [f for f in sorted(set(newfr) & set(oldfr)) if (oldfr[f][1], oldfr[f][2]) != (newfr[f][1], newfr[f][2])]
nico = [f for f in sorted(set(newfr) & set(oldfr)) if [n[0] for n in oldfr[f][7]] != [n[0] for n in newfr[f][7]]]
oldnm = OLDNOM.get('nm', {})
nom_alta = sorted(set(NM) - set(oldnm))
nom_baja = sorted(set(oldnm) - set(NM))
nom_cambio = [f for f in sorted(set(NM) & set(oldnm)) if NM[f] != oldnm[f]]
rg_cambio = sorted(f for f in set(RG) | set(OLDRG) if RG.get(f) != OLDRG.get(f))
log(f'   fracciones nuevas: {len(alta)}' + (' → ' + ', '.join(alta[:30]) if alta else ''))
log(f'   fracciones eliminadas: {len(baja)}' + (' → ' + ', '.join(baja[:30]) if baja else ''))
log(f'   cambios de arancel: {len(arancel)}')
for f, i0, i1, e0, e1 in arancel[:40]:
    log(f'     {f}: IMP {i0 or "—"} → {i1 or "—"} · EXP {e0 or "—"} → {e1 or "—"}')
log(f'   cambios de descripción o unidad: {len(desc)}' + (' → ' + ', '.join(desc[:20]) if desc else ''))
log(f'   cambios de NICO: {len(nico)}' + (' → ' + ', '.join(nico[:20]) if nico else ''))
log(f'   NOM: {len(nom_alta)} fracciones nuevas, {len(nom_baja)} quitadas, {len(nom_cambio)} con cambios')
log(f'   regulaciones/programas (Base Única): {len(rg_cambio)} fracciones con cambios')
if len(baja) > 400:
    fail(f'se eliminarían {len(baja)} fracciones de golpe; revisa el archivo del SNICE antes de publicar')

# ------------------------------------------------------------------ 7. escribir la página
J = lambda o: json.dumps(o, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
DATA = {'cap': CAP, 'pt': PT, 'fr': FR}
NOMD = {'nm': NM, 'ns': sechead, 'fnom': fnom, 'fbu': fecha_corta(bu_fecha)}
out = src
for m_, id_, obj in sorted([(m_rg, 'rg', RG), (m_nom, 'nom', NOMD), (m_data, 'data', DATA)], key=lambda x: -x[0].start()):
    out = out[:m_.start()] + f'<script id="{id_}" type="application/json">' + J(obj) + '</script>' + out[m_.end():]
d_fa = datetime.date(int(fa_fecha[:4]), int(fa_fecha[4:6]), int(fa_fecha[6:]))
d_nom = datetime.date(*reversed([int(x) for x in fnom.split('/')]))
stamp = f'TIGIE al {d_fa.day} {MESES[d_fa.month - 1][:3]} {d_fa.year} · NOM al {d_nom.day} {MESES[d_nom.month - 1][:3]} {d_nom.year}'
mods_txt = [v['dof'] for v in CFG.get('encabezados', {}).values()] + [m_['dof'] for m_ in CFG['modificaciones']]
mods_txt = sorted(set(mods_txt), key=dmy)
mods_txt = ', '.join(mods_txt[:-1]) + ' y ' + mods_txt[-1] if len(mods_txt) > 1 else ''.join(mods_txt)
base_d = datetime.date(*reversed([int(x) for x in CFG['base']['dof'].split('/')]))
fuentes = (f'Fracciones, aranceles y NICO: SNICE / Secretaría de Economía, archivo del {fecha_larga(fa_fecha)}. '
           f'Texto de partidas y subpartidas: «Base Única» del SNICE. '
           f'NOM: Anexo 2.4.1 del Acuerdo de reglas y criterios de la SE (versión del DOF del {base_d.day} de {MESES[base_d.month - 1]} de {base_d.year} con sus modificaciones del {mods_txt}). '
           f'Demás regulaciones, PROSEC, IMMEX, frontera y precios estimados: «Base Única» del SNICE del {fecha_larga(bu_fecha)}.')
for key, val in (('stamp', stamp), ('fuentes', fuentes)):
    out, n = re.subn(r'<!--M:' + key + r'-->.*?<!--/M-->', '<!--M:' + key + '-->' + html.escape(val, quote=False) + '<!--/M-->', out, flags=re.S)
    if n != 1:
        fail('no encontré la marca <!--M:' + key + '--> en index.html')
cambio = out != src
if cambio:
    open(PAGE, 'w', encoding='utf-8', newline='\n').write(out)
log('')
log('Página actualizada.' if cambio else 'Sin cambios: la página ya estaba al día.')
open(REPORT, 'w', encoding='utf-8').write(f'Actualización {datetime.datetime.now():%d/%m/%Y %H:%M}\n\n' + '\n'.join(rep_lines) + '\n')
sys.exit(0)
