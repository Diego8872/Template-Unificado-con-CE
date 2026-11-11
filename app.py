import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import pdfplumber
import re
import io
import math
import unicodedata
from collections import defaultdict
from pathlib import Path

# ═════════════════════════════════════════════════════════════════════════════
#  LÓGICA (sin dependencias de Streamlit)
# ═════════════════════════════════════════════════════════════════════════════

COL_CE = 'D:LPCOVUCE0044'
COL_ORDEN = 'I:NRO-ORDEN-LPCO'
COL_OPC = 'I:MINERIAOPC'
COL_MIN = 'V:MINERIA'
COLS_NUEVAS = [COL_CE, COL_ORDEN, COL_OPC, COL_MIN]
COLS_REQUERIDAS = ['NumeroDeFactura', 'Observaciones', 'PosicionArancelaria',
                   'Cantidad', 'UnidadMedida', 'Descripcion']

UM_NOMBRES = {'UNIDAD': '07', 'CENTIMETRO': '20', 'METRO': '02'}


def safe_float(v):
    try:
        f = float(str(v).replace(',', '.'))
        return 0.0 if math.isnan(f) else f
    except Exception:
        return 0.0


def parse_cantidad(s):
    s = str(s).strip().rstrip('.,')
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except Exception:
        return 0.0


def sin_acentos(s):
    return unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()


def norm_txt(s):
    return re.sub(r'[^a-z0-9]', '', sin_acentos(s).lower())


def solo_digitos(s):
    return re.sub(r'\D', '', str(s or ''))


def ncm8(s):
    return solo_digitos(s)[:8]


def codigo_um(v):
    """'07 - UNIDAD' → '07'; 'UNIDAD' → '07'; desconocida → None (comodín)."""
    s = sin_acentos(v).strip().upper()
    if not s or s == 'NAN':
        return None
    m = re.match(r'^(\d{1,2})\b', s)
    if m:
        return m.group(1).zfill(2)
    return UM_NOMBRES.get(s)


def um_compatible(a, b):
    return a is None or b is None or a == b


def v_celda(df, i, col):
    if col not in df.columns:
        return ''
    v = df.at[i, col]
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ''
    return str(v).strip()


# ── Lectura de PDFs ─────────────────────────────────────────────────────────

def leer_pdf(file):
    """Devuelve (texto, numero_documento_desde_anotaciones)."""
    data = file.getvalue() if hasattr(file, 'getvalue') else file.read()
    texto, nro = '', None
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                texto += (page.extract_text() or '') + '\n'
                if nro is None:
                    for annot in (page.annots or []):
                        d = annot.get('data', {})
                        t = d.get('T', b'')
                        if isinstance(t, bytes):
                            t = t.decode('latin-1', errors='ignore')
                        if 'numero_documento' in str(t).lower():
                            val = d.get('V', b'')
                            if isinstance(val, bytes):
                                val = val.decode('latin-1', errors='ignore')
                            val = str(val or '').strip()
                            if re.match(r'^(CE|IF)-', val):
                                nro = val
    except Exception:
        pass
    return texto, nro


def numero_desde_nombre(filename):
    m = re.search(r'((?:CE|IF)-\d{4}-\d+-APN-[A-Z]+?)(?:#|23)MEC', filename)
    return m.group(1) + '#MEC' if m else None


def numero_desde_texto(texto):
    m = re.search(r'N[úu]mero:\s*((?:CE|IF)-\d{4}-\d+-APN-[A-Z]+)\s*#\s*MEC', texto)
    return m.group(1) + '#MEC' if m else None


def parsear_documento(texto, nro):
    """Interpreta un CE o un IF (Régimen de Inversiones Mineras)."""
    tipo = nro[:2]
    d = {'num': nro, 'tipo': tipo, 'id': int(re.search(r'-(\d{5,})-', nro).group(1))}
    m = re.search(r'CUIT:\s*(\d{11})', texto)
    d['cuit'] = m.group(1) if m else None
    m = re.search(r'Referencia:\s*(EX-\d{4}-\d+)', texto)
    d['ex'] = m.group(1) if m else None
    d['ex_id'] = int(d['ex'].split('-')[2]) if d['ex'] else None
    m = re.search(r'Proyecto minero[^:\n]*:\s*(.+)', texto)
    d['proyecto'] = re.sub(r'\s+', ' ', m.group(1)).strip() if m else None

    d['obs'], d['mar'], d['factura'], d['parte'], d['int'] = None, None, None, None, None
    if tipo == 'IF':
        i = texto.rfind('Observaciones')
        if i >= 0:
            seg = texto[i + len('Observaciones'):].split('Digitally signed')[0]
            seg = re.sub(r'\s+', ' ', seg).strip()
            d['obs'] = seg
            m = re.search(r'(MAR\d+)', seg)
            d['mar'] = m.group(1) if m else None
            m = re.search(r'\b([A-Z]{1,2}\d{6,})\b', seg)
            d['factura'] = m.group(1) if m else None
            m = re.search(r'PARTE\s*(\d+)', seg, re.I)
            d['parte'] = int(m.group(1)) if m else None
            m = re.search(r'Int:\s*(\d+)', seg)
            d['int'] = int(m.group(1)) if m else None

    ordenes = []
    for b in re.split(r'N[úu]mero de Orden[^:]{0,200}:\s*', texto)[1:]:
        mn = re.match(r'\s*(\d+)', b)
        if not mn:
            continue
        o = {'n': int(mn.group(1))}
        m = re.search(r'NCM[^:]{0,120}:\s*([\d.]+)', b)
        o['ncm'] = ncm8(m.group(1)) if m else None
        m = re.search(r'Cantidad[^:]{0,120}:\s*([\d.,]+)', b)
        o['q'] = parse_cantidad(m.group(1)) if m else 0.0
        m = re.search(r'Unidad de medida[^:]{0,120}:\s*(\d+)', b)
        o['um'] = m.group(1).zfill(2) if m else None
        m = re.search(r'Descripci[óo]n comercial del bien\s*(.+?)'
                      r'(?=Informaci[óo]n sobre la mercader|Condici[oó]n de la mercader|'
                      r'Digitally signed|Declaraci[óo]n Jurada|$)', b, re.S)
        o['desc'] = re.sub(r'\s+', ' ', m.group(1)).strip() if m else ''
        ordenes.append(o)
    d['orders'] = ordenes
    return d


def clasificar_pdfs(archivos):
    """Devuelve (docs_por_numero, no_reconocidos)."""
    docs, no_rec = {}, []
    for f in archivos:
        texto, nro_anot = leer_pdf(f)
        nro = numero_desde_texto(texto) or nro_anot or numero_desde_nombre(f.name)
        if not nro:
            no_rec.append(f.name)
            continue
        d = parsear_documento(texto, nro)
        d['archivo'] = f.name
        if not d['orders']:
            no_rec.append(f.name)
            continue
        docs[nro] = d
    return docs, no_rec


# ── Emparejamiento CE ↔ IF ──────────────────────────────────────────────────

def firma(d, con_desc):
    return [(o['n'], o['ncm'], round(o['q'], 6), o['um'],
             norm_txt(o['desc']) if con_desc else None) for o in d['orders']]


def emparejar(ces, ifs):
    """Une cada CE con su IF por contenido. Devuelve (pares, ces_sin_if, ifs_sin_ce, ambiguos)."""
    cand = {}
    for c in ces:
        mismo_cuit = [i for i in ifs if i['cuit'] == c['cuit']]
        exactos = [i for i in mismo_cuit if firma(i, True) == firma(c, True)]
        if exactos:
            cand[c['num']] = [(i, False) for i in exactos]
        else:
            cand[c['num']] = [(i, True) for i in mismo_cuit if firma(i, False) == firma(c, False)]

    asign, usados = {}, set()
    pend = set(cand)
    cambio = True
    while cambio:
        cambio = False
        for cn in sorted(pend):
            libres = [(i, r) for i, r in cand[cn] if i['num'] not in usados]
            if len(libres) == 1:
                asign[cn] = libres[0]
                usados.add(libres[0][0]['num'])
                pend.discard(cn)
                cambio = True
    por_num = {c['num']: c for c in ces}
    ambiguos = []
    for cn in sorted(pend):
        libres = [(i, r) for i, r in cand[cn] if i['num'] not in usados]
        if not libres:
            continue
        ex = por_num[cn]['ex_id']

        def dist(par):
            i = par[0]
            return (ex - i['id']) if (ex and ex - i['id'] >= 0) else 10 ** 12
        libres.sort(key=dist)
        if len(libres) == 1 or dist(libres[0]) < dist(libres[1]):
            asign[cn] = libres[0]
            usados.add(libres[0][0]['num'])
        else:
            ambiguos.append((por_num[cn], [i['num'] for i, _ in libres]))
    pares = [{'ce': por_num[cn], 'if': i, 'relajado': r} for cn, (i, r) in asign.items()]
    ambiguos_ids = {c['num'] for c, _ in ambiguos}
    ces_sin_if = [c for c in ces if c['num'] not in asign and c['num'] not in ambiguos_ids]
    ifs_sin_ce = [i for i in ifs if i['num'] not in usados]
    return pares, ces_sin_if, ifs_sin_ce, ambiguos


# ── Asignación al template ──────────────────────────────────────────────────

def preparar_filas(df):
    filas = []
    for i in df.index:
        obs = v_celda(df, i, 'Observaciones')
        m = re.search(r'PARTE\s*(\d+)', obs, re.I)
        filas.append({
            'ix': i,
            'fac': v_celda(df, i, 'NumeroDeFactura').upper(),
            'cm': 'SIN CM' not in obs.upper(),
            'parte': int(m.group(1)) if m else None,
            'ncm': ncm8(v_celda(df, i, 'PosicionArancelaria')),
            'q': safe_float(v_celda(df, i, 'Cantidad')),
            'um': codigo_um(v_celda(df, i, 'UnidadMedida')),
            'desc': norm_txt(v_celda(df, i, 'Descripcion')),
            'desc_orig': v_celda(df, i, 'Descripcion'),
        })
    return filas


def col_item_de(df):
    for c in ('ITEM_DESPACHO', 'ITEM'):
        if c in df.columns:
            return c
    return None


def asignar_lineas(df, pares, ncm_map):
    col_item = col_item_de(df)
    filas = preparar_filas(df)
    usadas = set()
    asign = {}
    control, avisos, alertas = [], [], []

    def item_de(ix):
        return v_celda(df, ix, col_item) if col_item else ''

    por_fac = defaultdict(list)
    for p in pares:
        por_fac[(p['if']['factura'] or '').upper()].append(p)

    for fac, lista in por_fac.items():
        lista.sort(key=lambda p: (p['if']['parte'] or 0, p['if']['id']))
        filas_fac = [f for f in filas if f['fac'] == fac and f['cm']]
        if not fac or not filas_fac:
            for p in lista:
                alertas.append({'nivel': 'ALERTA', 'tipo': 'Factura no encontrada',
                                'ce': p['ce']['num'], 'orden': '', 'item': '', 'fila': '',
                                'detalle': f"El {p['if']['num']} indica la factura '{fac or 'sin dato'}' "
                                           f"y el template no tiene ítems con CM de esa factura."})
                p['res'] = {'asignadas': 0, 'sin_item': len(p['ce']['orders']), 'avisos': 0,
                            'ncm': 0, 'sobran': 0}
            continue

        for p in lista:
            ce, i_f = p['ce'], p['if']
            parte = i_f['parte']
            if parte:
                g = [f for f in filas_fac if f['parte'] == parte] or [f for f in filas_fac if f['parte'] is None]
                agrupado = True
            else:
                g = filas_fac
                agrupado = False
            p['grupo'], p['agrupado'] = g, agrupado
            libres = [f for f in g if f['ix'] not in usadas]
            ordenes = sorted(i_f['orders'], key=lambda o: o['n'])
            n_avisos = 0
            pendientes = []
            for o in ordenes:        # 1) coincidencia completa
                hit = next((f for f in libres if f['ncm'] == o['ncm'] and abs(f['q'] - o['q']) < 1e-6
                            and um_compatible(f['um'], o['um']) and f['desc'] == norm_txt(o['desc'])), None)
                if hit is None:
                    pendientes.append(o)
                else:
                    asign[hit['ix']] = (ce['num'], o['n'])
                    usadas.add(hit['ix'])
                    libres.remove(hit)
            sin_item = []
            for o in pendientes:     # 2) sin descripción → aviso
                hit = next((f for f in libres if f['ncm'] == o['ncm'] and abs(f['q'] - o['q']) < 1e-6
                            and um_compatible(f['um'], o['um'])), None)
                if hit is None:
                    sin_item.append(o)
                    alertas.append({'nivel': 'ALERTA', 'tipo': 'Orden sin ítem', 'ce': ce['num'],
                                    'orden': o['n'], 'item': '', 'fila': '',
                                    'detalle': f"No hay ítem en el template (factura {fac}) para NCM {o['ncm']}, "
                                               f"cantidad {o['q']:g}, '{o['desc'][:60]}'."})
                    continue
                asign[hit['ix']] = (ce['num'], o['n'])
                usadas.add(hit['ix'])
                libres.remove(hit)
                n_avisos += 1
                avisos.append({'nivel': 'AVISO', 'tipo': 'Descripción distinta', 'ce': ce['num'],
                               'orden': o['n'], 'item': item_de(hit['ix']), 'fila': hit['ix'] + 2,
                               'desc_ce': o['desc'], 'desc_xls': hit['desc_orig'],
                               'detalle': 'Coinciden NCM, cantidad y unidad; la descripción del IF/CE difiere del template.'})
            if p['relajado']:
                n_avisos += 1
                avisos.append({'nivel': 'AVISO', 'tipo': 'CE e IF con descripción distinta', 'ce': ce['num'],
                               'orden': '', 'item': '', 'fila': '', 'desc_ce': '', 'desc_xls': '',
                               'detalle': f"El {ce['num']} y el {i_f['num']} coinciden en NCM, cantidad y unidad, "
                                          f"pero alguna descripción difiere."})
            p['res'] = {'asignadas': len(ordenes) - len(sin_item), 'sin_item': len(sin_item),
                        'avisos': n_avisos, 'ncm': 0, 'sobran': 0}

        # ítems con CM que quedaron sin orden (por parte o por factura)
        sobrantes_fac = [f for f in filas_fac if f['ix'] not in usadas]
        if sobrantes_fac:
            if any(p.get('agrupado') for p in lista):
                for p in lista:
                    if p.get('agrupado'):
                        p['res']['sobran'] = len([f for f in p['grupo'] if f['ix'] not in usadas])
            else:
                lista[-1]['res']['sobran'] = len(sobrantes_fac)

    # ítems CM de facturas sin ningún CE/IF asociado
    sin_ce_por_fac = defaultdict(list)
    for f in filas:
        if f['cm'] and f['ix'] not in usadas:
            sin_ce_por_fac[f['fac']].append(f)
    for fac, fl in sin_ce_por_fac.items():
        alertas.append({'nivel': 'AVISO', 'tipo': 'Ítems con CM sin CE', 'ce': '', 'orden': '',
                        'item': ', '.join(item_de(f['ix']) for f in fl[:8]) + ('…' if len(fl) > 8 else ''),
                        'fila': '', 'detalle': f"Factura {fac}: {len(fl)} ítem(s) con CM sin CE asignado."})

    # MINERIA / NCM
    min_opc = {}
    fuera = defaultdict(list)
    for ix, (cen, n) in asign.items():
        pos_arancel = v_celda(df, ix, 'PosicionArancelaria')
        opc = ncm_map.get(ncm8(pos_arancel), '')
        min_opc[ix] = opc
        if not opc:
            fuera[(cen, pos_arancel)].append((n, item_de(ix), ix + 2))
            for p in pares:
                if p['ce']['num'] == cen and 'res' in p:
                    p['res']['ncm'] += 1
    for (cen, pos_arancel), lst in fuera.items():
        ordenes_txt = ', '.join(str(n) for n, _, _ in sorted(lst)[:12]) + ('…' if len(lst) > 12 else '')
        alertas.append({'nivel': 'ALERTA', 'tipo': 'NCM fuera del listado', 'ce': cen, 'orden': ordenes_txt,
                        'item': ', '.join(str(it) for _, it, _ in lst[:12]) + ('…' if len(lst) > 12 else ''),
                        'fila': ', '.join(str(f) for _, _, f in lst[:12]) + ('…' if len(lst) > 12 else ''),
                        'detalle': f"NCM {pos_arancel}: {len(lst)} ítem(s) con CE y no figura en el listado "
                                   f"Res. 73/26 SM (no se pudo completar I:MINERIAOPC)."})

    # proyecto / CUIT (coherencia IF ↔ template)
    for p in pares:
        i_f = p['if']
        fac = (i_f['factura'] or '').upper()
        filas_p = [f for f in filas if f['fac'] == fac and f['cm']]
        if not filas_p:
            continue
        proy_x = {v_celda(df, f['ix'], 'ProyectoMinero').upper() for f in filas_p} - {''}
        if proy_x and i_f['proyecto'] and i_f['proyecto'].upper().strip() not in proy_x:
            avisos.append({'nivel': 'AVISO', 'tipo': 'Proyecto minero distinto', 'ce': p['ce']['num'],
                           'orden': '', 'item': '', 'fila': '', 'desc_ce': i_f['proyecto'],
                           'desc_xls': ' / '.join(sorted(proy_x)),
                           'detalle': f"El proyecto del {i_f['num']} no coincide con el del template (factura {fac})."})
            p['res']['avisos'] += 1
        cuit_x = {solo_digitos(v_celda(df, f['ix'], 'CUIT')) for f in filas_p} - {''}
        if cuit_x and i_f['cuit'] and i_f['cuit'] not in cuit_x:
            avisos.append({'nivel': 'AVISO', 'tipo': 'CUIT distinto', 'ce': p['ce']['num'], 'orden': '',
                           'item': '', 'fila': '', 'desc_ce': i_f['cuit'], 'desc_xls': ' / '.join(sorted(cuit_x)),
                           'detalle': f"El CUIT del {i_f['num']} no coincide con el del template."})
            p['res']['avisos'] += 1

    # control por CE
    for p in pares:
        r = p['res']
        if r['sin_item'] or r['ncm'] or r['sobran']:
            estado = 'ALERTA'
        elif r['avisos']:
            estado = 'OK con aviso'
        else:
            estado = 'OK'
        control.append({'ce': p['ce']['num'], 'if': p['if']['num'], 'ex': p['ce']['ex'] or '',
                        'factura': p['if']['factura'] or '', 'parte': p['if']['parte'] or '',
                        'ord_ce': len(p['ce']['orders']), 'asignadas': r['asignadas'],
                        'sin_item': r['sin_item'], 'sobran': r['sobran'], 'estado': estado})
    control.sort(key=lambda c: (c['factura'], c['parte'] or 0, c['ce']))

    # salida: columnas nuevas junto a la columna de ítem
    out = df.copy()
    for c in COLS_NUEVAS:
        if c in out.columns:
            out = out.drop(columns=[c])
    col_item = col_item_de(out)
    pos = out.columns.tolist().index(col_item) + 1 if col_item else len(out.columns)
    ce_col = [asign[i][0] if i in asign else '' for i in out.index]
    or_col = [asign[i][1] if i in asign else '' for i in out.index]
    op_col = [min_opc.get(i, '') if i in asign else '' for i in out.index]
    mi_col = ['SI' if i in asign else '' for i in out.index]
    for k, (nombre, valores) in enumerate(zip(COLS_NUEVAS, [ce_col, or_col, op_col, mi_col])):
        out.insert(pos + k, nombre, valores)
    return out, control, avisos, alertas


def alertas_de_emparejamiento(ces_sin_if, ifs_sin_ce, ambiguos):
    al = []
    for c in ces_sin_if:
        al.append({'nivel': 'ALERTA', 'tipo': 'CE sin IF', 'ce': c['num'], 'orden': '', 'item': '', 'fila': '',
                   'detalle': f"No hay un IF con el mismo CUIT y las mismas {len(c['orders'])} órdenes "
                              f"(NCM, cantidad, unidad)."})
    for i in ifs_sin_ce:
        al.append({'nivel': 'ALERTA', 'tipo': 'IF sin CE', 'ce': '', 'orden': '', 'item': '', 'fila': '',
                   'detalle': f"{i['num']} (factura {i['factura'] or '?'}"
                              f"{', parte ' + str(i['parte']) if i['parte'] else ''}) no tiene CE cargado."})
    for c, nums in ambiguos:
        al.append({'nivel': 'ALERTA', 'tipo': 'CE ambiguo', 'ce': c['num'], 'orden': '', 'item': '', 'fila': '',
                   'detalle': 'Varios IF tienen el mismo contenido: ' + ', '.join(nums) + '. No se asignó ninguno.'})
    return al


def leer_listado_ncm(path):
    """Listado Anexo II Res. 73/26 SM → {ncm_8_digitos: MIN001/MIN002}."""
    try:
        d = pd.read_excel(path, dtype=str)
        return {ncm8(r['NCM']): str(r['A DECLARAR EN SIM']).strip()
                for _, r in d.iterrows()
                if pd.notna(r['NCM']) and pd.notna(r['A DECLARAR EN SIM'])}
    except Exception:
        return {}


def ruta_listado_ncm():
    for base in (Path(__file__).resolve().parent, Path.cwd()):
        p = base / 'ncm_mineria.xlsx'
        if p.exists():
            return p
    return None


# ── Excel de salida ─────────────────────────────────────────────────────────

def exportar_excel(df, control, avisos_y_alertas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'general'
    FILL_AM = PatternFill('solid', fgColor='FFFF00')
    FILL_TIT = PatternFill('solid', fgColor='FFF2CC')
    FH = Font(name='Arial', size=11, bold=True)
    FD = Font(name='Calibri', size=11)
    SIN_COLOR = {'ID', 'InscRUMP', 'ActiServ', 'NroInsc', 'RazonSocial', 'CUIT', 'ImpDirecta', 'CondMerca',
                 'SimiSira', 'ProyectoMinero', 'Radicacion', 'ClasificacionDeArticulo', 'TipoDeFactura',
                 'Observaciones', 'ITEM_DESPACHO', 'ITEM', 'D:CERTSM'} | set(COLS_NUEVAS)
    for ci, col in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font = FH
        cell.alignment = Alignment(horizontal='left', vertical='center')
        if col not in SIN_COLOR:
            cell.fill = FILL_AM
    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, val in enumerate(row, 1):
            v = None if (val == '' or (isinstance(val, float) and math.isnan(val))) else val
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.font = FD
            cell.alignment = Alignment(vertical='center')
    for col in ws.columns:
        ml = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(ml + 2, 10), 45)
    ws.freeze_panes = 'A2'

    def hoja(nombre, cabeceras, filas):
        s = wb.create_sheet(nombre)
        s.append(cabeceras)
        for f in filas:
            s.append(f)
        for c in s[1]:
            c.font = Font(name='Arial', size=11, bold=True)
            c.fill = FILL_TIT
        for col in s.columns:
            ml = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            s.column_dimensions[col[0].column_letter].width = min(max(ml + 2, 12), 60)
        s.freeze_panes = 'A2'

    hoja('Control_CE',
         ['CE', 'IF', 'EX', 'Factura', 'Parte', 'Órdenes CE', 'Órdenes asignadas', 'Órdenes sin ítem',
          'Ítems CM sin orden', 'Estado'],
         [[c['ce'], c['if'], c['ex'], c['factura'], c['parte'], c['ord_ce'], c['asignadas'], c['sin_item'],
           c['sobran'], c['estado']] for c in control])
    filas = [[a['nivel'], a['tipo'], a['ce'], a['orden'], a['item'], a['fila'],
              a.get('desc_ce', ''), a.get('desc_xls', ''), a['detalle']] for a in avisos_y_alertas]
    if not filas:
        filas = [['', 'Sin avisos', '', '', '', '', '', '', '']]
    hoja('Avisos', ['Nivel', 'Tipo', 'CE', 'Orden', 'ITEM_DESPACHO', 'Fila Excel',
                    'Descripción IF/CE', 'Descripción Excel', 'Detalle'], filas)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def procesar(df, docs, ncm_map):
    """Orquesta todo el análisis. Devuelve un dict con resultados."""
    ces = [d for d in docs.values() if d['tipo'] == 'CE']
    ifs = [d for d in docs.values() if d['tipo'] == 'IF']
    pares, ces_sin_if, ifs_sin_ce, ambiguos = emparejar(ces, ifs)
    out, control, avisos, alertas = asignar_lineas(df, pares, ncm_map)
    al_emp = alertas_de_emparejamiento(ces_sin_if, ifs_sin_ce, ambiguos)
    for c in ces_sin_if:
        control.append({'ce': c['num'], 'if': '', 'ex': c['ex'] or '', 'factura': '', 'parte': '',
                        'ord_ce': len(c['orders']), 'asignadas': 0, 'sin_item': len(c['orders']),
                        'sobran': 0, 'estado': 'SIN IF'})
    for i in ifs_sin_ce:
        control.append({'ce': '', 'if': i['num'], 'ex': '', 'factura': i['factura'] or '',
                        'parte': i['parte'] or '', 'ord_ce': len(i['orders']), 'asignadas': 0,
                        'sin_item': 0, 'sobran': 0, 'estado': 'SIN CE'})
    todos = al_emp + alertas + avisos
    todos.sort(key=lambda a: 0 if a['nivel'] == 'ALERTA' else 1)
    return {'df': out, 'control': control, 'avisos': todos, 'n_ce': len(ces), 'n_if': len(ifs),
            'n_pares': len(pares)}


# ═════════════════════════════════════════════════════════════════════════════
#  INTERFAZ
# ═════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Template Unificado con CE | INTERLOG", page_icon="📜", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    html, body, [class*="css"] { font-family: 'Roboto', sans-serif; background-color: #0e1117; color: #fafafa; }
    .block-container { padding-top: 3rem; }
    h1 { color: #00b4d8 !important; font-size: 1.8rem !important; }
    h2, h3 { color: #00b4d8 !important; }
    .subtitulo { font-size: 0.95rem; color: #8899aa; margin-bottom: 1.5rem; }
    .seccion { font-size: 1rem; font-weight: 600; color: #00b4d8; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 1rem; border-bottom: 1px solid #1e3a4a; padding-bottom: 6px; }
    .badge-ok { background: #0d2e1a; color: #3dd68c; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 500; border: 1px solid #3dd68c; display: inline-block; margin: 3px 2px; }
    .badge-err { background: #2e0d0d; color: #ff6b6b; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 500; border: 1px solid #ff6b6b; display: inline-block; margin: 3px 2px; }
    .badge-warn { background: #2e2200; color: #f0c040; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 500; border: 1px solid #f0c040; display: inline-block; margin: 3px 2px; }
    .match-box { background: #161b22; border: 1px solid #1e3a4a; border-radius: 8px; padding: 14px; margin: 6px 0; font-size: 0.92rem; color: #e0e0e0; }
    .match-ok { border-color: #3dd68c; }
    .match-warn { border-color: #f0c040; }
    .match-err { border-color: #ff6b6b; }
    .match-box b { color: #ffffff; }
    .match-box span { color: #aaaaaa; }
    .stButton > button { background: linear-gradient(135deg, #00b4d8, #0077b6) !important; color: white !important; font-weight: 600 !important; border: none !important; border-radius: 8px !important; }
    #MainMenu {visibility: hidden;} header {visibility: hidden;} footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


@st.cache_data
def cargar_ncm_mineria():
    """Listado Anexo II Res. 73/26 SM. El archivo ncm_mineria.xlsx debe estar junto a este archivo."""
    p = ruta_listado_ncm()
    return leer_listado_ncm(p) if p else {}


st.title("📜 Template Unificado con CE")
st.markdown('<p class="subtitulo">INTERLOG Comercio Exterior — Asignación de Certificados Mineros (CE) al Template Unificado</p>', unsafe_allow_html=True)
st.markdown("---")

nro_ref = st.text_input("Número de referencia de la operación", placeholder="ej: 982755")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    st.markdown('<p class="seccion">📊 Template Unificado</p>', unsafe_allow_html=True)
    f_unificado = st.file_uploader("Template Unificado (.xlsx)", type=["xlsx"])
with col2:
    st.markdown('<p class="seccion">📜 PDFs de CE</p>', unsafe_allow_html=True)
    f_ces = st.file_uploader("Certificados (CE-…) — múltiples", type=["pdf"], accept_multiple_files=True, key="ces")
with col3:
    st.markdown('<p class="seccion">📄 PDFs de IF</p>', unsafe_allow_html=True)
    f_ifs = st.file_uploader("Solicitudes (IF-…) — múltiples", type=["pdf"], accept_multiple_files=True, key="ifs")

st.markdown("---")

listo = bool(f_unificado and f_ces and f_ifs and nro_ref)
if not listo:
    st.info("📝 Ingresá el número de referencia, el Template Unificado y los PDFs de CE e IF para continuar.")

ncm_map = cargar_ncm_mineria()
if not ncm_map:
    st.error("No se encontró o no se pudo leer `ncm_mineria.xlsx` (debe estar junto a este archivo, con columnas "
             "'NCM' y 'A DECLARAR EN SIM'). Sin ese listado no se puede completar I:MINERIAOPC.")

if st.button("🔍 ANALIZAR Y ASIGNAR CE", disabled=not listo, use_container_width=True):
    with st.spinner("Procesando..."):
        df = pd.read_excel(f_unificado, dtype=str)
        df.columns = df.columns.str.strip()
        faltan = [c for c in COLS_REQUERIDAS if c not in df.columns]
        if faltan:
            st.error("Al Template le faltan columnas: " + ", ".join(faltan))
            st.session_state['procesado'] = False
        else:
            docs, no_rec = clasificar_pdfs(list(f_ces) + list(f_ifs))
            r = procesar(df, docs, ncm_map)
            r.update({'no_reconocidos': no_rec, 'nro_ref': nro_ref, 'procesado': True})
            st.session_state.update(r)

if st.session_state.get('procesado'):
    S = st.session_state
    control, avisos, df = S['control'], S['avisos'], S['df']
    alertas_rojas = [a for a in avisos if a['nivel'] == 'ALERTA']
    avisos_amar = [a for a in avisos if a['nivel'] != 'ALERTA']

    for nombre in S.get('no_reconocidos', []):
        st.markdown(f'<span class="badge-warn">⚠️ No se pudo leer como CE/IF: {nombre}</span>', unsafe_allow_html=True)

    if alertas_rojas:
        st.markdown("### ❌ Alertas")
        for a in alertas_rojas:
            ref = ' | '.join(str(x) for x in (a['ce'], f"orden {a['orden']}" if a['orden'] != '' else '',
                                              f"ITEM {a['item']}" if a['item'] else '') if x)
            st.markdown(f'<span class="badge-err">❌ {a["tipo"]}{" — " + ref if ref else ""}: {a["detalle"]}</span>',
                        unsafe_allow_html=True)
        st.warning("Revisá las alertas antes de descargar.")
    if avisos_amar:
        st.markdown("### ⚠️ Avisos")
        for a in avisos_amar:
            ref = ' | '.join(str(x) for x in (a['ce'], f"orden {a['orden']}" if a['orden'] != '' else '',
                                              f"ITEM {a['item']}" if a['item'] else '') if x)
            extra = ''
            if a.get('desc_ce') or a.get('desc_xls'):
                extra = f" — IF/CE: «{a.get('desc_ce', '')}» · Excel: «{a.get('desc_xls', '')}»"
            st.markdown(f'<span class="badge-warn">⚠️ {a["tipo"]}{" — " + ref if ref else ""}: {a["detalle"]}{extra}</span>',
                        unsafe_allow_html=True)

    st.markdown("### 📊 Resultado de la asignación")
    for c in control:
        if c['estado'] == 'OK':
            css, badge = 'match-ok', '<span class="badge-ok">✅ OK</span>'
        elif c['estado'] == 'OK con aviso':
            css, badge = 'match-warn', '<span class="badge-warn">⚠️ OK con aviso</span>'
        elif c['estado'] == 'SIN IF':
            css, badge = 'match-err', '<span class="badge-err">❌ CE sin IF</span>'
        elif c['estado'] == 'SIN CE':
            css, badge = 'match-err', '<span class="badge-err">❌ IF sin CE</span>'
        else:
            css, badge = 'match-err', '<span class="badge-err">❌ Alerta</span>'
        parte = f" — Parte {c['parte']}" if c['parte'] else ''
        st.markdown(f"""
        <div class="match-box {css}">
            {badge} <b>{c['ce'] or '(sin CE)'}</b><br>
            <span>IF: {c['if'] or '—'} | EX: {c['ex'] or '—'} | Factura: {c['factura'] or '—'}{parte}</span><br>
            <span>Órdenes: <b>{c['ord_ce']}</b> | Asignadas: <b>{c['asignadas']}</b> | Sin ítem: <b>{c['sin_item']}</b> | Ítems CM sin orden: <b>{c['sobran']}</b></span>
        </div>
        """, unsafe_allow_html=True)

    obs = df['Observaciones'].fillna('').astype(str).str.upper() if 'Observaciones' in df.columns else pd.Series([''] * len(df))
    total_sin_cm = int(obs.str.contains('SIN CM').sum())
    total_con_ce = int((df[COL_CE] != '').sum())
    total_cm_sin_ce = int(((df[COL_CE] == '') & (~obs.str.contains('SIN CM'))).sum())
    st.markdown(f"**CE: {S['n_ce']} | IF: {S['n_if']} | Pares: {S['n_pares']} — "
                f"Total: {total_con_ce} ítems con CE | {total_cm_sin_ce} ítems con CM sin CE | {total_sin_cm} SIN CM**")
    st.markdown("---")

    st.download_button(
        label="📥 DESCARGAR TEMPLATE UNIFICADO CON CE",
        data=exportar_excel(df, control, avisos),
        file_name=f'TEMPLATE_UNIFICADO_{S.get("nro_ref", "")}_CON_CE.xlsx',
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
