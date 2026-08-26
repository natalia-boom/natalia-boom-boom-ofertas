# -*- coding: utf-8 -*-
# Crea la tabla `facturas` (1 oferta -> N facturas) y carga la hoja "Ofertas Aprobadas".
# ACUMULA: dedup por dedup_key (N de factura si existe, si no un hash de la fila). Nunca borra.
import openpyxl, pg8000, re, hashlib, datetime

XLSX = r'C:\Users\nvargas\Downloads\CONTROL DE OFERTAS 2026-25 AGOSTO (3).xlsx'
DB = dict(host='hayabusa.proxy.rlwy.net', port=55500, database='railway',
          user='postgres', password='ihzPXTpSaAGzvHiUczbnFbKHUObbNZKj')

def s(v):
    if v is None: return None
    v = str(v).strip()
    return v if v else None

def norm_num(v):
    if v is None: return None
    txt = str(v).strip()
    m = re.match(r'^(\d{2})-(\d{3,4})', txt)
    if m: return m.group(1) + m.group(2).zfill(4)
    return None

def parse_valor(v):
    if v is None: return None
    if isinstance(v, (int, float)): return int(v)
    txt = str(v).replace('$', '').replace(',', '').replace(' ', '').strip()
    if txt in ('', '-', '.'): return None
    try: return int(float(txt))
    except: return None

# ---- leer hoja Ofertas Aprobadas ----
wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
ws = wb['Ofertas Aprobadas']
filas = []
rownum = 1
for r in ws.iter_rows(min_row=2, values_only=True):
    rownum += 1
    ref = s(r[1]); cliente = s(r[2]); estado = s(r[7])
    if not (cliente or estado):    # fila totalmente vacia
        continue
    mes = s(r[0])
    if isinstance(r[0], datetime.datetime): mes = None  # una fila trae fecha en MES, la ignoramos
    fila = dict(
        oferta_num = norm_num(ref),
        ref_original = ref,
        mes = mes,
        cliente = cliente,
        descripcion = s(r[3]),
        origen = s(r[4]),
        destino = s(r[5]),
        valor = parse_valor(r[6]),
        estado_proyecto = (estado.upper() if estado else None),
        responsable = s(r[10]),
        no_factura = s(r[11]),
    )
    # UNA FILA = UNA LINEA. OJO: en esta hoja un mismo N de factura cubre VARIAS ofertas,
    # por eso NO se deduplica por factura. Clave = fila fisica + hash de contenido
    # (idempotente si se reimporta el MISMO archivo; conserva todas las lineas).
    base = '|'.join(str(fila[k] or '') for k in
                    ('ref_original','descripcion','valor','estado_proyecto','mes','origen','destino','no_factura'))
    fila['dedup_key'] = 'R:%d:' % rownum + hashlib.md5(base.encode('utf-8')).hexdigest()[:12]
    filas.append(fila)

print('Filas leidas de la hoja:', len(filas))

c = pg8000.connect(**DB); cur = c.cursor()
cur.execute('''
    CREATE TABLE IF NOT EXISTS facturas (
        id bigserial PRIMARY KEY,
        oferta_num text,
        ref_original text,
        mes text,
        cliente text,
        descripcion text,
        origen text,
        destino text,
        valor bigint,
        estado_proyecto text,
        responsable text,
        no_factura text,
        dedup_key text UNIQUE,
        created_at timestamptz DEFAULT now()
    )
''')
cur.execute('CREATE INDEX IF NOT EXISTS idx_facturas_oferta ON facturas(oferta_num)')
c.commit()
print('[DB] tabla facturas lista.')

# dedup dentro del mismo archivo (por si dos filas comparten dedup_key) + insercion por lote
seen = set(); params = []
for f in filas:
    if f['dedup_key'] in seen: continue
    seen.add(f['dedup_key'])
    params.append((f['oferta_num'],f['ref_original'],f['mes'],f['cliente'],f['descripcion'],f['origen'],
                   f['destino'],f['valor'],f['estado_proyecto'],f['responsable'],f['no_factura'],f['dedup_key']))
cur.executemany('''INSERT INTO facturas
    (oferta_num,ref_original,mes,cliente,descripcion,origen,destino,valor,estado_proyecto,responsable,no_factura,dedup_key)
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (dedup_key) DO NOTHING''', params)
c.commit()
print(f'Filas enviadas al INSERT (unicas por clave): {len(params)}')

# ---- resumen ----
cur.execute('SELECT COUNT(*), COUNT(oferta_num) FILTER (WHERE oferta_num IS NOT NULL), COALESCE(SUM(valor),0) FROM facturas')
tot, conof, suma = cur.fetchone()
print(f'TOTAL facturas: {tot} | con N de oferta: {conof} | sin N (contratos/otros): {tot-conof}')
print(f'SUMA facturado en tabla facturas: {suma:,}')
cur.execute('SELECT COUNT(DISTINCT oferta_num) FROM facturas WHERE oferta_num IS NOT NULL')
print('Ofertas distintas enlazadas:', cur.fetchone()[0])
cur.execute("SELECT COALESCE(SUM(valor),0) FROM facturas WHERE estado_proyecto LIKE 'FACTURADO%'")
print('SUMA solo FACTURADO (real):', f'{cur.fetchone()[0]:,}')
cur.execute("SELECT estado_proyecto, COUNT(*), COALESCE(SUM(valor),0) FROM facturas GROUP BY estado_proyecto ORDER BY 2 DESC")
print('Por estado:')
for e,n,v in cur.fetchall(): print(f'   {e}: {n} filas, {v:,}')
c.close()
