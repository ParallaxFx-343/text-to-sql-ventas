"""Junta todos los movimientos de sep-oct 2026 por sucursal (local, concepto, fecha, paec, unidades)."""
import re, glob, csv, json, datetime as dt
from pathlib import Path
H = Path('/home/user/text-to-sql-ventas/ateneo/historial')
D = lambda d, m: dt.datetime(2026, m, d)
mov = []
def desglose_titulo(f):
    tot = {}
    for l in open(f, encoding='utf-8-sig'):
        m = re.match(r'^\s+(\d+)\s{2}(\S.*?)\s*$', l)
        if m and not l.strip().startswith(('ISBN', 'EAN')):
            tot[m.group(2)] = tot.get(m.group(2), 0) + int(m.group(1))
    return tot
def csv_u(f): return sum(int(l.split(';')[1]) for l in open(f, encoding='utf-8-sig').read().split() if ';' in l)
slug = lambda x: re.sub(r'[^A-Z0-9]+', '_', x).strip('_')
pad = json.load(open('/home/user/text-to-sql-ventas/ateneo/config.json'))['orden_paec'] + ['DEVOTO']
inv = {slug(p): p for p in pad}; inv.update({slug(p.upper()): p for p in pad})
# 11/09 promo 2x1 (un CSV por local)
for f in sorted(glob.glob(str(H / '2026-09/promo_2x1_csv/refuerzo_*.csv'))):
    mov.append((inv[Path(f).stem[len('refuerzo_'):]], 'Promo 2x1 $28.500', D(11, 9), None, csv_u(f)))
for f, c, d in ((H / '2026-09/dia_de_la_madre_csv/DESGLOSE_x_TITULO.txt', 'Refuerzo Día de la Madre', D(15, 9)),
                (H / '2026-09/DESGLOSE_AGENDA_Y_AUTORES.txt', 'Agenda + autores de octubre', D(17, 9)),
                (H / '2026-09/DESGLOSE_REPOSICION_21_AL_25_v22-09.txt', 'Repo 21 al 25/09', D(21, 9))):
    for loc, u in desglose_titulo(f).items():
        assert loc in pad, (f, loc); mov.append((loc, c, d, None, u))
# canales
for f, c, d in ((H / '2026-09/canales_ranking_17-09.csv', 'Canales ranking', D(17, 9)), (H / '2026-09/canales_refuerzo_23-09.csv', 'Canales refuerzo', D(23, 9)),
                (H / '2026-10/canales_refuerzo_09-10.csv', 'Canales refuerzo', D(9, 10)), (H / '2026-10/canales_refuerzo_14-10.csv', 'Canales refuerzo', D(14, 10)),
                (H / '2026-10/canales_refuerzo_19-10.csv', 'Canales refuerzo', D(19, 10))):
    mov.append(('CANALES', c, d, None, csv_u(f)))
mov.append(('DEVOTO', 'Repo Libros Sorpresa', D(5, 10), None, csv_u(H / '2026-10/reposicion_DEVOTO_libros_sorpresa.csv')))
# 07/10 refuerzo sorpresa con nro de pedido
txt = open(H / '2026-10/libros_sorpresa/refuerzo_195_07-10/DESGLOSE_x_SUCURSAL_REFUERZO_SORPRESA_con_pedidos.txt', encoding='utf-8-sig').read().split('\n')
loc = None; ped = {}; uu = {}
for l in txt:
    m = re.match(r'^(\S.*?)\s{2,}(\d+) u\b', l)
    if m and m.group(1) in pad: loc = m.group(1); uu[loc] = int(m.group(2))
    mp = re.search(r'\b25\s*-?\s*(\d{6})\s*$', l)
    if mp and loc: ped[loc] = f'25-{mp.group(1)}'
for loc, u in uu.items(): mov.append((loc, 'Refuerzo Libros Sorpresa', D(7, 10), ped.get(loc), u))
# 14/10
P14 = dict(zip(['ABASTO', 'ALTO COMAHUE', 'ALTO ROSARIO', 'ATENEO FLORIDA 340', 'ATENEO ROSARIO', 'AVELLANEDA', 'CABALLITO', 'NUEVO CENTRO SHOPPING', 'PASEO ALCORTA', 'PATIO OLMOS', 'PLAZA OESTE'], [f'25-{n}' for n in range(88251, 88262)]))
for f in glob.glob(str(H / '2026-10/reposicion_14-10/reposicion_*.csv')):
    l_ = inv[Path(f).stem[len('reposicion_'):]]; mov.append((l_, 'Repo Ranking (14/10)', D(14, 10), P14[l_], csv_u(f)))
# 15/10 repo + halloween
p15 = {}
for l in open(H / '2026-10/reposicion_15-10/PAECS_REPO_Y_HALLOWEEN_15-10_completo.txt', encoding='utf-8-sig'):
    m = re.match(r'\s*\d+\.\s+(\S.*?)\s{2,}\d+ u.*PAEC:\s*25\s*(\d{5})', l)
    if m: p15[m.group(1)] = f'25-{m.group(2)}'
    m2 = re.match(r'DEVOTO\s+25\s*(\d{5})', l.strip())
    if m2: p15['DEVOTO'] = f'25-{m2.group(1)}'
repo15 = {'SOLAR DE LA ABADIA', 'FLORES', 'MAR DEL PLATA PASEO ALDREY', 'MENDOZA PALMARES', 'SANTA FE Ribera Shopping'}
for f in glob.glob(str(H / '2026-10/reposicion_15-10/reposicion_*.csv')):
    l_ = inv[Path(f).stem[len('reposicion_'):]]
    mov.append((l_, 'Repo Ranking + Sugerido Halloween' if l_ in repo15 else 'Sugerido Halloween', D(15, 10), p15[l_], csv_u(f)))
# 19/10 recomposicion
r = {}
for row in csv.DictReader(open(H / '2026-10/recomposicion_19-10/detalle_recomposicion.csv', encoding='utf-8-sig'), delimiter=';'):
    r[row['local']] = r.get(row['local'], 0) + int(row['cant'])
for l_, u in r.items(): mov.append((l_, 'Recomposición Octubre', D(19, 10), None, u))
for f in glob.glob(str(H / '2026-10/reposicion_20-10/reposicion_*.csv')):
    mov.append((inv[Path(f).stem[len('reposicion_'):]], 'Repo Ranking (20/10)', D(20, 10), None, csv_u(f)))
