"""Refuerzo fiestas (pedido del 24/11/2026). UN SOLO PAEC, rótulo en el TXT.
Planilla REfu_navidad.xlsx: por título, una 'x' en L (completar a 10), M (a 20), N (a 5), o en O el texto 'Completar a 30/40'.
Van todas las librerías salvo Grand Splendid, Unicenter, Bahía Blanca Shopping, Ezeiza, Patio Bullrich, Patio Olmos,
Paseo Alcorta, Palermo Portal, Ushuaia y Trelew (seguimiento específico). Stock en locales y disponible: SII.
Lo que está en camino no se descuenta (criterio de Franco).
Uso: python -P fiestas.py PLANILLA.xlsx SII.xls config.json amba.txt SALIDA 'SII dd/mm hh:mm' 'ROTULO'"""
import sys, json, re, openpyxl, pandas as pd
from pathlib import Path
fpl, fsii, fcfg, famba, out, etq, rot = sys.argv[1:8]; out = Path(out); tag = out.name.upper()
NO_VA = {'ATENEO GRAND SPLENDID', 'UNICENTER', 'BAHIA BLANCA SHOPPING', 'EZEIZA', 'PATIO BULLRICH', 'PATIO OLMOS',
         'PASEO ALCORTA', 'PALERMO PORTAL', 'USHUAIA', 'TRELEW'}
pad = json.load(open(fcfg))['orden_paec'] + ['DEVOTO']
assert NO_VA <= set(pad), NO_VA - set(pad)
locs = [p for p in pad if p not in NO_VA]
AMBA = set(open(famba, encoding='utf-8').read().split('\n'))
ws = openpyxl.load_workbook(fpl).active
TIT = []
for r in range(4, ws.max_row + 1):
    isbn, idart = ws.cell(r, 1).value, ws.cell(r, 2).value
    m = {c: ws.cell(r, c).value for c in (12, 13, 14, 15)}
    if not any(v is not None for v in m.values()):
        continue
    assert isinstance(isbn, int) and isinstance(idart, int), (r, isbn)
    objs = [o for c, o in ((12, 10), (13, 20), (14, 5)) if m[c] is not None]
    if m[15] is not None:
        objs.append(int(re.search(r'(\d+)', str(m[15])).group(1)))
    assert len(objs) == 1, (r, m)
    assert not ws.row_dimensions[r].hidden, f'fila {r} oculta con marca'
    TIT.append((isbn, idart, objs[0], r))
assert len({i for i, *_ in TIT}) == len(TIT), 'ISBN repetido'
s = pd.read_excel(fsii); s = s[s['ID ARTICULO'].notna()].copy(); s['n'] = pd.to_numeric(s.ISBN, errors='coerce'); by = s.set_index('n')
assert all(p in s.columns for p in locs)
falta = [(i, r) for i, _, _, r in TIT if i not in by.index]; assert not falta, falta
mal = [(i, a, int(by.loc[i, 'ID ARTICULO'])) for i, a, _, _ in TIT if int(by.loc[i, 'ID ARTICULO']) != a]; assert not mal, mal
filas = []
for i, idart, obj, _ in TIT:
    for l in locs:
        st = by.loc[i, l]; st = 0 if pd.isna(st) else int(st)
        if obj - st > 0: filas.append((i, l, st, obj, obj - st))
rep = pd.DataFrame(filas, columns=['isbn', 'local', 'stock', 'objetivo', 'cant'])
dem = rep.groupby('isbn', sort=False).cant.sum()
T = lambda i: str(by.loc[i, 'TITULO']); ID = lambda i: int(by.loc[i, 'ID ARTICULO']); DISP = lambda i: int(by.loc[i, 'DISPONIBLE'])
OBJ = {i: o for i, _, o, _ in TIT}
orden = [i for i, *_ in TIT if i in dem.index]
corto = [(i, T(i), int(dem[i]), DISP(i)) for i in orden if DISP(i) < dem[i]]
(out / f'PAEC_{tag}.csv').write_bytes(('\r\n'.join(f'{i};{int(dem[i])}' for i in orden) + '\r\n').encode())
cab = [f'ROTULAR COMO: "{rot}"', f'REFUERZO FIESTAS - UN SOLO PAEC',
       'Completar a 5 / 10 / 20 / 30 / 40 según la planilla. No van: Grand Splendid, Unicenter, Bahía Blanca Shopping, Ezeiza, '
       'Patio Bullrich, Patio Olmos, Paseo Alcorta, Palermo Portal, Ushuaia y Trelew.',
       f'Stock en locales y disponible: {etq}. {len(orden)} títulos | {rep.local.nunique()} locales | {int(rep.cant.sum())} unidades.',
       ('Disponible: alcanza para todo.' if not corto else f'OJO: {len(corto)} títulos piden más que el disponible (marcados con *).'), '']
L = list(cab)
for k, i in enumerate(orden, 1):
    r = rep[rep.isbn == i].sort_values('local')
    L += ['=' * 72, f'{k}. {T(i)[:50]:<50}{int(dem[i]):>6} u' + ('  *' if DISP(i) < dem[i] else ''),
          f'   ISBN {i}   ID {ID(i)}   a {OBJ[i]}   en {len(r)} locales   disp {DISP(i)}', '=' * 72]
    L += [f'   {c:>4}  {l:<40} (tiene {st} -> {o})' for l, st, o, c in zip(r.local, r.stock, r.objetivo, r.cant)] + ['']
L += ['=' * 72, f'TOTAL {int(rep.cant.sum())} u']
(out / f'DESGLOSE_x_TITULO_{tag}.txt').write_text('\n'.join(L) + '\n', encoding='utf-8-sig')
M = list(cab); M[1] += ' - POR SUCURSAL (primero AMBA, después interior)'
for zona in ('AMBA', 'INTERIOR'):
    zl = sorted(l for l in rep.local.unique() if (l in AMBA) == (zona == 'AMBA'))
    M += ['#' * 72, f'# {zona}: {len(zl)} locales, {int(rep[rep.local.isin(zl)].cant.sum())} u', '#' * 72, '']
    for l in zl:
        sub = rep[rep.local == l].set_index('isbn'); sub = sub.loc[[i for i in orden if i in sub.index]]
        M += ['=' * 72, f'{l:<58}{int(sub.cant.sum()):>6} u', '=' * 72]
        M += [f'   {int(x.cant):>4}  {i}  {T(i)[:40]:<40} (tiene {int(x.stock)} -> {int(x.objetivo)})' for i, x in sub.iterrows()] + ['']
M += ['=' * 72, f'TOTAL {int(rep.cant.sum())} u']
(out / f'DESGLOSE_x_SUCURSAL_{tag}.txt').write_text('\n'.join(M) + '\n', encoding='utf-8-sig')
u = rep.groupby('local').cant.sum()
(out / f'UNIDADES_x_SUCURSAL_{tag}.txt').write_text('\n'.join(f'{l}\t{int(u[l])}' for l in pad if l in u.index) + '\n', encoding='utf-8-sig')
rep.to_csv(out / f'detalle_{tag.lower()}.csv', sep=';', index=False, encoding='utf-8-sig')
print('titulos marcados', len(TIT), '| con pedido', len(orden), '| sin pedido (todos completos):', [T(i) for i, *_ in TIT if i not in dem.index])
print('TOTAL', int(rep.cant.sum()), 'u |', rep.local.nunique(), 'de', len(locs), 'locales | AMBA', int(rep[rep.local.isin(AMBA)].cant.sum()),
      'interior', int(rep[~rep.local.isin(AMBA)].cant.sum()))
print('por objetivo', {o: int(rep[rep.objetivo == o].cant.sum()) for o in sorted(rep.objetivo.unique())})
print('CORTOS', len(corto), 'faltan', sum(a - b for _, _, a, b in corto))
for i, t, a, b in sorted(corto, key=lambda x: x[3] - x[2]): print(f'   {t[:45]:<45} pide {a:>4} disp {b:>5}')
