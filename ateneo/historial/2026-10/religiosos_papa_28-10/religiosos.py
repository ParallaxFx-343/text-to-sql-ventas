"""Exhibición de libros religiosos por la visita del Papa (noviembre 2026), pedido del 28/10. UN SOLO PAEC.
Completar a 10 en todos los locales; a 5 en Patio Bullrich, Bahía Blanca Shopping, Trelew, Ezeiza y Palermo Portal;
Ateneo Juramento no va. Stock en locales y disponible: SII. Lo que está en camino (repo 27/10) no se descuenta.
Uso: python -P religiosos.py SII.xls config.json amba.txt SALIDA 'SII dd/mm hh:mm' [REPO_EN_CAMINO_DIR]"""
import sys, json, glob, pandas as pd
from pathlib import Path
fsii, fcfg, famba, out, etq = sys.argv[1:6]; out = Path(out)
TIT = [(9789500218078, 748667), (9789500216685, 734151), (9789500210560, 675902), (9789500210577, 675903), (9789500216678, 734152)]
CHICOS = {'PATIO BULLRICH', 'BAHIA BLANCA SHOPPING', 'TRELEW', 'EZEIZA', 'PALERMO PORTAL'}
NO_VA = {'ATENEO JURAMENTO'}
pad = json.load(open(fcfg))['orden_paec'] + ['DEVOTO']
locs = [p for p in pad if p not in NO_VA]
AMBA = set(open(famba, encoding='utf-8').read().split('\n'))
assert CHICOS <= set(locs)
s = pd.read_excel(fsii); s = s[s['ID ARTICULO'].notna()].copy(); s['n'] = pd.to_numeric(s.ISBN, errors='coerce'); by = s.set_index('n')
assert all(p in s.columns for p in locs)
camino = {}
if len(sys.argv) > 6:
    for f in glob.glob(sys.argv[6] + '/reposicion_*.csv'):
        for l in open(f, 'rb').read().decode().split('\r\n'):
            if l: i, q = map(int, l.split(';')); camino[i] = camino.get(i, 0) + q
filas = []
for i, idart in TIT:
    assert int(by.loc[i, 'ID ARTICULO']) == idart, (i, idart)
    for l in locs:
        st = by.loc[i, l]; st = 0 if pd.isna(st) else int(st); obj = 5 if l in CHICOS else 10
        if obj - st > 0: filas.append((i, l, st, obj, obj - st))
rep = pd.DataFrame(filas, columns=['isbn', 'local', 'stock', 'objetivo', 'cant'])
dem = rep.groupby('isbn', sort=False).cant.sum()
T = lambda i: str(by.loc[i, 'TITULO']); ID = lambda i: int(by.loc[i, 'ID ARTICULO']); DISP = lambda i: int(by.loc[i, 'DISPONIBLE'])
corto = [(T(i), int(dem[i]), DISP(i)) for i, _ in TIT if DISP(i) < dem.get(i, 0)]
orden = [i for i, _ in TIT if i in dem.index]
(out / 'PAEC_RELIGIOSOS_PAPA_28-10.csv').write_bytes(('\r\n'.join(f'{i};{int(dem[i])}' for i in orden) + '\r\n').encode())
cab = ['EXHIBICION LIBROS RELIGIOSOS - VISITA DEL PAPA (NOVIEMBRE 2026) - UN SOLO PAEC',
       'Completar a 10 en todos los locales; a 5 en Patio Bullrich, Bahía Blanca Shopping, Trelew, Ezeiza y Palermo Portal. Ateneo Juramento no va.',
       f'Stock en locales y disponible: {etq}. {len(orden)} títulos | {rep.local.nunique()} locales | {int(rep.cant.sum())} unidades.',
       ('Disponible: alcanza para todo.' if not corto else 'OJO, disponible que no alcanza: ' + '; '.join(f'{t} pide {a}, hay {b}' for t, a, b in corto)), '']
L = list(cab)
for k, i in enumerate(orden, 1):
    r = rep[rep.isbn == i].sort_values('local')
    L += ['=' * 72, f'{k}. {T(i)[:52]:<52}{int(dem[i]):>6} u', f'   ISBN {i}   ID {ID(i)}   en {len(r)} locales   disp {DISP(i)}', '=' * 72]
    L += [f'   {c:>4}  {l:<40} (tiene {st} -> {o})' for l, st, o, c in zip(r.local, r.stock, r.objetivo, r.cant)] + ['']
L += ['=' * 72, f'TOTAL {int(rep.cant.sum())} u']
(out / 'DESGLOSE_x_TITULO_RELIGIOSOS_PAPA_28-10.txt').write_text('\n'.join(L) + '\n', encoding='utf-8-sig')
M = list(cab); M[0] += ' - POR SUCURSAL (primero AMBA, después interior)'
for zona in ('AMBA', 'INTERIOR'):
    zl = sorted(l for l in rep.local.unique() if (l in AMBA) == (zona == 'AMBA'))
    M += ['#' * 72, f'# {zona}: {len(zl)} locales, {int(rep[rep.local.isin(zl)].cant.sum())} u', '#' * 72, '']
    for l in zl:
        r = rep[rep.local == l].set_index('isbn').loc[[i for i in orden if i in set(rep[rep.local == l].isbn)]]
        M += ['=' * 72, f'{l:<58}{int(r.cant.sum()):>6} u' + ('   (local chico: a 5)' if l in CHICOS else ''), '=' * 72]
        M += [f'   {int(x.cant):>4}  {i}  {T(i)[:40]:<40} (tiene {int(x.stock)})' for i, x in r.iterrows()] + ['']
M += ['=' * 72, f'TOTAL {int(rep.cant.sum())} u']
(out / 'DESGLOSE_x_SUCURSAL_RELIGIOSOS_PAPA_28-10.txt').write_text('\n'.join(M) + '\n', encoding='utf-8-sig')
rep.to_csv(out / 'detalle_religiosos.csv', sep=';', index=False, encoding='utf-8-sig')
for i in orden:
    print(f'{T(i)[:40]:<40} {int(dem[i]):>5} u  {int((rep.isbn == i).sum()):>2} locales  disp {DISP(i)}  en camino repo 27/10: {camino.get(i, 0)}')
print('TOTAL', int(rep.cant.sum()), 'u |', rep.local.nunique(), 'locales | sin pedido:', sorted(set(locs) - set(rep.local)), '| corto', corto)
print('AMBA', int(rep[rep.local.isin(AMBA)].cant.sum()), 'interior', int(rep[~rep.local.isin(AMBA)].cant.sum()))
