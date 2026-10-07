"""Repo de Libros Sorpresa para AMBA: titulos con menos de 3 dias de stock.

Reglas de Franco (07/10/2026):
- se mide por titulo en cada local: dias = stock del titulo (SII) / ritmo
  ritmo = vendido del titulo del primer dia de venta del local al ultimo dia completo,
  dividido por esos dias
- entra si dias < 3 (estricto); se manda lo vendido de ese titulo en ese periodo
- AMBA = CABA + conurbano + Ateneo La Plata + Calle 12 + San Miguel (La Plata no)
- Ateneo Grand Splendid (nunca recibio): 8 de cada uno
- un solo PAEC (EAN;total) + TXT por titulo con locales en orden alfabetico
Uso: python -P repo_sorpresa.py DETALLE.xlsx SII.xlsx SALIDA_DIR
"""
import math, sys
from pathlib import Path
import pandas as pd

f_det, f_sii, out = sys.argv[1], sys.argv[2], Path(sys.argv[3])
AMBA = ['ABASTO', 'AEROPARQUE', 'ALTO PALERMO', 'ATENEO FLORIDA 340', 'ATENEO FLORIDA 632',
        'ATENEO GRAND SPLENDID', 'ATENEO JURAMENTO', 'BORGES', 'CABALLITO', 'DOT BAIRES SHOPPING',
        'FLORES', 'PALERMO PORTAL', 'PASEO ALCORTA', 'PATIO BULLRICH', 'SOLAR DE LA ABADIA',
        'VILLA CRESPO', 'VILLA URQUIZA', 'DEVOTO',
        'AVELLANEDA', 'EZEIZA', 'LOMAS DE ZAMORA', 'PALMAS DEL PILAR', 'PLAZA OESTE', 'QUILMES CENTRO',
        'REMEROS (NORDELTA)', 'SAN JUSTO', 'SAN MIGUEL', 'TORTUGAS OPEN MALL', 'UNICENTER',
        'ATENEO LA PLATA', 'CALLE 12']
CARGA_INICIAL = {'ATENEO GRAND SPLENDID': 8}
CORTE = 3

det = pd.read_excel(f_det, dtype={'ISBN': str})
det['local'] = det.Sucursal.str.replace(r'\s*\(\d+\)\s*$', '', regex=True).str.strip()
det['dia'] = det.Fecha.dt.date
dias = sorted(det.dia.unique())
ult_c = dias[-2]                                   # el ultimo dia del export es parcial
d = det[det.dia <= ult_c]
primera = det[det.Cantidad > 0].groupby('local').dia.min()

sii = pd.read_excel(f_sii)
cols = list(sii.columns)
suc = cols[cols.index('CLASIFIC') + 1:]
assert all(a in suc for a in AMBA), [a for a in AMBA if a not in suc]
s = sii[sii.TITULO.astype(str).str.startswith('LIBRO SORPRESA')].copy()
assert len(s) == 6
info = {int(r['ID ARTICULO']): (str(int(r.ISBN)), r.TITULO.replace('LIBRO SORPRESA ', '').strip())
        for _, r in s.iterrows()}

lineas = []                                         # (local, id, cantidad, stock, vendido, ritmo, dias)
for loc in AMBA:
    for _, r in s.iterrows():
        art = int(r['ID ARTICULO']); stock = int(0 if pd.isna(r[loc]) else r[loc])
        if loc in CARGA_INICIAL:
            lineas.append((loc, art, CARGA_INICIAL[loc] - stock, stock, 0, None, None)); continue
        if loc not in primera.index:
            continue                                # sin ventas: no hay ritmo
        n = (ult_c - primera[loc]).days + 1
        vend = int(d[(d.local == loc) & (d.Articulo == art)].Cantidad.sum())
        if vend <= 0:
            continue
        ritmo = vend / n
        cob = stock / ritmo
        if cob < CORTE:
            lineas.append((loc, art, vend, stock, vend, round(ritmo, 2), round(cob, 1)))
rep = pd.DataFrame(lineas, columns=['local', 'id', 'cant', 'stock', 'vendido', 'ritmo', 'dias'])
rep = rep[rep.cant > 0]
tot = rep.groupby('id').cant.sum().sort_values(ascending=False)

out.mkdir(parents=True, exist_ok=True)
(out / 'PAEC_REPO_SORPRESA_AMBA.csv').write_bytes(
    ('\r\n'.join(f'{info[i][0]};{int(q)}' for i, q in tot.items()) + '\r\n').encode())
L = ['REPO LIBROS SORPRESA - AMBA - DESGLOSE POR TITULO',
     f'Stock: SII 07/10 07:00. Venta: {dias[0]:%d/%m} al {ult_c:%d/%m} (el {dias[-1]:%d/%m} es parcial y no se usa).',
     f'Regla: por titulo en cada local, entra si el stock le alcanza para menos de {CORTE} dias',
     'al ritmo de venta del local desde su primera venta; se repone lo vendido en el periodo.',
     'ATENEO GRAND SPLENDID (nunca recibio): 8 de cada uno.',
     f'{len(tot)} titulos | {rep.local.nunique()} locales | {int(rep.cant.sum())} unidades', '']
for k, (i, q) in enumerate(tot.items(), 1):
    b = rep[rep.id == i].sort_values('local')
    L += ['=' * 66, f'{k}. LIBRO SORPRESA {info[i][1]:<34}{int(q):>6} u',
          f'   EAN {info[i][0]}   ID {i}   en {len(b)} locales', '=' * 66]
    L += [f'   {int(r.cant):>4}  {r.local}' for r in b.itertuples()] + ['']
L += ['=' * 66, f'{"TOTAL":<52}{int(rep.cant.sum()):>6} u']
(out / 'DESGLOSE_REPO_SORPRESA_AMBA.txt').write_text('\n'.join(L) + '\n', encoding='utf-8-sig')
rep.assign(ean=rep.id.map(lambda i: info[i][0]), titulo=rep.id.map(lambda i: info[i][1])) \
   .to_csv(out / 'detalle_repo.csv', sep=';', index=False, encoding='utf-8-sig')
print(f'{len(tot)} titulos | {rep.local.nunique()} locales | {len(rep)} lineas | {int(rep.cant.sum())} u')
print({info[i][1]: int(q) for i, q in tot.items()})
print(rep.groupby('local').cant.sum().sort_values(ascending=False).to_dict())
