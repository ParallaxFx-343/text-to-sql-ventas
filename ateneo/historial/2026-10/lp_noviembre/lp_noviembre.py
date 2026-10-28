"""LP de noviembre 2026 para subir (hoja 'Sheet', ID e Importe sin encabezado, desde la fila 1).
1. Base: LP_OCTUBRE_26_CON_NOVEDADES (se mantiene el orden de filas).
2. Se aplican los precios de CAMBIO_DE_PRECIOS_NOVIEMBRE_2026 (cambios de precio, promo Nov-Dic, venta en firme).
3. Lo que está en LP_ARGENTINA_NOVIEMBRE_2026 y falta en la base se agrega al final.
4. Controles: mismo orden que la base, sin IDs repetidos, solo cambian los precios del CAMBIO, LP Argentina = resultado.
Uso: python -P lp_noviembre.py BASE.xlsx CAMBIO.xlsx LP_ARGENTINA.xlsx SALIDA_DIR"""
import sys, pandas as pd
from pathlib import Path
from openpyxl import Workbook
fb, fc, fa, out = sys.argv[1:]; out = Path(out)
b = pd.read_excel(fb, dtype={'Código / ISBN': str})
def rd(f):
    d = pd.read_excel(f, header=2); d = d[d['I.D.'].notna()].copy(); d['I.D.'] = d['I.D.'].astype(int)
    d['P.V.P. $'] = d['P.V.P. $'].round(2); return d
c = rd(fc); a = rd(fa).drop_duplicates('I.D.')        # LP Argentina trae 4 artículos repetidos (iguales)
assert not b['Nro. Art.'].duplicated().any() and not c['I.D.'].duplicated().any()
assert c['I.D.'].isin(b['Nro. Art.']).all(), 'hay cambios de artículos que no están en la base'
nuevo = dict(zip(c['I.D.'], c['P.V.P. $']))
lp = [(int(i), nuevo.get(int(i), p)) for i, p in zip(b['Nro. Art.'], b.Importe)]
agreg = a[~a['I.D.'].isin(b['Nro. Art.'])]
lp += [(int(r['I.D.']), r['P.V.P. $']) for _, r in agreg.iterrows()]
ids = [i for i, _ in lp]; P = dict(lp)
assert len(ids) == len(set(ids)) and ids[:len(b)] == list(b['Nro. Art.'])
cambia = {i for i, p in zip(b['Nro. Art.'], b.Importe) if P[i] != p}
assert cambia <= set(c['I.D.'])
dif = [(i, p, P[i]) for i, p in zip(a['I.D.'], a['P.V.P. $']) if P[i] != p]
assert not dif, dif
assert P[734075] == 0.05
wb = Workbook(); ws = wb.active; ws.title = 'Sheet'
for i, p in lp:
    ws.append([i, int(p) if float(p).is_integer() else float(p)])
wb.save(out / 'LP_NOVIEMBRE_26_PARA_SUBIR.xlsx')
# control para Franco
bp = dict(zip(b['Nro. Art.'], b.Importe)); bt = dict(zip(b['Nro. Art.'], b['Título']))
ctl = c[['I.S.B.N.', 'I.D.', 'T i t u l o', 'Novedad/Cambio de precio', 'P.V.P. $']].copy()
ctl.insert(4, 'Precio base', ctl['I.D.'].map(bp)); ctl['Dif'] = ctl['P.V.P. $'] - ctl['Precio base']
ctl.columns = ['ISBN', 'ID', 'Título', 'Tipo', 'Precio base (oct)', 'Precio noviembre', 'Diferencia']
ag = agreg[['I.S.B.N.', 'I.D.', 'T i t u l o', 'P.V.P. $', 'STKDISP']]; ag.columns = ['ISBN', 'ID', 'Título', 'Precio', 'Stock disp. LP Arg']
with pd.ExcelWriter(out / 'LP_NOVIEMBRE_26_CONTROL.xlsx') as w:
    ctl.to_excel(w, sheet_name='Cambios aplicados', index=False); ag.to_excel(w, sheet_name='Agregados al final', index=False)
print('base', len(b), '| cambios en archivo', len(c), '| precios que cambian', len(cambia), '| ya estaban al precio nuevo', len(c) - len(cambia),
      '| agregados', len(agreg), '| total', len(lp))
print(ctl.groupby('Tipo').agg(n=('ID', 'size'), sube=('Diferencia', lambda s: (s > 0).sum()), baja=('Diferencia', lambda s: (s < 0).sum()), igual=('Diferencia', lambda s: (s == 0).sum())))
print(ag.to_string())
