"""LP de diciembre 2026 rehecha el 27/11 por la suba por el dólar (~10%).
Base: LP_DICIEMBRE_26_PARA_SUBIR_v20-11_REEMPLAZADA.xlsx (la del 20/11: ID e Importe, mismo orden).
Se aplica CAMBIO_DE_PRECIOS_DICIEMBRE_2026 con las correcciones que confirmó Franco (27/11):
  - Amarantia 747595, Aventureras 753349, Resetea ed. aniversario 705123: el CAMBIO dice 3.800 -> 38.000 (faltaba un cero).
  - Voy a compartir TD 704437: el CAMBIO dice 29.200 (precio de la rústica) -> 34.000 como el resto de los "Voy a" TD.
  - Vacas TD viejas (649061, 679580, 689236) a 35.000 como vienen (se equiparan a las actuales).
  - 7 novedades de diciembre: precio de la LP Argentina (+10%); Cuaderno del despeinado 754497 queda en 34.500.
  - El espejo opaco 678137 baja a 16.000 (distribución), como la LP Argentina.
Controles: mismo orden, sin IDs repetidos, ningún saldado (<2.000) cambia, nada queda por debajo de 300 salvo los que ya estaban,
LP Argentina = resultado salvo las 4 correcciones.
Uso: python -P lp_diciembre_v2.py BASE.xlsx CAMBIO.xlsx LP_ARGENTINA.xlsx SII.xls SALIDA_DIR"""
import sys, pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
fb, fc, fa, fs, out = sys.argv[1:]; out = Path(out)
b = pd.read_excel(fb, header=None, names=['id', 'p'])
rd = lambda f: (lambda d: d[d['I.D.'].notna()].astype({'I.D.': int}))(pd.read_excel(f, header=2))
c = rd(fc); a = rd(fa)
assert not b.id.duplicated().any() and not c['I.D.'].duplicated().any() and not a['I.D.'].duplicated().any()
assert c['I.D.'].isin(b.id).all() and a['I.D.'].isin(b.id).all()
bp = dict(zip(b.id, b.p)); ap = dict(zip(a['I.D.'], a['P.V.P. $']))
CORRIGE = {747595: 38000, 753349: 38000, 705123: 38000, 704437: 34000}
NOVEDADES = [754498, 754499, 754500, 754501, 754502, 754503, 754504]   # precio LP Argentina
QUEDA = {754497: 34500}                                                 # Cuaderno del despeinado
EXTRA = {678137: 16000}                                                 # El espejo opaco (distribución)
for i, v in CORRIGE.items(): assert i in set(c['I.D.']) and ap[i] != v, i
nuevo = dict(zip(c['I.D.'], c['P.V.P. $'])); nuevo.update(CORRIGE)
nuevo.update({i: ap[i] for i in NOVEDADES}); nuevo.update(EXTRA)
P = {i: nuevo.get(i, p) for i, p in bp.items()}
assert P[754497] == 34500 and P[734075] == 0.05
lp = [(int(i), P[i]) for i in b.id]
# controles
sald = [i for i, p in bp.items() if p < 2000 and P[i] != p]; assert not sald, sald
bajo = [i for i in P if P[i] < 300 and bp[i] >= 300]; assert not bajo, bajo
baja = [i for i in P if P[i] < bp[i] and i not in EXTRA]; assert not baja, baja
dif = {i for i in a['I.D.'] if P[i] != ap[i]}; assert dif == set(CORRIGE), dif
ch = {i for i in P if P[i] != bp[i]}
pct = pd.Series({i: P[i] / bp[i] - 1 for i in ch})
fuera = pct[(pct < .08) | (pct > .12)]
wb = Workbook(); ws = wb.active; ws.title = 'Sheet'
for i, p in lp: ws.append([i, int(p) if float(p).is_integer() else float(p)])
wb.save(out / 'LP_DICIEMBRE_26_PARA_SUBIR.xlsx')
# control
s = pd.read_excel(fs); s = s[s['ID ARTICULO'].notna()]; tit = dict(zip(s['ID ARTICULO'].astype(int), s.TITULO))
tit.update(dict(zip(a['I.D.'], a['T i t u l o'])))
motivo = {**{i: 'CORREGIDO (Franco 27/11): el CAMBIO decía ' + f'{int(ap[i]):,}'.replace(',', '.') for i in CORRIGE},
          **{i: 'Novedad dic: precio LP Argentina' for i in NOVEDADES}, **{i: 'Venta en firme / distribución: baja a 16.000 (Franco 27/11)' for i in EXTRA}}
tipo = dict(zip(c['I.D.'], c['Novedad/Cambio de precio']))
rows = [(i, tit.get(i, ''), bp[i], P[i], P[i] - bp[i], round(pct[i] * 100, 1), motivo.get(i, tipo.get(i, ''))) for i in b.id if i in ch]
ctl = pd.DataFrame(rows, columns=['ID', 'Título', 'Precio 20/11', 'Precio nuevo', 'Diferencia', '% suba', 'Motivo'])
res = pd.DataFrame([
    ('Artículos en la LP', len(lp)), ('Precios que cambian', len(ch)), ('  del CAMBIO (con 4 corregidos)', len(set(c['I.D.']) & ch)),
    ('  novedades dic (LP Argentina)', len(NOVEDADES)), ('  El espejo opaco (baja)', 1),
    ('Suba promedio simple (sin El espejo opaco) %', round(pct.drop(list(EXTRA)).mean() * 100, 2)),
    ('Saldados (< $2.000) que cambian', 0), ('Precios que quedan < $300 y antes no', 0),
    ('LP Argentina distinta del resultado', f'{len(dif)} (las 4 corregidas)'),
    ('Cambios fuera de +8% / +12%', len(fuera))], columns=['Control', 'Valor'])
quieto = b[(b.p >= 2000) & (~b.id.isin(list(ch)))].copy(); quieto['Título'] = quieto.id.map(tit)
quieto['Disp. SII 26/11'] = quieto.id.map(dict(zip(s['ID ARTICULO'].astype(int), s.DISPONIBLE)))
quieto['Marca LP Arg'] = quieto.id.map(dict(zip(a['I.D.'], a['Novedad/Cambio de precio']))).fillna('')
quieto = quieto.rename(columns={'id': 'ID', 'p': 'Precio (sin cambio)'})[['ID', 'Título', 'Precio (sin cambio)', 'Disp. SII 26/11', 'Marca LP Arg']].sort_values('Disp. SII 26/11', ascending=False)
with pd.ExcelWriter(out / 'LP_DICIEMBRE_26_CONTROL.xlsx') as w:
    res.to_excel(w, sheet_name='Resumen', index=False); ctl.to_excel(w, sheet_name='Cambios', index=False); quieto.to_excel(w, sheet_name='No saldados sin cambio', index=False)
    for ws in w.book.worksheets:
        for col, wd in zip('ABCDEFG', (14, 48, 14, 14, 12, 10, 52)): ws.column_dimensions[col].width = wd
        for cell in ws[1]: cell.font = Font(bold=True)
    wc = w.book['Cambios']
    for r in range(2, wc.max_row + 1):
        if wc.cell(r, 1).value in motivo:
            for k in range(1, 8): wc.cell(r, k).fill = PatternFill('solid', fgColor='FFF2CC')
print(res.to_string(index=False)); print('fuera de rango:'); print(ctl[ctl.ID.isin(fuera.index)].to_string(index=False))
