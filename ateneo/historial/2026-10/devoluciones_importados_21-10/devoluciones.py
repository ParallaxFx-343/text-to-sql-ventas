"""Candidatos a devolución de Gatito, Dino, Stickers Dinosaurios, Stickers Granja y Banderas del mundo
a partir de la Sabana consolidada de Importados. Uso: python -P devoluciones.py SALIDA.xlsx SABANA.xlsx [SABANA2.xlsx ...]
(la de importados y la de BANDERAS DEL MUNDO, que es nacional y viene aparte)
Agrupa por cuenta + sucursal (cada punto de entrega es una devolución distinta).
Criterio: Prioridad alta = stock >= 20 y liquidó menos del 25 % de lo entregado;
Pedir = stock >= 10 y < 25 %; Revisar = stock >= 10 sin ninguna liquidación (venta 0 puede ser falta de
liquidar) o con 25-50 %; el resto, No."""
import sys, pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
fout, fins = sys.argv[1], sys.argv[2:]
T = {747643: 'Gatito', 747645: 'Dino', 747646: 'Stickers Dinosaurios', 747647: 'Stickers Granja', 744138: 'Banderas del mundo'}
df = pd.concat([pd.read_excel(f, sheet_name=0).dropna(subset=['CODIGO']) for f in fins], ignore_index=True)
for c in ['CUENTA', 'CODIGO', 'ISBN', 'ENTREGADA', 'VENTA', 'DEVOL', 'AJUSTE', 'TRANSF', 'SALDO']:
    df[c] = df[c].astype('int64')
s = df[df.CODIGO.isin(T)].copy()
assert (s.ENTREGADA - s.VENTA - s.DEVOL + s.AJUSTE + s.TRANSF == s.SALDO).all()
s['Título'] = s.CODIGO.map(T)
s['Últ. liq.'] = pd.to_datetime(s.FECHAULTLIQ, unit='D', origin='1899-12-30')
k = ['CUENTA', 'CLIENTE', 'SUCURSAL', 'NOMBRE_VENDEDOR']
g = s.groupby(k).agg(Entregada=('ENTREGADA', 'sum'), Liquidada=('VENTA', 'sum'), Devuelta=('DEVOL', 'sum'),
                     Stock=('SALDO', 'sum'), liq=('Últ. liq.', 'max')).reset_index()
pt = s.pivot_table(index=k, columns='Título', values='SALDO', aggfunc='sum', fill_value=0)
pt = pt[[t for t in T.values() if t in pt.columns]].add_prefix('Stock ').reset_index()
g = g.merge(pt, on=k)
g['% liquidado'] = g.Liquidada / g.Entregada
n_suc = df.groupby('CUENTA').SUCURSAL.nunique()
g['Sucursales de la cuenta'] = g.CUENTA.map(n_suc)
def sug(r):
    if r.Stock >= 10 and pd.isna(r.liq): return 'Revisar (nunca liquidó)'
    if r.Stock >= 20 and r['% liquidado'] < .25: return 'Prioridad alta'
    if r.Stock >= 10 and r['% liquidado'] < .25: return 'Pedir'
    if r.Stock >= 10 and r['% liquidado'] < .5: return 'Revisar'
    return 'No'
g['Sugerencia'] = g.apply(sug, axis=1)
orden = {'Prioridad alta': 0, 'Pedir': 1, 'Revisar (nunca liquidó)': 2, 'Revisar': 3, 'No': 4}
g = g.sort_values(['Sugerencia', 'Stock'], key=lambda c: c.map(orden) if c.name == 'Sugerencia' else -c)
g = g.rename(columns={'CUENTA': 'Cuenta', 'CLIENTE': 'Cliente', 'SUCURSAL': 'Sucursal', 'NOMBRE_VENDEDOR': 'Vendedor', 'liq': 'Últ. liq.'})
cols = ['Sugerencia', 'Cuenta', 'Cliente', 'Sucursal', 'Sucursales de la cuenta', 'Vendedor', 'Últ. liq.',
        'Entregada', 'Liquidada', 'Devuelta', 'Stock', '% liquidado'] + [c for c in g.columns if c.startswith('Stock ')]
g = g[cols]
det = s.rename(columns={'CUENTA': 'Cuenta', 'CLIENTE': 'Cliente', 'SUCURSAL': 'Sucursal', 'NOMBRE_VENDEDOR': 'Vendedor',
                        'ENTREGADA': 'Entregada', 'VENTA': 'Liquidada', 'DEVOL': 'Devuelta', 'SALDO': 'Stock'})
det['% liquidado'] = det.Liquidada / det.Entregada
det = det[['Cuenta', 'Cliente', 'Sucursal', 'Vendedor', 'Título', 'CODIGO', 'ISBN', 'Últ. liq.', 'Entregada', 'Liquidada',
           'Devuelta', 'Stock', '% liquidado']].rename(columns={'CODIGO': 'ID'}).sort_values(['Stock'], ascending=False)
tot = s.groupby(['CODIGO', 'ISBN', 'Título']).agg(Clientes=('CUENTA', 'nunique'), Entregada=('ENTREGADA', 'sum'),
      Liquidada=('VENTA', 'sum'), Devuelta=('DEVOL', 'sum'), Stock=('SALDO', 'sum')).reset_index().rename(columns={'CODIGO': 'ID'})
tot['% liquidado'] = tot.Liquidada / tot.Entregada
res = g.groupby('Sugerencia', sort=False).agg(Puntos=('Cuenta', 'size'), Stock=('Stock', 'sum')).reset_index()
with pd.ExcelWriter(fout, engine='openpyxl') as w:
    g.to_excel(w, sheet_name='Por cliente', index=False)
    det.to_excel(w, sheet_name='Por cliente y título', index=False)
    tot.to_excel(w, sheet_name='Totales por título', index=False)
    res.to_excel(w, sheet_name='Totales por título', index=False, startrow=len(tot) + 3)
wb = load_workbook(fout)
col = {'Prioridad alta': 'F8CBAD', 'Pedir': 'FCE4D6', 'Revisar (nunca liquidó)': 'FFF2CC', 'Revisar': 'FFF2CC'}
for ws in wb:
    for c in ws[1]: c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.freeze_panes = 'A2'
    for i, h in enumerate([c.value for c in ws[1]], 1):
        L = ws.cell(1, i).column_letter
        ws.column_dimensions[L].width = {'Cliente': 40, 'Sucursal': 32, 'Vendedor': 22, 'Sugerencia': 22, 'Título': 22}.get(h, 12)
        for r in range(2, ws.max_row + 1):
            c = ws.cell(r, i)
            if h == '% liquidado': c.number_format = '0%'
            elif h == 'Últ. liq.': c.number_format = 'DD/MM/YYYY'
            elif h in ('ISBN',): c.number_format = '0'
ws = wb['Por cliente']
for r in range(2, ws.max_row + 1):
    f = col.get(ws.cell(r, 1).value)
    if f:
        for c in range(1, ws.max_column + 1): ws.cell(r, c).fill = PatternFill('solid', fgColor=f)
ws.auto_filter.ref = ws.dimensions
wb['Por cliente y título'].auto_filter.ref = wb['Por cliente y título'].dimensions
wb.save(fout)
print(tot.to_string()); print(res.to_string())
print(g[g.Sugerencia.isin(['Prioridad alta'])][['Cliente', 'Sucursal', 'Sucursales de la cuenta', 'Vendedor', 'Entregada', 'Liquidada', 'Stock', '% liquidado']].to_string())
