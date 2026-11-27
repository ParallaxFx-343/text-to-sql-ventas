"""27/11: el Detalle nuevo vino solo desde el 01/11. Octubre sale de la hoja Datos de la planilla del 20/11
(Nro, Local, Fecha, Artículo, ISBN, Título corto, Cantidad, Total $), pasado al formato de normalizar_ventas.py.
Control: 01-19/11 de la planilla del 20/11 = export nuevo (Nro x Fecha x Artículo, 563 celdas, 632 u, 652 filas).
Uso: python -P armar_detalle.py PLANILLA_ANTERIOR.xlsx detalle_sorpresa_nuevo.xlsx DESDE(aaaa-mm-dd) SALIDA.xlsx"""
import sys, pandas as pd
fp, fn, desde, out = sys.argv[1:5]
p = pd.read_excel(fp, sheet_name='Datos', dtype={'ISBN': str}); p['Fecha'] = pd.to_datetime(p.Fecha)
n = pd.read_excel(fn, dtype={'ISBN': str}); n['Fecha'] = pd.to_datetime(n.Fecha)
assert n.Fecha.min() >= pd.Timestamp(desde)
nro = n.Sucursal.str.extract(r'\((\d+)\)\s*$')[0].astype(int)
corte = p.Fecha.max()                                  # el último día de la planilla anterior era parcial
a = p[(p.Fecha >= desde) & (p.Fecha < corte)].groupby(['Nro', 'Fecha', 'Artículo']).Cantidad.sum()
b = n[n.Fecha < corte].assign(Nro=nro).groupby(['Nro', 'Fecha', 'Articulo']).Cantidad.sum(); b.index.names = a.index.names
c = pd.concat([a, b], axis=1).fillna(0); assert (c.iloc[:, 0] == c.iloc[:, 1]).all(), c[c.iloc[:, 0] != c.iloc[:, 1]]
tit = dict(zip(n.Articulo, n.Titulo)); o = p[p.Fecha < desde]; assert set(o['Artículo']) <= set(tit)
r = pd.DataFrame({'Sucursal': o.Local + '  (' + o.Nro.astype(str) + ')', 'Fecha': o.Fecha, 'Articulo': o['Artículo'],
                  'Titulo': o['Artículo'].map(tit), 'Cantidad': o.Cantidad, 'Bruto': o['Total $'], 'Total': o['Total $'], 'ISBN': o.ISBN})
m = pd.concat([r, n]).sort_values(['Fecha', 'Sucursal'], kind='stable'); m.to_excel(out, index=False)
print(f'control {len(c)} celdas iguales ({int(a.sum())} u) | anterior {int(r.Cantidad.sum())} u + nuevo {int(n.Cantidad.sum())} u = {int(m.Cantidad.sum())} u')
