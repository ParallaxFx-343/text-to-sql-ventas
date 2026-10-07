"""Tablero de ventas de los Libros Sorpresa por local (solo locales fisicos).

Uso: python -P tablero_sorpresa.py DETALLE.xlsx AGRUPADAS.xlsx config.json SALIDA.xlsx
- DETALLE: "Detalle Venta Articulo" (una fila por venta: Sucursal "NOMBRE  (nro)",
  Fecha, Articulo, Titulo, Cantidad, Bruto, Total, ISBN)
- AGRUPADAS: "Ventas Proveedor Agrupadas Articulo" (solo para cuadrar y para
  saber que locales vendieron Editorial pero no Sorpresa)
- config.json: padron de locales fisicos (orden_paec)
"""
import json
import sys
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

f_det, f_agr, f_cfg, salida = sys.argv[1:5]
padron = set(json.load(open(f_cfg, encoding='utf-8'))['orden_paec']) | {'DEVOTO'}

# ---------- datos ----------
det = pd.read_excel(f_det, dtype={'ISBN': str})
det['nro'] = det.Sucursal.str.extract(r'\((\d+)\)\s*$')[0].astype(int)
det['local'] = det.Sucursal.str.replace(r'\s*\(\d+\)\s*$', '', regex=True).str.strip()
det = det[det.local.isin(padron)].copy()          # solo locales fisicos
NOMBRES = {'HISTORICA': 'Histórica', 'ROMANTICA': 'Romántica', 'CONTEMPORANEA': 'Contemporánea',
           'MISTERIO': 'Misterio', 'EROTICA': 'Erótica', 'BIO': 'Bío'}
det['titulo'] = det.Titulo.str.replace('LIBRO SORPRESA ', '', regex=False).str.strip().map(NOMBRES)
assert det.titulo.notna().all(), 'hay un Libro Sorpresa nuevo sin nombre en NOMBRES'
det['dia'] = det.Fecha.dt.date

agr = pd.read_excel(f_agr, dtype={'Código / ISBN': str})
agr = agr[agr.Canal.isna() | (agr.Canal.astype(str).str.strip() == '')]   # sin canales online
agr = agr[agr['Nombre Sucursal'].isin(padron)]
sor = agr[agr.TITULO.astype(str).str.contains('LIBRO SORPRESA')]

# cuadre por dos caminos: detalle vs agrupadas, local x articulo
a = sor.groupby(['Código Sucursal', 'ID ARTICULO']).Cantidad.sum()
d = det.groupby(['nro', 'Articulo']).Cantidad.sum()
a.index.names = d.index.names = ['s', 'a']
cuadre = pd.concat([a, d], axis=1).fillna(0)
assert (cuadre.iloc[:, 0] == cuadre.iloc[:, 1]).all(), 'detalle y agrupadas no cuadran'

dias = sorted(det.dia.unique())
ult = dias[-1]                                      # dia parcial (export a la mañana)
hasta = max(det.Fecha)                              # para el subtitulo
DIAS_ES = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']
etq = {x: f'{DIAS_ES[x.weekday()]} {x:%d/%m}' for x in dias}

tot_u = int(det.Cantidad.sum())
tot_ars = int(det.Total.sum())
por_dia = det.groupby('dia').Cantidad.sum().reindex(dias, fill_value=0)
por_tit = det.groupby('titulo').Cantidad.sum().sort_values(ascending=False)
mat = det.pivot_table(index=['nro', 'local'], columns='dia', values='Cantidad',
                      aggfunc='sum', fill_value=0).reindex(columns=dias, fill_value=0)
mat['Total'] = mat.sum(axis=1)
primera = det[det.Cantidad > 0].groupby(['nro', 'local']).dia.min()
ult_completo = dias[-2]
def ritmo(row):
    p = primera[row.name]
    n = (ult_completo - p).days + 1
    u = sum(row[x] for x in dias if p <= x <= ult_completo)
    return round(u / n, 1) if n > 0 else None
mat['Primera venta'] = [primera[i] for i in mat.index]
mat['Ritmo'] = mat.apply(ritmo, axis=1)
devol = det[det.Cantidad < 0]
mat['Devol.'] = devol.groupby(['nro', 'local']).Cantidad.sum().reindex(mat.index, fill_value=0)
mat = mat.sort_values(['Total', 'Ritmo'], ascending=False)
tit_cols = list(por_tit.index)
mt = det.pivot_table(index=['nro', 'local'], columns='titulo', values='Cantidad',
                     aggfunc='sum', fill_value=0).reindex(columns=tit_cols, fill_value=0)
mt['Total'] = mt.sum(axis=1)
mt = mt.loc[mat.index]

con_venta = set(det.local)
sin_venta = sorted(padron - con_venta)
otra_editorial = set(agr['Nombre Sucursal'])       # vendieron Editorial en el periodo
n_locales = len(con_venta)
n_padron = len(padron)

assert int(mat.Total.sum()) == tot_u == int(mt.Total.sum()) == int(por_tit.sum()) == int(por_dia.sum())

# ---------- estilo ----------
AZUL, AZUL_OSC, GRIS = '2A78D6', '1C5CAB', 'C3C2B7'
TINTA, TINTA2, FONDO, BORDE = '0B0B0B', '52514E', 'F5F6F8', 'D9DCE1'
fino = Side(style='thin', color=BORDE)
caja = Border(left=fino, right=fino, top=fino, bottom=fino)
H = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
HF = PatternFill('solid', fgColor='1F4E78')
def cab(ws, fila, textos, col0=1):
    for i, t in enumerate(textos):
        c = ws.cell(fila, col0 + i, t)
        c.font, c.fill, c.border = H, HF, caja
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
def calor(ws, rango):
    ws.conditional_formatting.add(rango, ColorScaleRule(
        start_type='num', start_value=0, start_color='FFFFFF',
        end_type='max', end_color='6DA7EC'))

wb = Workbook()

# ---------- hoja Resumenes (fuente de los graficos y vista de tabla) ----------
rs = wb.active
rs.title = 'Resúmenes'
rs['A1'] = 'Unidades por día'
rs['A1'].font = Font(bold=True, size=12)
cab(rs, 2, ['Día', 'Unidades'])
for i, x in enumerate(dias, start=3):
    rs.cell(i, 1, etq[x] + (' (parcial)' if x == ult else ''))
    rs.cell(i, 2, int(por_dia[x]))
fd = 2 + len(dias)
rs['D1'] = 'Unidades por título'
rs['D1'].font = Font(bold=True, size=12)
cab(rs, 2, ['Título', 'Unidades', '% del total', 'Locales'], col0=4)
for i, (t, u) in enumerate(por_tit.items(), start=3):
    rs.cell(i, 4, t); rs.cell(i, 5, int(u))
    rs.cell(i, 6, u / tot_u).number_format = '0%'
    rs.cell(i, 7, int(det[(det.titulo == t) & (det.Cantidad > 0)].local.nunique()))
ft = 2 + len(por_tit)
rs['I1'] = 'Top 15 locales'
rs['I1'].font = Font(bold=True, size=12)
cab(rs, 2, ['Local', 'Unidades'], col0=9)
top = mat.Total.head(15)[::-1]                      # al reves: la barra mas larga arriba
for i, ((n, loc), u) in enumerate(top.items(), start=3):
    rs.cell(i, 9, loc.title()); rs.cell(i, 10, int(u))
fl = 2 + len(top)
for c, w in zip('ABCDEFGHIJ', (16, 10, 3, 16, 10, 11, 9, 3, 30, 10)):
    rs.column_dimensions[c].width = w

# ---------- hoja Tablero ----------
tb = wb.create_sheet('Tablero', 0)
tb.sheet_view.showGridLines = False
for c in range(1, 17):
    tb.column_dimensions[L(c)].width = 10.5
for r in range(1, 70):
    for c in range(1, 17):
        tb.cell(r, c).fill = PatternFill('solid', fgColor=FONDO)
tb.merge_cells('A1:P1')
tb['A1'] = 'Libros Sorpresa · ventas por local'
tb['A1'].font = Font(name='Calibri', bold=True, size=20, color=TINTA)
tb.merge_cells('A2:P2')
tb['A2'] = (f'Solo locales físicos · del {dias[0]:%d/%m} al {ult:%d/%m/%Y} · '
            f'el {ult:%d/%m} es parcial (el reporte se sacó ese mismo día a la mañana) · '
            f'fuente: Detalle Venta Artículo, cuadrado contra Ventas Proveedor Agrupadas')
tb['A2'].font = Font(size=10, color=TINTA2)
tb.row_dimensions[1].height = 30

mejor = por_dia[por_dia.index != ult].idxmax()
kpis = [
    ('Unidades netas', f'{tot_u:,}'.replace(',', '.'), f'{len(devol)} devoluciones ya restadas'),
    ('Facturación', '$ ' + f'{tot_ars:,}'.replace(',', '.'), f'a $ {int(det.Total.sum() / tot_u):,} c/u'.replace(',', '.')),
    ('Locales con venta', f'{n_locales} de {n_padron}', f'{len(sin_venta)} sin venta (ver hoja)'),
    ('Mejor día', f'{etq[mejor]}', f'{int(por_dia[mejor])} u · {por_dia[mejor] / tot_u:.0%} del total'),
    ('Promedio por local', f'{tot_u / n_locales:.1f}'.replace('.', ','), 'unidades en el período'),
]
for k, (lab, val, nota) in enumerate(kpis):
    c0 = 1 + k * 3 + (k > 0) * 0
    c0 = [1, 4, 7, 10, 13][k]
    for r in (4, 5, 6):
        tb.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 2)
        for c in range(c0, c0 + 3):
            tb.cell(r, c).fill = PatternFill('solid', fgColor='FFFFFF')
            tb.cell(r, c).border = Border(left=fino if c == c0 else None, right=fino if c == c0 + 2 else None,
                                          top=fino if r == 4 else None, bottom=fino if r == 6 else None)
    tb.cell(4, c0, lab).font = Font(size=10, color=TINTA2)
    tb.cell(5, c0, val).font = Font(size=20, bold=True, color=TINTA)
    tb.cell(6, c0, nota).font = Font(size=9, color=TINTA2)
    for r in (4, 5, 6):
        tb.cell(r, c0).alignment = Alignment(horizontal='left', vertical='center', indent=1)
tb.row_dimensions[5].height = 30

def barras(titulo, cats, vals, horizontal=False, ancho=15, alto=7.5):
    ch = BarChart()
    ch.type = 'bar' if horizontal else 'col'
    ch.title = titulo
    ch.style = 10
    ch.legend = None
    ch.add_data(vals, titles_from_data=False)
    ch.set_categories(cats)
    s = ch.series[0]
    s.graphicalProperties = GraphicalProperties(solidFill=AZUL)
    s.graphicalProperties.line = LineProperties(noFill=True)
    ch.gapWidth = 60
    ch.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False,
                                   showLegendKey=False, showPercent=False, showLeaderLines=False)
    ch.y_axis.majorGridlines = None
    ch.y_axis.delete = True
    ch.x_axis.delete = False
    ch.width, ch.height = ancho, alto
    return ch

c1 = barras('Unidades por día', Reference(rs, min_col=1, min_row=3, max_row=fd),
            Reference(rs, min_col=2, min_row=3, max_row=fd))
p = DataPoint(idx=len(dias) - 1)                    # el dia parcial en gris
p.graphicalProperties = GraphicalProperties(solidFill=GRIS)
c1.series[0].dPt.append(p)
tb.add_chart(c1, 'A8')
c2 = barras('Unidades por título', Reference(rs, min_col=4, min_row=3, max_row=ft),
            Reference(rs, min_col=5, min_row=3, max_row=ft))
tb.add_chart(c2, 'I8')
c3 = barras('Top 15 locales (unidades)', Reference(rs, min_col=9, min_row=3, max_row=fl),
            Reference(rs, min_col=10, min_row=3, max_row=fl), horizontal=True, ancho=31, alto=10)
tb.add_chart(c3, 'A24')

# lectura
r0 = 45
tb.cell(r0, 1, 'Lo que muestran los datos').font = Font(bold=True, size=13, color=TINTA)
arranque = primera.value_counts().sort_index()
lideres = por_tit.head(2)
bio = por_tit.get('Bío', 0)
lineas = [
    f'{lideres.index[0]} ({int(lideres.iloc[0])} u, {lideres.iloc[0] / tot_u:.0%}) y {lideres.index[1]} '
    f'({int(lideres.iloc[1])} u) son los más vendidos; Bío quedó muy abajo ({int(bio)} u, vendido en '
    f'{det[(det.titulo == "Bío") & (det.Cantidad > 0)].local.nunique()} de {n_locales} locales).',
    f'El pico fue el {etq[mejor]} con {int(por_dia[mejor])} u ({por_dia[mejor] / tot_u:.0%} de todo lo vendido); '
    'después: ' + ', '.join(f'{etq[x]} {int(por_dia[x])} u' for x in dias if mejor < x < ult) +
    '. Con tan pocos días todavía no se puede hablar de tendencia.',
    'Los locales no arrancaron todos juntos: ' + ', '.join(
        f'{int(n)} el {x:%d/%m}' for x, n in arranque.items()) +
    '. Para comparar locales que arrancaron en días distintos, usar la columna "Ritmo" de la hoja Local x día.',
    f'Sin ninguna venta de Sorpresa: {", ".join(s.title() for s in sin_venta)}. Todos menos Devoto sí '
    'vendieron otros títulos de la Editorial según Ventas Agrupadas, así que conviene confirmar si les llegaron. '
    'A Devoto la repo se le pasó recién el 05/10.',
]
for i, t in enumerate(lineas):
    tb.merge_cells(start_row=r0 + 1 + i, start_column=1, end_row=r0 + 1 + i, end_column=16)
    c = tb.cell(r0 + 1 + i, 1, '•  ' + t)
    c.font = Font(size=10.5, color=TINTA)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    tb.row_dimensions[r0 + 1 + i].height = 30

# ---------- hoja Local x dia ----------
ld = wb.create_sheet('Local x día', 1)
ld.freeze_panes = 'C3'
ld['A1'] = (f'Unidades por local y por día · ordenado por total · el {ult:%d/%m} es parcial · '
            f'Ritmo = unidades por día desde su primera venta hasta el {ult_completo:%d/%m}')
ld['A1'].font = Font(italic=True, size=9, color=TINTA2)
cols = ['Nro', 'Local'] + [etq[x] + ('*' if x == ult else '') for x in dias] + \
       ['Total', 'Primera venta', 'Ritmo u/día', 'Devol.']
cab(ld, 2, cols)
for i, ((n, loc), row) in enumerate(mat.iterrows(), start=3):
    vals = [n, loc] + [int(row[x]) for x in dias] + [int(row.Total), row['Primera venta'],
                                                     row.Ritmo, int(row['Devol.'])]
    for j, v in enumerate(vals, start=1):
        c = ld.cell(i, j, v); c.border = caja
        c.alignment = Alignment(horizontal='left' if j == 2 else 'center')
    ld.cell(i, 3 + len(dias)).font = Font(bold=True)
    ld.cell(i, 4 + len(dias)).number_format = 'DD/MM'
    ld.cell(i, 5 + len(dias)).number_format = '0.0'
    if row['Devol.'] == 0:
        ld.cell(i, 6 + len(dias)).value = None
uf = 2 + len(mat)
tr = uf + 1
ld.cell(tr, 2, 'TOTAL').font = Font(bold=True)
for j in range(3, 4 + len(dias)):
    c = ld.cell(tr, j, f'=SUM({L(j)}3:{L(j)}{uf})')
    c.font = Font(bold=True); c.border = caja; c.alignment = Alignment(horizontal='center')
calor(ld, f'C3:{L(2 + len(dias))}{uf}')
for c, w in zip(range(1, len(cols) + 1), [6, 30] + [10] * len(dias) + [8, 10, 9, 7]):
    ld.column_dimensions[L(c)].width = w

# ---------- hoja Local x titulo ----------
lt = wb.create_sheet('Local x título', 2)
lt.freeze_panes = 'C3'
lt['A1'] = 'Unidades por local y por título · mismo orden que Local x día'
lt['A1'].font = Font(italic=True, size=9, color=TINTA2)
cab(lt, 2, ['Nro', 'Local'] + tit_cols + ['Total'])
for i, ((n, loc), row) in enumerate(mt.iterrows(), start=3):
    for j, v in enumerate([n, loc] + [int(row[t]) for t in tit_cols] + [int(row.Total)], start=1):
        c = lt.cell(i, j, v); c.border = caja
        c.alignment = Alignment(horizontal='left' if j == 2 else 'center')
    lt.cell(i, 3 + len(tit_cols)).font = Font(bold=True)
uf2 = 2 + len(mt)
lt.cell(uf2 + 1, 2, 'TOTAL').font = Font(bold=True)
for j in range(3, 4 + len(tit_cols)):
    c = lt.cell(uf2 + 1, j, f'=SUM({L(j)}3:{L(j)}{uf2})')
    c.font = Font(bold=True); c.border = caja; c.alignment = Alignment(horizontal='center')
calor(lt, f'C3:{L(2 + len(tit_cols))}{uf2}')
for c, w in zip(range(1, len(tit_cols) + 4), [6, 30] + [13] * len(tit_cols) + [8]):
    lt.column_dimensions[L(c)].width = w

# ---------- hoja Sin venta ----------
sv = wb.create_sheet('Sin venta', 3)
cab(sv, 1, ['Local físico', '¿Tiene ventas de otros títulos de la Editorial en Ventas Agrupadas?'])
for i, s in enumerate(sin_venta, start=2):
    sv.cell(i, 1, s).border = caja
    txt = 'Sí' if s in otra_editorial else 'No'
    if s == 'DEVOTO':
        txt += ' (la repo de Sorpresa se pasó el 05/10)'
    sv.cell(i, 2, txt).border = caja
sv.column_dimensions['A'].width = 30
sv.column_dimensions['B'].width = 55

# ---------- hoja Devoluciones ----------
dv = wb.create_sheet('Devoluciones', 4)
cab(dv, 1, ['Local', 'Fecha', 'Título', 'Cantidad', 'Total $'])
for i, r in enumerate(devol.sort_values(['Fecha', 'local']).itertuples(), start=2):
    for j, v in enumerate([r.local, r.Fecha, r.titulo, int(r.Cantidad), int(r.Total)], start=1):
        c = dv.cell(i, j, v); c.border = caja
    dv.cell(i, 2).number_format = 'DD/MM/YYYY'
for c, w in zip('ABCDE', (30, 12, 16, 10, 10)):
    dv.column_dimensions[c].width = w

# ---------- hoja Datos ----------
dt = wb.create_sheet('Datos')
cab(dt, 1, ['Nro', 'Local', 'Fecha', 'Artículo', 'ISBN', 'Título', 'Cantidad', 'Total $'])
for i, r in enumerate(det.sort_values(['local', 'Fecha']).itertuples(), start=2):
    for j, v in enumerate([r.nro, r.local, r.Fecha, r.Articulo, r.ISBN, r.titulo,
                           int(r.Cantidad), int(r.Total)], start=1):
        dt.cell(i, j, v)
    dt.cell(i, 3).number_format = 'DD/MM/YYYY'
dt.auto_filter.ref = f'A1:H{len(det) + 1}'
dt.freeze_panes = 'A2'
for c, w in zip('ABCDEFGH', (6, 30, 12, 10, 16, 16, 10, 10)):
    dt.column_dimensions[c].width = w

# impresion: el tablero en una hoja apaisada; las tablas al ancho de la pagina
for ws, alto in ((tb, 1), (ld, 0), (lt, 0)):
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = alto
tb.print_area = 'A1:P50'
wb.save(salida)
print(f'{tot_u} u | $ {tot_ars:,} | {n_locales} locales de {n_padron} | {len(dias)} dias | '
      f'devol {len(devol)} | sin venta {sin_venta} | cuadre agrupadas OK ({len(cuadre)} celdas)')
print('por dia', {etq[x]: int(v) for x, v in por_dia.items()})
print('por titulo', por_tit.to_dict())
for t in lineas: print(' -', t)
