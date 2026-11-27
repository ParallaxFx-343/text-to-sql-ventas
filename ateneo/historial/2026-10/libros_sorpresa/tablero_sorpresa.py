"""Tablero de ventas de los Libros Sorpresa por local (solo locales fisicos).

Uso: [EXPORT=aaaa-mm-dd] python -P tablero_sorpresa.py DETALLE.xlsx AGRUPADAS.xlsx config.json SALIDA.xlsx [SII [REFUERZO.txt [SII_BASE]]]
- DETALLE: "Detalle Venta Articulo" (una fila por venta: Sucursal "NOMBRE  (nro)",
  Fecha, Articulo, Titulo, Cantidad, Bruto, Total, ISBN)
- AGRUPADAS: "Ventas Proveedor Agrupadas Articulo" ("-" si no se tiene; entonces no hay cuadre) (solo para cuadrar y para
  saber que locales vendieron Editorial pero no Sorpresa)
- config.json: padron de locales fisicos (orden_paec)
- SII (opcional): stock por local, para "cuanto le queda" y "le alcanza" (dias).
  Del deposito no se muestra nada (Franco: el stock de deposito se va a liberar).
- REFUERZO (opcional, requiere SII): el DESGLOSE_x_SUCURSAL del refuerzo que se cargo,
  con los nros de pedido que Franco anota al lado de cada local ("25 - 277602").
  Agrega la hoja "Refuerzo 07-10" y, junto a "le quedan", lo que tendra cada local
  con el refuerzo (stock del SII + refuerzo, sin descontar lo vendido despues del SII).
- SII_BASE (opcional): el SII con el que se calculo el refuerzo. Con el, para cada
  local y titulo se estima cuanto del refuerzo ya entro al local:
  entro = SII - (SII_BASE - vendido entre los dos SII), acotado entre 0 y el refuerzo;
  "falta llegar" = refuerzo - entro. Sin SII_BASE se toma que no llego nada.
Los exports crudos del sistema (estilos rotos, todos los articulos) se normalizan
antes con normalizar_ventas.py.
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
f_sii = sys.argv[5] if len(sys.argv) > 5 else None
f_ref = sys.argv[6] if len(sys.argv) > 6 else None
f_base = sys.argv[7] if len(sys.argv) > 7 else None
padron = set(json.load(open(f_cfg, encoding='utf-8'))['orden_paec']) | {'DEVOTO'}

# ---------- datos ----------
det = pd.read_excel(f_det, dtype={'ISBN': str})
det['nro'] = det.Sucursal.str.extract(r'\((\d+)\)\s*$')[0].astype(int)
det['local'] = det.Sucursal.str.replace(r'\s*\(\d+\)\s*$', '', regex=True).str.strip()
det = det[det.local.isin(padron)].copy()          # solo locales fisicos
NOMBRES = {'HISTORICA': 'Histórica', 'ROMANTICA': 'Romántica', 'CONTEMPORANEA': 'Contemporánea',
           'MISTERIO': 'Misterio', 'EROTICA': 'Erótica', 'BIO': 'Bío'}
ACENTOS = {'SUPERACION': 'Superación', 'FICCION': 'Ficción'}
def nombre(x):
    """'LIBRO SORPRESA SUPERACION PERSONAL 1' -> 'Superación personal 1' (los 6 de octubre, como siempre)."""
    x = str(x).replace('LIBRO SORPRESA ', '').strip()
    if x in NOMBRES:
        return NOMBRES[x]
    w = x.split()
    return ' '.join([ACENTOS.get(w[0], w[0].capitalize())] + [v.lower() for v in w[1:]])
ORIG = set(NOMBRES.values())                       # coleccion de octubre (lanzada el 01/10)
coleccion = lambda tt: 'Octubre' if tt in ORIG else 'Noviembre'
det['titulo'] = det.Titulo.map(nombre)
assert det.groupby('titulo').Articulo.nunique().max() == 1, 'dos articulos con el mismo nombre'
det['dia'] = det.Fecha.dt.date

agr = None
cuadre = []
extra_agr = 0
if f_agr != '-':
    agr = pd.read_excel(f_agr, dtype={'Código / ISBN': str})
    agr = agr[agr.Canal.isna() | (agr.Canal.astype(str).str.strip() == '')]   # sin canales online
    agr = agr[agr['Nombre Sucursal'].isin(padron)]
    sor = agr[agr.TITULO.astype(str).str.contains('LIBRO SORPRESA')]

    # cuadre por dos caminos: detalle vs agrupadas, local x articulo
    a = sor.groupby(['Código Sucursal', 'ID ARTICULO']).Cantidad.sum()
    # CUADRE_DESDE=aaaa-mm-dd: Agrupadas cubre solo desde esa fecha (27/11: Detalle nuevo desde el 01/11, octubre sale de la planilla anterior)
    _cd = __import__('os').environ.get('CUADRE_DESDE')
    d = (det[det.Fecha >= pd.Timestamp(_cd)] if _cd else det).groupby(['nro', 'Articulo']).Cantidad.sum()
    a.index.names = d.index.names = ['s', 'a']
    cuadre = pd.concat([a, d], axis=1).fillna(0)
    dif_c = cuadre[cuadre.iloc[:, 0] != cuadre.iloc[:, 1]]
    # tolerancia: Agrupadas sacado despues que el Detalle puede traer ventas extra del dia parcial
    assert (dif_c.iloc[:, 0] >= dif_c.iloc[:, 1]).all() and (dif_c.iloc[:, 0] - dif_c.iloc[:, 1]).sum() <= 10, \
        f'detalle y agrupadas no cuadran:\n{dif_c}'
    extra_agr = int((dif_c.iloc[:, 0] - dif_c.iloc[:, 1]).sum())

dias = sorted(det.dia.unique())
ult = dias[-1]                                      # ultimo dia con ventas
# EXPORT=aaaa-mm-dd (fecha en que se saco el Detalle): si es posterior al ultimo dia, ese dia esta completo.
import os as _os
_exp = _os.environ.get('EXPORT')
parcial = ult if (not _exp or pd.Timestamp(_exp).date() <= ult) else None   # dia parcial (export a la mañana) o ninguno
hasta = max(det.Fecha)                              # para el subtitulo
DIAS_ES = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']
etq = {x: f'{DIAS_ES[x.weekday()]} {x:%d/%m}' for x in dias}

tot_u = int(det.Cantidad.sum())
tot_ars = int(det.Total.sum())
por_dia = det.groupby('dia').Cantidad.sum().reindex(dias, fill_value=0)
por_tit = det.groupby('titulo').Cantidad.sum()
por_tit = por_tit.loc[sorted(por_tit.index, key=lambda x: (coleccion(x) != 'Octubre', -por_tit[x], x))]   # octubre primero
lanz = det[det.Cantidad > 0].groupby('titulo').dia.min()          # primera venta de cada titulo en la cadena
mat = det.pivot_table(index=['nro', 'local'], columns='dia', values='Cantidad',
                      aggfunc='sum', fill_value=0).reindex(columns=dias, fill_value=0)
mat['Total'] = mat.sum(axis=1)
primera = det[det.Cantidad > 0].groupby(['nro', 'local']).dia.min()
ult_completo = dias[-2] if parcial is not None else dias[-1]
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

# ---------- stock en locales (SII) ----------
stock = None
if f_sii:
    def leer_stock(f):
        sii = pd.read_excel(f)
        cols_s = list(sii.columns)
        suc_s = [c for c in cols_s[cols_s.index('CLASIFIC') + 1:] if c in padron]
        assert len(suc_s) == len(padron), f'faltan locales en el SII: {padron - set(suc_s)}'
        s = sii[sii.TITULO.astype(str).str.startswith('LIBRO SORPRESA')].copy()
        s['titulo'] = s.TITULO.map(nombre)
        assert s.titulo.is_unique
        st = s.set_index('titulo')[suc_s].fillna(0).astype(int).T   # local x titulo
        sin_v = sorted(set(st.columns) - set(por_tit.index), key=lambda x: (coleccion(x) != 'Octubre', x))
        st = st.reindex(columns=list(por_tit.index) + sin_v, fill_value=0)   # al final, los que todavia no vendieron
        st['Total'] = st.sum(axis=1)
        return st

    def fecha_de(f):
        import re as _r
        m = _r.search(r'(20\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', f.split('/')[-1])
        return pd.Timestamp(*map(int, m.groups())) if m else None
    stock = leer_stock(f_sii)
    loc_de = {loc: (n, loc) for n, loc in mat.index}
    mat['Stock'] = [int(stock.loc[loc, 'Total']) for _, loc in mat.index]
    mat['Alcanza'] = [round(r.Stock / r.Ritmo, 1) if r.Ritmo else None for _, r in mat.iterrows()]
    import re as _re
    m_fecha = _re.search(r'(20\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', f_sii.split('/')[-1])
    fecha_sii = f'{m_fecha.group(3)}/{m_fecha.group(2)} {m_fecha.group(4)}:{m_fecha.group(5)}' if m_fecha else 'SII'
    llenos = [x for x in sorted(det.dia.unique()) if x != parcial]   # sin el dia parcial
    dias_t = {x: (ult_completo - lanz[x]).days + 1 for x in lanz.index}      # dias desde que salio cada titulo
    MIN_DIAS = 7                                          # con menos de una semana de venta el ritmo no dice nada
    vend_d = det[det.dia.isin(llenos)].groupby('titulo').Cantidad.sum() / pd.Series(dias_t)
    vend_d = vend_d[[dias_t[x] >= MIN_DIAS for x in vend_d.index]]
    alcanza_tit = {t: (stock[t].sum() / vend_d[t]) for t in por_tit.index if vend_d.get(t, 0) > 0}

# ---------- refuerzo cargado (TXT por sucursal con nros de pedido) ----------
ref = None
if f_ref:
    import re as _re
    filas_r, pedidos, zona_de, tot_hdr = [], {}, {}, {}
    zona = loc = None
    for lin in open(f_ref, encoding='utf-8-sig').read().split('\n'):
        m_z = _re.match(r'#\s*(AMBA|INTERIOR):', lin)
        if m_z:
            zona = m_z.group(1); continue
        if not lin.strip() or lin.startswith(('#', '=')):
            continue
        m_i = _re.match(r'\s+(\d+)\s+LIBRO SORPRESA (\w+)\s+EAN (\d+)\s+ID (\d+)(.*)$', lin)
        m_p = _re.search(r'\b25\s*-?\s*(\d{6})\s*$', lin)
        if m_i:
            filas_r.append((loc, NOMBRES[m_i.group(2)], int(m_i.group(1))))
        else:
            m_h = _re.match(r'(\S.*?)\s{2,}(\d+) u\b', lin)
            if not m_h or zona is None or m_h.group(1) == 'TOTAL':
                continue
            loc = m_h.group(1).strip()
            assert loc in padron, f'local del refuerzo que no esta en el padron: {loc}'
            zona_de[loc], tot_hdr[loc] = zona, int(m_h.group(2))
        if m_p and loc:
            pedidos[loc] = f'25-{m_p.group(1)}'
    ref = pd.DataFrame(filas_r, columns=['local', 'titulo', 'cant']).pivot_table(
        index='local', columns='titulo', values='cant', aggfunc='sum', fill_value=0)
    ref = ref.reindex(columns=list(por_tit.index), fill_value=0)
    ref['Total'] = ref.sum(axis=1)
    assert all(ref.loc[l, 'Total'] == t for l, t in tot_hdr.items()), 'el total de un local no cuadra'
    sin_pedido = sorted(set(ref.index) - set(pedidos))
    ref_full = ref.reindex(stock.index, fill_value=0)          # todos los locales del padron
    tits = list(por_tit.index)
    entro = None
    if f_base:
        base = leer_stock(f_base)
        f0, f1 = fecha_de(f_base), fecha_de(f_sii)
        assert f0 is not None and f1 is not None and f0 < f1, 'no se pudo leer la fecha de los SII'
        entre = det[(det.dia >= f0.date()) & (det.dia < f1.date())]
        v_entre = entre.pivot_table(index='local', columns='titulo', values='Cantidad', aggfunc='sum',
                                    fill_value=0).reindex(index=stock.index, columns=tits, fill_value=0)
        mov = stock[tits] - (base[tits] - v_entre)              # entradas netas entre los dos SII
        entro = mov.clip(lower=0).where(mov.clip(lower=0) < ref_full[tits], ref_full[tits])
        otras = (mov - ref_full[tits]).clip(lower=0).sum(axis=1)  # entro mas que el refuerzo
        entro['Total'] = entro.sum(axis=1)
        fecha_base = f'{f0:%d/%m %H:%M}'
    pend = ref_full[tits] - (entro[tits] if entro is not None else 0)
    pend['Total'] = pend.sum(axis=1)
    con_ref = stock + pend

con_venta = set(det.local)
sin_venta = sorted(padron - con_venta)
otra_editorial = set(agr['Nombre Sucursal']) if agr is not None else None   # vendieron Editorial en el periodo
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
    rs.cell(i, 1, etq[x] + (' (parcial)' if x == parcial else ''))
    rs.cell(i, 2, int(por_dia[x]))
fd = 2 + len(dias)
rs['D1'] = 'Unidades por título'
rs['D1'].font = Font(bold=True, size=12)
cab(rs, 2, ['Título', 'Unidades', '% del total', 'Locales', 'Colección', 'Primera venta'], col0=4)
for i, (t, u) in enumerate(por_tit.items(), start=3):
    rs.cell(i, 4, t); rs.cell(i, 5, int(u))
    rs.cell(i, 6, u / tot_u).number_format = '0%'
    rs.cell(i, 7, int(det[(det.titulo == t) & (det.Cantidad > 0)].local.nunique()))
    rs.cell(i, 8, coleccion(t)); rs.cell(i, 9, lanz.get(t)).number_format = 'DD/MM'
ft = 2 + len(por_tit)
rs['K1'] = 'Top 15 locales'
rs['K1'].font = Font(bold=True, size=12)
cab(rs, 2, ['Local', 'Unidades'], col0=11)
top = mat.Total.head(15)[::-1]                      # al reves: la barra mas larga arriba
for i, ((n, loc), u) in enumerate(top.items(), start=3):
    rs.cell(i, 11, loc.title()); rs.cell(i, 12, int(u))
fl = 2 + len(top)
# por semana y coleccion (lunes a domingo)
det['semana'] = [x - pd.Timedelta(days=x.weekday()) for x in pd.to_datetime(det.dia)]
sem = det.pivot_table(index='semana', columns=det.titulo.map(coleccion), values='Cantidad', aggfunc='sum', fill_value=0)
sem = sem.reindex(columns=['Octubre', 'Noviembre'], fill_value=0)
rs['N1'] = 'Unidades por semana (lun a dom)'
rs['N1'].font = Font(bold=True, size=12)
cab(rs, 2, ['Semana', 'Colección octubre', 'Colección noviembre', 'Total'], col0=14)
for i, (w, row) in enumerate(sem.iterrows(), start=3):
    ini = max(w, pd.Timestamp(dias[0])); fin = min(w + pd.Timedelta(days=6), pd.Timestamp(ult))
    rs.cell(i, 14, f'{ini:%d/%m} al {fin:%d/%m}' + (' (parcial)' if fin.date() == ult and (parcial is not None or fin.weekday() != 6) else ''))
    rs.cell(i, 15, int(row.Octubre)); rs.cell(i, 16, int(row.Noviembre)); rs.cell(i, 17, int(row.sum()))
fs = 2 + len(sem)
rs.column_dimensions['N'].width = 22
for c in 'OPQ':
    rs.column_dimensions[c].width = 12
for c, w in zip('ABCDEFGHIJKL', (16, 10, 3, 24, 10, 11, 9, 11, 12, 3, 30, 10)):
    rs.column_dimensions[c].width = w

# ---------- hoja Tablero ----------
tb = wb.create_sheet('Tablero', 0)
tb.sheet_view.showGridLines = False
for c in range(1, 17):
    tb.column_dimensions[L(c)].width = 10.5
for r in range(1, 80):
    for c in range(1, 17):
        tb.cell(r, c).fill = PatternFill('solid', fgColor=FONDO)
tb.merge_cells('A1:P1')
tb['A1'] = 'Libros Sorpresa · ventas por local'
tb['A1'].font = Font(name='Calibri', bold=True, size=20, color=TINTA)
tb.merge_cells('A2:P2')
tb['A2'] = (f'Solo locales físicos · del {dias[0]:%d/%m} al {ult:%d/%m/%Y} · '
            + (f'el {ult:%d/%m} es parcial (el reporte se sacó ese mismo día) · ' if parcial is not None else f'datos completos hasta el {ult:%d/%m} · ')
            + (('fuente: Detalle Venta Artículo, cuadrado contra Ventas Proveedor Agrupadas'
               + (f' desde el {pd.Timestamp(_cd):%d/%m}' if _cd else '')
               + (f' (Agrupadas trae {extra_agr} u más: se sacó después, son ventas del {(ult if parcial is not None else pd.Timestamp(_exp).date()):%d/%m})' if extra_agr else ''))
              if agr is not None else 'fuente: Detalle Venta Artículo (sin Ventas Agrupadas para cuadrar)'))
tb['A2'].font = Font(size=10, color=TINTA2)
tb.row_dimensions[1].height = 30

mejor = por_dia[por_dia.index != parcial].idxmax()
nov_t = [x for x in por_tit.index if coleccion(x) == 'Noviembre']
u_nov = int(por_tit[nov_t].sum()); n_nov = int((por_tit[nov_t] > 0).sum()); d_nov = min(lanz[x] for x in nov_t)
u_oct = int(por_tit[[x for x in por_tit.index if coleccion(x) == 'Octubre']].sum())
kpis = [
    ('Unidades netas', f'{tot_u:,}'.replace(',', '.'), f'{len(devol)} devoluciones ya restadas'),
    ('Facturación', '$ ' + f'{tot_ars:,}'.replace(',', '.'), f'a $ {int(det.Total.sum() / tot_u):,} c/u'.replace(',', '.')),
    ('Locales con venta', f'{n_locales} de {n_padron}', f'{len(sin_venta)} sin venta (ver hoja)'),
    ('Colección noviembre', f'{u_nov:,}'.replace(',', '.') + ' u',
     f'{n_nov} títulos con venta desde el {d_nov:%d/%m}'),
    (('Le queda en los locales', f'{int(stock.Total.sum()):,}'.replace(',', '.') + ' u',
      f'SII {fecha_sii} · ver hoja Le quedan x título') if stock is not None else
     ('Promedio por local', f'{tot_u / n_locales:.1f}'.replace('.', ','), 'unidades en el período')),
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
if parcial is not None:
    p = DataPoint(idx=len(dias) - 1)                # el dia parcial en gris
    p.graphicalProperties = GraphicalProperties(solidFill=GRIS)
    c1.series[0].dPt.append(p)
if len(dias) > 20:
    c1.dataLabels = None                            # con muchos dias las etiquetas se pisan
tb.add_chart(c1, 'A8')
c2 = barras('Unidades por semana', Reference(rs, min_col=14, min_row=3, max_row=fs),
            Reference(rs, min_col=17, min_row=3, max_row=fs))
tb.add_chart(c2, 'I8')
c3 = barras('Top 15 locales (unidades)', Reference(rs, min_col=11, min_row=3, max_row=fl),
            Reference(rs, min_col=12, min_row=3, max_row=fl), horizontal=True, ancho=15, alto=15)
tb.add_chart(c3, 'A24')
c4 = barras('Unidades por título (primero la colección de octubre)', Reference(rs, min_col=4, min_row=3, max_row=ft),
            Reference(rs, min_col=5, min_row=3, max_row=ft), horizontal=True, ancho=15, alto=15)
c4.y_axis.delete = True
c4.x_axis.tickLblSkip = 1
c4.x_axis.scaling.orientation = 'maxMin'            # el primero arriba
tb.add_chart(c4, 'I24')

# lectura
r0 = 54
tb.cell(r0, 1, 'Lo que muestran los datos').font = Font(bold=True, size=13, color=TINTA)
arranque = primera.value_counts().sort_index()
lideres = por_tit.head(2)
bio = por_tit.get('Bío', 0)
def dos_semanas(ll):
    """Compara la primera mitad de los dias completos con la segunda, y los sabados."""
    h = len(ll) // 2
    a, b = int(por_dia[ll[:h]].sum()), int(por_dia[ll[h:2 * h]].sum())
    m = lambda n: f'{n:,}'.replace(',', '.')
    return (f'Del {ll[0]:%d/%m} al {ll[h - 1]:%d/%m} se vendieron {m(a)} u y del {ll[h]:%d/%m} al '
            f'{ll[2 * h - 1]:%d/%m}, {m(b)} u ({b / a - 1:+.0%}). Los sábados son el pico: '
            + ', '.join(f'{etq[x]} {int(por_dia[x])} u' for x in ll if x.weekday() == 5)
            + f', contra {por_dia[[x for x in ll if x.weekday() != 5]].mean():.0f} u de promedio el resto de los días.')


po = por_tit[[x for x in por_tit.index if coleccion(x) == 'Octubre']].sort_values(ascending=False)
pn = por_tit[nov_t].sort_values(ascending=False)
lineas = [
    f'Colección de octubre (6 títulos, desde el 01/10): ' + f'{u_oct:,}'.replace(',', '.') + ' u; '
    + ', '.join(f'{x} {int(v)}' for x, v in po.items()) + '. '
    + f'Colección de noviembre ({n_nov} títulos con venta, desde el {d_nov:%d/%m}): ' + f'{u_nov:,}'.replace(',', '.') + ' u'
    + '; los que más venden: ' + ', '.join(f'{x} {int(v)}' for x, v in pn.head(5).items()) + '.',
    (f'El pico fue el {etq[mejor]} con {int(por_dia[mejor])} u ({por_dia[mejor] / tot_u:.0%} de todo lo vendido); '
     'después: ' + ', '.join(f'{etq[x]} {int(por_dia[x])} u' for x in dias if mejor < x and x != parcial) +
     '. Con tan pocos días todavía no se puede hablar de tendencia.') if len(dias) - 1 < 10 else
    dos_semanas([x for x in dias if x != parcial]),
    'Los locales no arrancaron todos juntos: ' + ', '.join(
        f'{int(n)} el {x:%d/%m}' for x, n in arranque.items()) +
    '. Para comparar locales que arrancaron en días distintos, usar la columna "Vendió por día" de la hoja Local x día.',
]
if stock is None:
    lineas.append(
        f'Sin ninguna venta de Sorpresa: {", ".join(s.title() for s in sin_venta)}. Todos menos Devoto sí '
        'vendieron otros títulos de la Editorial según Ventas Agrupadas, así que conviene confirmar si les llegaron. '
        'A Devoto la repo se le pasó recién el 05/10.')
else:
    tc_all = [c for c in stock.columns if c != 'Total']
    cero = [s for s in sin_venta if stock.loc[s, 'Total'] == 0]
    con = [s for s in sin_venta if stock.loc[s, 'Total'] > 0]
    orden = sorted(alcanza_tit.items(), key=lambda x: x[1])
    sabados = [x for x in llenos if x.weekday() == 5]
    lineas.append(
        f'Lo que le queda a los locales (SII {fecha_sii}), al ritmo de cada título desde su primera venta '
        f'(hasta el {llenos[-1]:%d/%m}; sin los que tienen menos de {MIN_DIAS} días de venta), alcanza para: '
        + ', '.join(f'{t} {d:.1f} días'.replace('.', ',') for t, d in orden[:3])
        + f'; {orden[-1][0]}, en cambio, tiene para {orden[-1][1]:.0f} días.'
        + (f' Un sábado se vende {por_dia[sabados].mean() / por_dia[llenos].mean():.1f}'.replace('.', ',')
           + ' veces el promedio diario, así que un fin de semana consume bastante más.' if sabados else ''))
    so = [x for x in tc_all if coleccion(x) == 'Octubre']; sn = [x for x in tc_all if coleccion(x) == 'Noviembre']
    m_ = lambda n: f'{int(n):,}'.replace(',', '.')
    nuevos_n = [x for x in sn if x not in lanz.index or (ult_completo - lanz[x]).days + 1 < MIN_DIAS]
    locs_n = lambda x: int((stock[x] > 0).sum())
    lineas.append(
        f'Stock en locales (SII {fecha_sii}): la colección de octubre está casi agotada, quedan {m_(stock[so].sum().sum())} u '
        f'(' + ', '.join(f'{x} {m_(stock[x].sum())}' for x in so) + '). '
        f'De la de noviembre hay {m_(stock[sn].sum().sum())} u de {len(sn)} títulos. '
        + (f'{len(nuevos_n)} títulos ({", ".join(nuevos_n)}) están solo en '
           f'{max(locs_n(x) for x in nuevos_n)} locales y empezaron a venderse el '
           f'{min((lanz[x] for x in nuevos_n if x in lanz.index), default=ult):%d/%m}: todavía no hay ritmo para medirlos. '
           if nuevos_n else '')
        + ('Sin stock de la colección de noviembre: ' + ', '.join(l.title() for l in stock.index if stock.loc[l, sn].sum() == 0) + '.'
           if any(stock.loc[l, sn].sum() == 0 for l in stock.index) else ''))
    partes = []
    if cero:
        partes.append(', '.join(s.title() for s in cero) + ' no tienen stock de Sorpresa')
    if con:
        partes.append(', '.join(f'{s.title()} ({int(stock.loc[s, "Total"])} u)' for s in con)
                      + (' tiene stock y no vendió' if len(con) == 1 else ' tienen stock y no vendieron')
                      )
    if partes:
        lineas.append('Sin ninguna venta en todo el período: ' + '; '.join(partes) + '.')
if ref is not None:
    miles = lambda n: f'{int(n):,}'.replace(',', '.')
    alc_r = sorted(((t, con_ref[t].sum() / vend_d[t]) for t in por_tit.index if vend_d.get(t, 0) > 0),
                   key=lambda x: x[1])
    if entro is not None:
        en_r = entro.loc[ref.index, 'Total']
        llego = [l for l in ref.index if pend.loc[l, 'Total'] == 0]
        nada = [l for l in ref.index if en_r[l] == 0]
        lineas.append(
            f'Refuerzo del 07/10 ({miles(ref.Total.sum())} u, {len(ref)} locales): según el SII del {fecha_sii} '
            f'ya entraron {miles(en_r.sum())} u; llegó completo a {len(llego)} locales, a {len(nada)} no les llegó nada '
            f'y {len(ref) - len(llego) - len(nada)} recibieron una parte (ver hoja Refuerzo 07-10). '
            f'Faltan llegar {miles(pend.Total.sum())} u: con eso los locales pasarían a {miles(con_ref.Total.sum())} u, '
            'que alcanzan para ' + ', '.join(f'{t} {d:.1f} días'.replace('.', ',') for t, d in alc_r[:-1])
            + f' y {alc_r[-1][0]} {alc_r[-1][1]:.0f} días.')
        mas = otras[otras > 2].sort_values(ascending=False)
        if len(mas):
            lineas.append(
                'Entre los dos SII entró más mercadería de la que tenía el refuerzo en: '
                + ', '.join(f'{l.title()} (+{int(v)} u)' for l, v in mas.items())
                + '; no sale de este refuerzo, así que hay otro envío que no está en estos archivos.')
    else:
        lineas.append(
            f'Refuerzo del 07/10: {miles(ref.Total.sum())} u en {len(ref)} locales (ver hoja Refuerzo 07-10). '
            f'Sumado a lo que quedaba, la cadena pasa a {miles(con_ref.Total.sum())} u, '
            'que alcanzan para: '
            + ', '.join(f'{t} {d:.1f} días'.replace('.', ',') for t, d in alc_r[:-1])
            + f' y {alc_r[-1][0]} {alc_r[-1][1]:.0f} días (no descuenta lo vendido después del SII).')
for i, t in enumerate(lineas):
    tb.merge_cells(start_row=r0 + 1 + i, start_column=1, end_row=r0 + 1 + i, end_column=16)
    c = tb.cell(r0 + 1 + i, 1, '•  ' + t)
    c.font = Font(size=10.5, color=TINTA)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    tb.row_dimensions[r0 + 1 + i].height = 30 if len(t) < 230 else 45

# ---------- hoja Local x dia ----------
ld = wb.create_sheet('Local x día', 1)
ld.freeze_panes = 'C3'
ld['A1'] = (f'Unidades por local y por día · ordenado por total · ' + (f'el {ult:%d/%m} es parcial · ' if parcial is not None else '') +
            f'"Vendió por día" = unidades por día desde su primera venta hasta el {ult_completo:%d/%m}'
            + (f' · "Le quedan" = stock del local en el SII del {fecha_sii} · "Le alcanza" = le quedan / vendió por día; '
               'en rojo, 3 días o menos' if stock is not None else '')
            + (' · "Refuerzo 07/10" = lo cargado ese día; "Falta llegar" = la parte que todavía no figura en el SII; '
               '"Con lo que falta" = le quedan + falta llegar' if ref is not None else ''))
ld['A1'].font = Font(italic=True, size=9, color=TINTA2)
cols = ['Nro', 'Local'] + [etq[x] + ('*' if x == parcial else '') for x in dias] + \
       ['Total', 'Primera venta', 'Vendió por día'] + \
       (['Le quedan', 'Le alcanza (días)'] if stock is not None else []) + \
       (['Refuerzo 07/10', 'Falta llegar', 'Con lo que falta', 'Le alcanza con eso (días)']
        if ref is not None else []) + ['Devol.']
extra = (2 if stock is not None else 0) + (4 if ref is not None else 0)
cab(ld, 2, cols)
for i, ((n, loc), row) in enumerate(mat.iterrows(), start=3):
    vals = [n, loc] + [int(row[x]) for x in dias] + [int(row.Total), row['Primera venta'], row.Ritmo] + \
           ([int(row.Stock), row.Alcanza] if stock is not None else []) + \
           ([int(ref_full.loc[loc, 'Total']) or None, int(pend.loc[loc, 'Total']) or None,
             int(con_ref.loc[loc, 'Total']),
             round(con_ref.loc[loc, 'Total'] / row.Ritmo, 1) if row.Ritmo else None]
            if ref is not None else []) + [int(row['Devol.'])]
    for j, v in enumerate(vals, start=1):
        c = ld.cell(i, j, v); c.border = caja
        c.alignment = Alignment(horizontal='left' if j == 2 else 'center')
    ld.cell(i, 3 + len(dias)).font = Font(bold=True)
    ld.cell(i, 4 + len(dias)).number_format = 'DD/MM'
    ld.cell(i, 5 + len(dias)).number_format = '0.0'
    if stock is not None:
        ld.cell(i, 7 + len(dias)).number_format = '0.0'
    if ref is not None:
        ld.cell(i, 11 + len(dias)).number_format = '0.0'
    if row['Devol.'] == 0:
        ld.cell(i, 6 + len(dias) + extra).value = None
uf = 2 + len(mat)
tr = uf + 1
ld.cell(tr, 2, 'TOTAL').font = Font(bold=True)
for j in list(range(3, 4 + len(dias))) + ([6 + len(dias)] if stock is not None else []) + \
        ([8 + len(dias), 9 + len(dias), 10 + len(dias)] if ref is not None else []):
    c = ld.cell(tr, j, f'=SUM({L(j)}3:{L(j)}{uf})')
    c.font = Font(bold=True); c.border = caja; c.alignment = Alignment(horizontal='center')
calor(ld, f'C3:{L(2 + len(dias))}{uf}')
if stock is not None:
    from openpyxl.formatting.rule import CellIsRule
    col_alc = L(7 + len(dias))
    ld.conditional_formatting.add(f'{col_alc}3:{col_alc}{uf}', CellIsRule(
        operator='lessThanOrEqual', formula=['3'], fill=PatternFill('solid', fgColor='F8D7D7'),
        font=Font(bold=True, color='9C1C1C')))
    if ref is not None:
        col_alc2 = L(11 + len(dias))
        ld.conditional_formatting.add(f'{col_alc2}3:{col_alc2}{uf}', CellIsRule(
            operator='lessThanOrEqual', formula=['3'], fill=PatternFill('solid', fgColor='F8D7D7'),
            font=Font(bold=True, color='9C1C1C')))
for c, w in zip(range(1, len(cols) + 1), [6, 30] + [10] * len(dias) + [8, 10, 9]
                + ([9, 10] if stock is not None else []) + ([10, 9, 10, 12] if ref is not None else []) + [7]):
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
cab(sv, 1, ['Local físico', '¿Tiene ventas de otros títulos de la Editorial en Ventas Agrupadas?']
    + ([f'Le quedan Sorpresa ({fecha_sii})'] if stock is not None else [])
    + (['Refuerzo 07/10', 'Falta llegar', 'Con lo que falta'] if ref is not None else []))
for i, s in enumerate(sin_venta, start=2):
    sv.cell(i, 1, s).border = caja
    txt = ('Sí' if s in otra_editorial else 'No') if otra_editorial is not None else 's/d (no se pasó Ventas Agrupadas)'
    sv.cell(i, 2, txt).border = caja
    if stock is not None:
        c = sv.cell(i, 3, int(stock.loc[s, 'Total'])); c.border = caja
        c.alignment = Alignment(horizontal='center')
    if ref is not None:
        for j, v in ((4, int(ref_full.loc[s, 'Total'])), (5, int(pend.loc[s, 'Total'])),
                     (6, int(con_ref.loc[s, 'Total']))):
            c = sv.cell(i, j, v); c.border = caja; c.alignment = Alignment(horizontal='center')
sv.column_dimensions['A'].width = 30
sv.column_dimensions['B'].width = 55
sv.column_dimensions['C'].width = 22
sv.column_dimensions['D'].width = 14
sv.column_dimensions['E'].width = 14
sv.column_dimensions['F'].width = 14

# ---------- hoja Stock x titulo ----------
if stock is not None:
    tc_s = [c for c in stock.columns if c != 'Total']   # tambien los que todavia no vendieron
    sk = wb.create_sheet('Le quedan x título', 3)
    sk.freeze_panes = 'C3'
    sk['A1'] = (f'Stock de cada local en el SII del {fecha_sii}, por título · mismo orden que Local x día; '
                'al final los locales sin venta · abajo, el total de la cadena y para cuántos días alcanza: por título, '
                'al ritmo desde su primera venta; el total, al ritmo de los últimos 7 días completos')
    sk['A1'].font = Font(italic=True, size=9, color=TINTA2)
    cab(sk, 2, ['Nro', 'Local'] + tc_s + ['Total'])
    filas = [(n, loc) for n, loc in mat.index] + [('', s) for s in sin_venta]
    for i, (n, loc) in enumerate(filas, start=3):
        for j, v in enumerate([n, loc] + [int(stock.loc[loc, t]) for t in tc_s]
                              + [int(stock.loc[loc, 'Total'])], start=1):
            c = sk.cell(i, j, v); c.border = caja
            c.alignment = Alignment(horizontal='left' if j == 2 else 'center')
        sk.cell(i, 3 + len(tc_s)).font = Font(bold=True)
    uf3 = 2 + len(filas)
    calor(sk, f'C3:{L(2 + len(tc_s))}{uf3}')
    dias_llenos = [x for x in dias if x != parcial]
    vend = det[det.dia.isin(dias_llenos)].groupby('titulo').Cantidad.sum()
    ritmo_t = [vend_d.get(t, 0) for t in tc_s]
    ult7 = dias_llenos[-7:]                               # total: al ritmo de los ultimos 7 dias completos
    ritmo_tot = det[det.dia.isin(ult7)].Cantidad.sum() / len(ult7)
    resumen = [('Le queda en la cadena', [int(stock[t].sum()) for t in tc_s] + [int(stock.Total.sum())], '#,##0'),
               (f'Vendido hasta el {dias_llenos[-1]:%d/%m}',
                [int(vend.get(t, 0)) for t in tc_s] + [int(vend.sum())], '#,##0'),
               ('Desde (primera venta)', [lanz[t] if t in lanz.index else None for t in tc_s] + [None], 'DD/MM'),
               (f'Vendido por día desde que salió (con {MIN_DIAS}+ días)', [round(r, 1) if r else None for r in ritmo_t] + [round(ritmo_tot, 1)], '0.0'),
               ('Le alcanza (días)', [round(stock[t].sum() / r, 1) if r > 0 else None for t, r in zip(tc_s, ritmo_t)]
                + [round(stock.Total.sum() / ritmo_tot, 1)], '0.0')]
    if ref is not None:
        por_dia_t = [vend.get(t, 0) / len(dias_llenos) for t in tc_s] + [vend.sum() / len(dias_llenos)]
        resumen += [('Refuerzo 07/10', [int(ref_full[t].sum()) for t in tc_s] + [int(ref_full.Total.sum())], '#,##0'),
                    ('Falta llegar', [int(pend[t].sum()) for t in tc_s] + [int(pend.Total.sum())], '#,##0'),
                    ('Le quedan + lo que falta', [int(con_ref[t].sum()) for t in tc_s] + [int(con_ref.Total.sum())], '#,##0'),
                    ('Le alcanza con eso (días)',
                     [round(v / d, 1) if d else None for v, d in
                      zip([con_ref[t].sum() for t in tc_s] + [con_ref.Total.sum()], por_dia_t)], '0.0')]
    for k, (lab, vals, fmt) in enumerate(resumen):
        r = uf3 + 2 + k
        sk.cell(r, 2, lab).font = Font(bold=True)
        for j, v in enumerate(vals, start=3):
            c = sk.cell(r, j, v); c.font = Font(bold=True); c.border = caja
            c.number_format = fmt; c.alignment = Alignment(horizontal='center')
    from openpyxl.formatting.rule import CellIsRule
    sk.conditional_formatting.add(f'C{uf3 + 6}:{L(3 + len(tc_s))}{uf3 + 6}', CellIsRule(
        operator='lessThanOrEqual', formula=['3'], fill=PatternFill('solid', fgColor='F8D7D7'),
        font=Font(bold=True, color='9C1C1C')))
    if ref is not None:
        sk.conditional_formatting.add(f'C{uf3 + 9}:{L(3 + len(tc_s))}{uf3 + 9}', CellIsRule(
            operator='lessThanOrEqual', formula=['3'], fill=PatternFill('solid', fgColor='F8D7D7'),
            font=Font(bold=True, color='9C1C1C')))
    for c, w in zip(range(1, len(tc_s) + 4), [6, 30] + [13] * len(tc_s) + [8]):
        sk.column_dimensions[L(c)].width = w
    sk.column_dimensions['B'].width = 32 if ref is not None else 30
    sk.page_setup.orientation = 'landscape'
    sk.sheet_properties.pageSetUpPr.fitToPage = True
    sk.page_setup.fitToWidth = 1
    sk.page_setup.fitToHeight = 0
    cob_tit = dict(zip(tc_s, resumen[4][1][:-1]))
    cob_tot = resumen[4][1][-1]

# ---------- hoja Refuerzo ----------
if ref is not None:
    rf = wb.create_sheet('Refuerzo 07-10', 4)
    rf.freeze_panes = 'D3'
    rf['A1'] = ('Refuerzo cargado el 07/10 (195 de cada título) · nro de pedido según el TXT anotado · '
                'primero AMBA y después interior, como se cargó'
                + (f' · "Entró" = lo que muestra el SII del {fecha_sii} contra el del {fecha_base} descontando lo vendido '
                   '(estimado, puede mezclarse con otros envíos)' if entro is not None else
                   ' · no hay SII anterior para comparar: se toma que no llegó nada')
                + ' · "Con lo que falta" = le quedan + falta llegar')
    rf['A1'].font = Font(italic=True, size=9, color=TINTA2)
    cab(rf, 2, ['Zona', 'Local', 'Pedido'] + tit_cols + ['Total', 'Entró', 'Falta llegar', 'Estado',
                                                          f'Le quedan ({fecha_sii})', 'Con lo que falta'])
    orden_r = sorted(ref.index, key=lambda l: (zona_de[l] != 'AMBA', l))
    for i, loc in enumerate(orden_r, start=3):
        vals = [zona_de[loc], loc, pedidos.get(loc, 'SIN NRO')] + [int(ref.loc[loc, t]) or None for t in tit_cols] \
            + [int(ref.loc[loc, 'Total']), int(entro.loc[loc, 'Total']) if entro is not None else None,
               int(pend.loc[loc, 'Total']),
               ('Llegó' if pend.loc[loc, 'Total'] == 0 else 'No llegó' if entro is None or entro.loc[loc, 'Total'] == 0
                else 'Parcial'),
               int(stock.loc[loc, 'Total']), int(con_ref.loc[loc, 'Total'])]
        for j, v in enumerate(vals, start=1):
            c = rf.cell(i, j, v); c.border = caja
            c.alignment = Alignment(horizontal='left' if j == 2 else 'center')
        rf.cell(i, 4 + len(tit_cols)).font = Font(bold=True)
        est = rf.cell(i, 7 + len(tit_cols))
        est.fill = PatternFill('solid', fgColor={'Llegó': 'DFF0D8', 'Parcial': 'FCF3CF'}.get(est.value, 'F8D7D7'))
        if loc not in pedidos:
            rf.cell(i, 3).font = Font(bold=True, color='9C1C1C')
            rf.cell(i, 3).fill = PatternFill('solid', fgColor='F8D7D7')
    uf4 = 2 + len(orden_r)
    rf.cell(uf4 + 1, 2, 'TOTAL').font = Font(bold=True)
    for j in [k for k in range(4, 10 + len(tit_cols)) if k != 7 + len(tit_cols)]:
        c = rf.cell(uf4 + 1, j, f'=SUM({L(j)}3:{L(j)}{uf4})')
        c.font = Font(bold=True); c.border = caja; c.alignment = Alignment(horizontal='center')
    calor(rf, f'D3:{L(3 + len(tit_cols))}{uf4}')
    for c, w in zip(range(1, len(tit_cols) + 10), [10, 30, 12] + [15] * len(tit_cols) + [8, 8, 9, 10, 13, 12]):
        rf.column_dimensions[L(c)].width = w
    rf.print_title_rows = '2:2'
    rf.page_setup.orientation = 'landscape'
    rf.sheet_properties.pageSetUpPr.fitToPage = True
    rf.page_setup.fitToWidth = 1
    rf.page_setup.fitToHeight = 0

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
tb.print_area = f'A1:P{r0 + 1 + len(lineas)}'
for nombre in ('Local x día', 'Local x título', 'Le quedan x título'):
    if nombre in wb.sheetnames:
        wb[nombre].print_title_rows = '2:2'     # encabezado en cada pagina impresa
wb.save(salida)
if ref is not None:
    if entro is not None:
        print(f'entro {int(entro.Total.sum())} | falta llegar {int(pend.Total.sum())} | otras entradas {otras[otras > 2].to_dict()}')
    print(f'refuerzo {int(ref.Total.sum())} u | {len(ref)} locales | por titulo {ref.drop(columns="Total").sum().to_dict()} '
          f'| sin nro de pedido: {sin_pedido}')
print(f'{tot_u} u | $ {tot_ars:,} | {n_locales} locales de {n_padron} | {len(dias)} dias | '
      f'devol {len(devol)} | sin venta {sin_venta} | ' + (f'cuadre agrupadas OK ({len(cuadre)} celdas)' if agr is not None else 'SIN cuadre (no hay agrupadas)'))
print('por dia', {etq[x]: int(v) for x, v in por_dia.items()})
print('por titulo', por_tit.to_dict())
for t in lineas: print(' -', t)
