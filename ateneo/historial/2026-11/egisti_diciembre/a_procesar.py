"""Arma el A PROCESAR de Servicio de Novedades de un vendedor siguiendo el modelo de noviembre (Egisti):
primera hoja 'Hoja4' = carga: fila 1 encabezado (IdCliente, DirEnt, tipocomer, DirTrans, Trans Retira, Observacion, Lista Precio, vacía,
IDs desde la I), una fila por cliente de la Hoja1 (sin la RAZON SOCIAL), Observacion = la de la Hoja1, cantidades vacías = 0.
Los IDs salen por ISBN de la LP Argentina (novedades) / SII; también se escriben en la fila 5 de la Hoja1.
Con MODELO (el A PROCESAR del mes anterior) copia el formato de su Hoja4: colores, negrita, bordes, anchos y altos; la última
novedad, la última reimpresión y la última fila toman el formato de las últimas del modelo (bordes de cierre).
Uso: python -P a_procesar.py GRILLA.xlsx LP_ARGENTINA.xlsx LISTA SALIDA.xlsx [MODELO.xlsx]"""
import sys, copy, openpyxl, pandas as pd
from openpyxl.utils import get_column_letter as L
fg, fa, lista, fout = sys.argv[1:5]; lista = int(lista)
a = pd.read_excel(fa, header=2); a = a[a['I.D.'].notna()]
ID = dict(zip(a['I.S.B.N.'].astype('int64'), a['I.D.'].astype(int)))
wb = openpyxl.load_workbook(fg); h1 = wb['Hoja1']
assert h1.cell(5, 1).value == 'IdCliente' and h1.cell(5, 3).value == 'DirEnt'
cols = [c for c in range(10, h1.max_column + 1) if h1.cell(3, c).value]
ids = [ID[int(h1.cell(3, c).value)] for c in cols]
assert len(set(ids)) == len(ids)
filas = [r for r in range(6, 66) if h1.cell(r, 1).value]
assert filas == list(range(6, filas[-1] + 1)) and not h1.cell(filas[-1] + 1, 1).value
for c in range(10, h1.max_column + 1):              # nada cargado en columnas sin ISBN, salvo ceros
    for r in filas:
        v = h1.cell(r, c).value
        assert c in cols or v in (None, 0), (r, c, v)
for c, i in zip(cols, ids):
    h1.cell(5, c).value = i
h4 = wb.create_sheet('Hoja4', 0)
h4.append(['IdCliente', 'DirEnt', 'tipocomer', 'DirTrans', 'Trans Retira', 'Observacion ', 'Lista Precio', None] + ids)
tot = [0] * len(cols)
for r in filas:
    q = []
    for c in cols:
        v = h1.cell(r, c).value
        assert v is None or isinstance(v, (int, float)), (r, c, v)
        q.append(int(v or 0))
    tot = [x + y for x, y in zip(tot, q)]
    h4.append([h1.cell(r, 1).value, h1.cell(r, 3).value, h1.cell(r, 4).value, h1.cell(r, 5).value, h1.cell(r, 6).value,
               h1.cell(r, 7).value, lista, None] + q)
if len(sys.argv) > 5:
    mw = openpyxl.load_workbook(sys.argv[5]); mh, m1 = mw['Hoja4'], mw['Hoja1']
    def secciones(ws, c0):                       # columnas con ISBN antes y después de REIMPRESIONES (fila 1)
        rc = next(c for c in range(1, ws.max_column + 1) if str(ws.cell(1, c).value).strip().upper() == 'REIMPRESIONES')
        cs = [c for c in range(10, ws.max_column + 1) if ws.cell(3, c).value]
        return [c for c in cs if c < rc], [c for c in cs if c >= rc]
    nn, nr = (len(x) for x in secciones(h1, 10)); mn, mr = (len(x) for x in secciones(m1, 10))
    assert mh.max_column == 8 + mn + mr, 'la Hoja4 del modelo no coincide con su Hoja1'
    def mapa(n, m, base):
        return [base + (m - 1 if k == n - 1 else min(k, m - 2)) for k in range(n)]
    colmap = list(range(1, 9)) + mapa(nn, mn, 9) + mapa(nr, mr, 9 + mn)
    ultima = max(r for r in range(2, mh.max_row + 1) if mh.cell(r, 1).value)
    n = len(filas)
    rowmap = [1] + [2 + (ultima - 2 if i == n - 1 else min(i, ultima - 3)) for i in range(n)]
    for i, mr_ in enumerate(rowmap, start=1):
        if mh.row_dimensions[mr_].height: h4.row_dimensions[i].height = mh.row_dimensions[mr_].height
        for j, mc in enumerate(colmap, start=1):
            h4.cell(i, j)._style = copy.copy(mh.cell(mr_, mc)._style)
    for j, mc in enumerate(colmap, start=1):
        h4.column_dimensions[L(j)].width = mh.column_dimensions[L(mc)].width
    for cf in mh.conditional_formatting:
        for rule in cf.rules: h4.conditional_formatting.add(str(cf.sqref), rule)
    h4.sheet_format.defaultRowHeight = mh.sheet_format.defaultRowHeight
wb.active = 0
for ws in wb: ws.sheet_view.tabSelected = ws.title == 'Hoja4'
wb.save(fout)
print(len(filas), 'filas,', len(ids), 'IDs, lista', lista, '| unidades', sum(tot))
for c, i, t in zip(cols, ids, tot):
    print(f'  {L(c):>3} {i}  {t:>4} u  cuota {h1.cell(filas[-1] + 2, c).value}  {h1.cell(2, c).value}')
