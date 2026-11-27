"""Arma el A PROCESAR de Servicio de Novedades de un vendedor siguiendo el modelo de noviembre (Egisti):
primera hoja 'Hoja4' = carga: fila 1 encabezado (IdCliente, DirEnt, tipocomer, DirTrans, Trans Retira, Observacion, Lista Precio, vacía,
IDs desde la I), una fila por cliente de la Hoja1 (sin la RAZON SOCIAL), Observacion = la de la Hoja1, cantidades vacías = 0.
Los IDs salen por ISBN de la LP Argentina (novedades) / SII; también se escriben en la fila 5 de la Hoja1.
Uso: python -P a_procesar.py GRILLA.xlsx LP_ARGENTINA.xlsx LISTA SALIDA.xlsx"""
import sys, openpyxl, pandas as pd
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
wb.active = 0
for ws in wb: ws.sheet_view.tabSelected = ws.title == 'Hoja4'
wb.save(fout)
print(len(filas), 'filas,', len(ids), 'IDs, lista', lista, '| unidades', sum(tot))
for c, i, t in zip(cols, ids, tot):
    print(f'  {L(c):>3} {i}  {t:>4} u  cuota {h1.cell(filas[-1] + 2, c).value}  {h1.cell(2, c).value}')
