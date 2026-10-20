"""Rearma el bloque CANALES de la hoja 'Reposición x local' con los PAEC de los mails con
Damián Winnik (01, 09, 17 y 21/09). Los pedidos del 01, 09 y 21/09 no tenían archivo:
las unidades son las pedidas en el mail (canales_mails_septiembre.csv)."""
import os, copy, csv, datetime as dt, openpyxl
AQUI = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(AQUI, 'Cadena_Stock_26.27_ACTUALIZADO.xlsx')
SEP = os.path.join(AQUI, '..', '..', '2026-09')
D = lambda d, m: dt.datetime(2026, m, d)
filas = [('Canales refuerzo', D(1, 9), '25-87894', 'canales_refuerzo_01-09.csv'),
         ('Canales refuerzo', D(9, 9), '25-87940', 'canales_refuerzo_09-09.csv'),
         ('Canales ranking', D(17, 9), '25-88029', None),
         ('Canales refuerzo', D(21, 9), '25-88068', 'canales_refuerzo_21-09_a.csv'),
         ('Canales refuerzo', D(21, 9), '25-88083', 'canales_refuerzo_21-09_b.csv'),
         ('Canales refuerzo', D(23, 9), None, None)]
def unid(f):
    return sum(int(r[1]) for r in csv.reader(open(os.path.join(SEP, f)), delimiter=';') if r)
wb = openpyxl.load_workbook(XLSX); ws = wb['Reposición x local']
r0 = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == 'CANALES')
r1 = r0
while ws.cell(r1 + 1, 1).value is None and ws.cell(r1 + 1, 2).value: r1 += 1
viejas = [tuple(ws.cell(r, c).value for c in range(2, 6)) for r in range(r0, r1 + 1)]
assert [v[:2] for v in viejas[:2]] == [('Canales ranking', D(17, 9)), ('Canales refuerzo', D(23, 9))], viejas
assert all(v[2] is None for v in viejas), viejas
nuevas = []
for c, d, p, f in filas:
    u = unid(f) if f else next(v[3] for v in viejas if v[:2] == (c, d))
    nuevas.append((c, d, p, u))
nuevas += viejas[2:]
n = len(nuevas) - len(viejas)
ws.insert_rows(r1 + 1, n)
for k in range(n):
    for c in range(1, 6):
        ws.cell(r1 + 1 + k, c)._style = copy.copy(ws.cell(r1, c)._style)
for g in list(ws.merged_cells.ranges):
    if g.min_row > r1:
        ws.merged_cells.ranges.discard(g) if hasattr(ws.merged_cells.ranges, 'discard') else ws.merged_cells.ranges.remove(g)
        ws.merge_cells(start_row=g.min_row + n, end_row=g.max_row + n, start_column=g.min_col, end_column=g.max_col)
for k, v in enumerate(nuevas):
    for c, x in zip(range(2, 6), v): ws.cell(r0 + k, c).value = x
wb.save(XLSX)
for v in nuevas: print(v[0], v[1].strftime('%d/%m'), v[2], v[3])
