"""Saca de 'Reposición x local' los movimientos que Franco no cargó (20/10):
Repo Libros Sorpresa de DEVOTO del 05/10 y los Canales refuerzo sin PAEC (23/09, 09, 14 y 19/10)."""
import os, openpyxl
AQUI = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(AQUI, 'Cadena_Stock_26.27_ACTUALIZADO.xlsx')
wb = openpyxl.load_workbook(XLSX); ws = wb['Reposición x local']
val = lambda r: [ws.cell(r, c).value for c in range(1, 6)]
# DEVOTO: es la primera fila del bloque, se suben las de abajo para no perder el nombre ni el formato
r = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == 'DEVOTO')
assert val(r)[1:5:3] == ['Repo Libros Sorpresa', 40] and val(r)[3] is None, val(r)
for k in range(r, ws.max_row):
    for c in range(2, 6): ws.cell(k, c).value = ws.cell(k + 1, c).value
ws.delete_rows(ws.max_row)
# CANALES sin PAEC
r0 = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == 'CANALES')
borrar = []
r = r0
while r == r0 or (ws.cell(r, 1).value is None and ws.cell(r, 2).value):
    if ws.cell(r, 4).value is None:
        assert r != r0 and ws.cell(r, 2).value == 'Canales refuerzo', val(r); borrar.append(r)
    r += 1
assert len(borrar) == 4 and borrar == list(range(borrar[0], borrar[0] + 4)), borrar
for b in borrar: print('sale', val(b)[1:])
ws.delete_rows(borrar[0], 4)
for g in list(ws.merged_cells.ranges):
    if g.min_row > borrar[0]:
        ws.merged_cells.ranges.discard(g) if hasattr(ws.merged_cells.ranges, 'discard') else ws.merged_cells.ranges.remove(g)
        ws.merge_cells(start_row=g.min_row - 4, end_row=g.max_row - 4, start_column=g.min_col, end_column=g.max_col)
wb.save(XLSX)
