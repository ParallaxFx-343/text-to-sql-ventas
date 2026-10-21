"""Pasa el historial completo de movimientos (nuestra Cadena_Stock_26.27_ACTUALIZADO.xlsx, que tiene Concepto)
a la planilla de Yamina (formato 21/10: Sucursal | Fecha | Paec | Cant. Unid., una fila por sucursal).
Cada movimiento va en una fila: el primero en la fila de la sucursal y los demás en filas insertadas debajo,
con el mismo formato. No se agregan columnas. Lo que la planilla ya tenía se respeta y no se duplica.
Las cargas propias de la planilla que no estén en el historial (p. ej. la repo del día) van al final de su sucursal.
Las sucursales que faltan (DEVOTO) se agregan al final con el formato de la última fila.
Uso: python -P historial_a_planilla_yamina.py ACTUALIZADO.xlsx PLANILLA_YAMINA.xlsx SALIDA.xlsx"""
import sys, copy, openpyxl
fa, fy, fout = sys.argv[1:4]
a = openpyxl.load_workbook(fa)['Reposición x local']
mov = {}; cur = None
for r in a.iter_rows(min_row=2, values_only=True):
    if r[0] and str(r[0]).startswith('GRUPO'): cur = None; continue
    if r[0]: cur = r[0]; mov[cur] = []
    if cur and any(v is not None for v in r[2:5]): mov[cur].append(tuple(r[2:5]))
wb = openpyxl.load_workbook(fy); ws = wb['Reposición x local']
NC = ws.max_column
def estilo(dst, src):
    for c in range(1, NC + 1):
        ws.cell(dst, c)._style = copy.copy(ws.cell(src, c)._style)
fmt = {c: ws.cell(2, c).number_format for c in (2, 3, 4)}   # fila con datos de Yamina (Grand Splendid)
filas = [(r, ws.cell(r, 1).value) for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value]
nombres = [n for _, n in filas]
assert len(set(nombres)) == len(nombres)
faltan = [n for n in mov if n not in nombres]
# DEVOTO (u otras) al final
for n in faltan:
    r = ws.max_row + 1; estilo(r, ws.max_row); ws.cell(r, 1).value = n; filas.append((r, n))
log = {}
for r, n in sorted(filas, reverse=True):          # de abajo hacia arriba, así no se corren las filas pendientes
    propio = tuple(ws.cell(r, c).value for c in (2, 3, 4))
    L = list(mov.get(n, []))
    if any(v is not None for v in propio):
        if propio not in L:
            # misma fecha y unidades que un movimiento del historial pero sin PAEC: es el mismo, no se duplica
            if not any(m[0] == propio[0] and m[2] == propio[2] for m in L): L.append(propio)
    if not L: continue
    for c, v in zip((2, 3, 4), L[0]): ws.cell(r, c).value = v
    if len(L) > 1:
        ws.insert_rows(r + 1, len(L) - 1)
        for k, m in enumerate(L[1:], 1):
            estilo(r + k, r); ws.cell(r + k, 1).value = None
            for c, v in zip((2, 3, 4), m): ws.cell(r + k, c).value = v
    for k in range(len(L)):
        for c in (2, 3, 4): ws.cell(r + k, c).number_format = fmt[c]
    log[n] = len(L)
wb.save(fout)
print('sucursales', len(log), 'filas de movimientos', sum(log.values()), '| agregadas al final:', faltan)
