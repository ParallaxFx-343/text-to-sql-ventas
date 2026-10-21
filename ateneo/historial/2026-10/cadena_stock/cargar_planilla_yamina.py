"""Carga una acción (un PAEC por sucursal) en la planilla 'Cadena. Stock 26.27' de Yamina
(hoja 'Reposición x local': Sucursal | Fecha | Paec | Cant. Unid.; nombre en A solo en la primera fila
de cada sucursal, un movimiento por fila; desde el 21/10 trae todo el historial).
- Movimiento nuevo: se inserta una fila al final del bloque de la sucursal, con el formato de esa fila.
- Si ya hay una fila con la misma fecha y unidades: no se duplica; si le falta el PAEC, se completa.
- Nunca se pisa un PAEC distinto ya cargado (avisa).
Uso: python -P cargar_planilla_yamina.py ENTRADA.xlsx SALIDA.xlsx dd/mm/aaaa DETALLE.txt [PAECS.txt]"""
import sys, re, copy, datetime as dt, openpyxl
fin, fout, fecha, det = sys.argv[1:5]
paecs = {}
if len(sys.argv) > 5:
    paecs = {m[1].strip(): '25-' + m[2] for m in (re.match(r'(.+?):\s*25[ -]?(\d{5})\s*$', l) for l in open(sys.argv[5], encoding='utf-8-sig')) if m}
f = dt.datetime.strptime(fecha, '%d/%m/%Y')
u = {}
for l in open(det, encoding='utf-8-sig'):
    m = re.match(r'^([A-Z0-9][^=]*?)\s{2,}(\d+) u\s+\d+ títulos', l)
    if m: u[m[1].strip()] = int(m[2])
ALIAS = {'QUILMES CENTRO': 'QUILMES', 'REMEROS (NORDELTA)': 'REMEROS', 'MAR DEL PLATA M. BENDU': 'BENDÚ',
         'MAR DEL PLATA LOS GALLEGOS': 'LOS GALLEGOS', 'MAR DEL PLATA PASEO ALDREY': 'MAR DEL PLATA SHOP'}
wb = openpyxl.load_workbook(fin); ws = wb['Reposición x local']
assert [ws.cell(1, c).value for c in range(1, 5)] == ['Sucursal', 'Fecha', 'Paec', 'Cant. Unid.']
NC = ws.max_column
def bloques():
    b = {}; cur = None
    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value; vals = [ws.cell(r, c).value for c in (2, 3, 4)]
        if a: cur = a.strip(); assert cur not in b, cur; b[cur] = [r, r]
        elif all(v is None for v in vals): cur = None
        elif cur: b[cur][1] = r
    return b
# de abajo hacia arriba para que las inserciones no corran las filas pendientes
pos = bloques()
for suc in sorted(u, key=lambda s: -pos[ALIAS.get(s, s)][0]):
    nom = ALIAS.get(suc, suc); r0, r1 = pos[nom]; q = u[suc]; p = paecs.get(suc)
    hit = [r for r in range(r0, r1 + 1) if ws.cell(r, 2).value == f and ws.cell(r, 4).value == q]
    if hit:
        r = hit[0]; c = ws.cell(r, 3).value
        if c is None and p: ws.cell(r, 3).value = p; print(f'{nom:<25} fila {r}: PAEC {p} completado')
        elif c not in (None, p): print(f'{nom:<25} fila {r}: YA TIENE OTRO PAEC ({c}), no se toca')
        else: print(f'{nom:<25} fila {r}: ya estaba')
        continue
    vacia = all(ws.cell(r0, c).value is None for c in (2, 3, 4))
    if vacia: r = r0
    else:
        r = r1 + 1; ws.insert_rows(r)
        for c in range(1, NC + 1): ws.cell(r, c)._style = copy.copy(ws.cell(r1, c)._style)
    for c, v in ((2, f), (3, p), (4, q)): ws.cell(r, c).value = v
    print(f'{nom:<25} fila {r}: {fecha} {p or "(sin PAEC todavía)"} {q} u')
wb.save(fout)
