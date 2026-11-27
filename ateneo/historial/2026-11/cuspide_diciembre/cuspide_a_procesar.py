"""A PROCESAR de Cúspide (10842) a partir de la grilla 'El Ateneo - Ampliación locales', con el A PROCESAR del mes anterior como modelo.
Hoja1 (adelante, activa) = transpuesta de la grilla: una fila por local (orden de la grilla), A IdCliente 10842, B DirEnt (fila 2 de
la grilla; Castelar no trae número: 163), C 'Q', D-E DirTrans/Trans Retira (los del modelo para ese local), F NRO PEDIDO (fila 3),
G Lista, H vacía, desde la I un ID por columna (filas de artículos de la grilla) con la cantidad (vacío o ' ' = 0).
Redirigidos a Ascasubi 149: Luján 39, Gral. Rodríguez 140, Berazategui 144 → DirEnt 149 y pedido como texto 'NNN DESTINO'.
ID repetido en la grilla (ORA-00001 si un local tiene cantidad en las dos filas): una sola columna con la SUMA (criterio Franco 27/11).
Formato: copia el de la Hoja1 del modelo (encabezado, filas, anchos). Uso:
python -P cuspide_a_procesar.py GRILLA.xlsx MODELO.xlsx LISTA SALIDA.xlsx"""
import sys, copy, openpyxl
from collections import OrderedDict
from openpyxl.utils import get_column_letter as L
fg, fm, lista, fout = sys.argv[1:5]; lista = int(lista)
REDIR = {39: 'LUJAN', 140: 'GENERAL RODRIGUEZ', 144: 'BERAZATEGUI'}
CASTELAR = 163
wb = openpyxl.load_workbook(fg); g = wb['El Ateneo - Ampliación locales']
mw = openpyxl.load_workbook(fm); mh = mw['Hoja1']; mg = mw['El Ateneo - Ampliación locales']
def locales(ws):
    out = []
    for c in range(4, ws.max_column + 1):
        n, p, nom = ws.cell(2, c).value, ws.cell(3, c).value, str(ws.cell(4, c).value or '').strip()
        if isinstance(n, (int, float)): out.append((c, int(n), p, nom))
        elif p not in (None, '') and str(p).strip() and 'CASTELAR' in nom.upper(): out.append((c, CASTELAR, p, nom))
        else: assert p in (None, '') or not str(p).strip(), (L(c), n, p, nom)
    return out
# DirTrans/Trans Retira por local según el modelo (clave = número propio del local en la grilla del modelo)
mloc = locales(mg); trans = {}
for k, (c, n, p, nom) in enumerate(mloc):
    r = 2 + k; assert str(mh.cell(r, 6).value).split()[0] == str(p), (r, p)
    trans[n] = (mh.cell(r, 4).value, mh.cell(r, 5).value)
loc = locales(g)
falta = [n for _, n, _, _ in loc if n not in trans]; assert not falta, f'locales sin DirTrans en el modelo: {falta}'
pedidos = [p for _, _, p, _ in loc]; assert len(set(pedidos)) == len(pedidos), 'pedido repetido'
arts = [r for r in range(5, g.max_row + 1) if g.cell(r, 2).value is not None]
for r in range(5, g.max_row + 1):          # nada cargado en filas sin ID
    if g.cell(r, 2).value is None:
        assert not any(isinstance(g.cell(r, c).value, (int, float)) and g.cell(r, c).value for c, *_ in loc), r
def q(v):
    if v is None or (isinstance(v, str) and not v.strip()): return 0
    assert isinstance(v, (int, float)) and v >= 0 and float(v).is_integer(), v
    return int(v)
cols = OrderedDict()                        # ID -> filas de la grilla
for r in arts: cols.setdefault(int(g.cell(r, 2).value), []).append(r)
rep = {i: rs for i, rs in cols.items() if len(rs) > 1}
h = wb.create_sheet('Hoja1', 0)
h.append(['IdCliente', 'DirEnt', 'tipocomer', 'DirTrans', 'Trans Retira', 'NRO PEDIDO', 'Lista Precio', None] + list(cols))
tot = 0
for c, n, p, nom in loc:
    de, ped = (149, f'{int(p)} {REDIR[n]}') if n in REDIR else (n, int(p))
    vals = [sum(q(g.cell(r, c).value) for r in rs) for rs in cols.values()]; tot += sum(vals)
    h.append([10842, de, 'Q', *trans[n], ped, lista, None] + vals)
# formato del modelo
mlast = mh.max_row; n_rows = len(loc) + 1; mcols = mh.max_column
rowmap = [1] + [min(i, mlast - 1) if i < n_rows else mlast for i in range(2, n_rows + 1)]
ncols = 8 + len(cols)
colmap = list(range(1, 9)) + [min(9 + k, mcols - 1) if 9 + k < ncols else mcols for k in range(len(cols))]
for i, mr in enumerate(rowmap, start=1):
    if mh.row_dimensions[mr].height: h.row_dimensions[i].height = mh.row_dimensions[mr].height
    for j, mc in enumerate(colmap, start=1):
        s, d = mh.cell(mr, mc), h.cell(i, j)     # entre libros distintos se copia cada atributo (los índices de estilo no coinciden)
        d.font, d.fill, d.border = copy.copy(s.font), copy.copy(s.fill), copy.copy(s.border)
        d.alignment, d.number_format, d.protection = copy.copy(s.alignment), s.number_format, copy.copy(s.protection)
for j, mc in enumerate(colmap, start=1):
    if mh.column_dimensions[L(mc)].width: h.column_dimensions[L(j)].width = mh.column_dimensions[L(mc)].width
wb.active = 0
for ws in wb: ws.sheet_view.tabSelected = ws.title == 'Hoja1'
wb.save(fout)
gt = sum(q(g.cell(r, c).value) for r in arts for c, *_ in loc)
print(f'{len(loc)} filas (locales) | {len(cols)} IDs ({len(arts)} filas de artículos) | {tot} u | grilla {gt} u | lista {lista}')
assert tot == gt
for i, rs in rep.items(): print(f'  ID {i} repetido en filas {rs} ({g.cell(rs[0], 3).value}): sumado en una columna')
for c, n, p, nom in loc:
    if n in REDIR or n == CASTELAR: print(f'  {nom[:35]:<35} grilla {n} -> DirEnt {149 if n in REDIR else n}, pedido {p}')
