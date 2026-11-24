"""Repo El Rincón Cultural (22330, RODRIGUEZ) para SILOMA / Servicio de Novedades (20/11 y 24/11/2026).
Formato ampliación: solo Hoja1, ISBN (formato 0) + cantidad, sin encabezado, desde la fila 7.
Decisiones de Franco (24/11): Resetea va la edición vieja 9789500212687 como la pidieron; Preguntón 3 de cada
uno de los 7; Teo y Ana 3 de cada uno de los 9; "promesa de sangre" venía con el ISBN de La memoria del mal -> 9789500218160.
Uso: python -P rincon_22330.py PLANILLA.xlsx SII.xls SALIDA.xlsx"""
import sys, openpyxl, pandas as pd
from openpyxl import Workbook
fpl, fsii, fout = sys.argv[1:4]
s = pd.read_excel(fsii); s = s[s['ID ARTICULO'].notna()].copy(); s['n'] = pd.to_numeric(s.ISBN, errors='coerce')
by = s.set_index('n')
lin = []
for r in openpyxl.load_workbook(fpl).active.iter_rows(min_row=2, values_only=True):
    isbn, tit, q = r[:3]
    t = str(tit).upper()
    if isbn is None and 'PREGUNTON' in t:
        lin += [(int(i), 3, x) for i, x in s[s.TITULO.astype(str).str.startswith('PREGUNTON')][['n', 'TITULO']].values]
    elif isbn is None and 'TEO Y ANA' in t:
        lin += [(int(i), 3, x) for i, x in s[s.TITULO.astype(str).str.startswith('TEO Y ANA')][['n', 'TITULO']].values]
    elif t.strip() == 'PROMESA DE SANGRE':
        assert int(isbn) == 9789500214193; lin.append((9789500218160, int(q), tit))
    else:
        lin.append((int(isbn), int(q), tit))
assert len({i for i, *_ in lin}) == len(lin), 'ISBN repetido'
assert all(i in by.index for i, *_ in lin)
wb = Workbook(); ws = wb.active; ws.title = 'Hoja1'
for k, (i, q, _) in enumerate(lin, 7):
    ws.cell(k, 1, i).number_format = '0'; ws.cell(k, 2, q)
wb.save(fout)
print(len(lin), 'líneas', sum(q for _, q, _ in lin), 'u')
for i, q, t in lin:
    print(f'{i} x{q}  ID {int(by.loc[i, "ID ARTICULO"])}  disp {int(by.loc[i, "DISPONIBLE"]):>5}  {by.loc[i, "TITULO"]}')
