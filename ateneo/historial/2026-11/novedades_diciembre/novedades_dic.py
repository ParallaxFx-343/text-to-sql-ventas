"""Servicio de Novedades diciembre 2026: escribe títulos (fila 2), ISBN (fila 3) y cuotas de cada vendedor en su grilla,
y en la grilla de Cúspide (ISBN en A, título en C, cuota en la columna "Cuotas").
Sale todo de la planilla "Cuotas novedades DICIEMBRE 2026": novedades = columnas con TIRADA y cuotas (C..J),
reimpresiones = las que siguen (K..P, sin cuotas). Sin IDs (Franco no los tiene todavía).
Reglas: solo se escriben valores en celdas VACÍAS; nunca se toca una celda con valor (en particular los 0 en gris de los
clientes que no reciben importados); no se tocan formatos, colores, merges, filas ni columnas. Mismo nombre de archivo.
Uso: python -P novedades_dic.py CARPETA_ENTRADA CARPETA_SALIDA"""
import sys, os, glob, openpyxl
from openpyxl.utils import get_column_letter as L
src, out = sys.argv[1:3]; os.makedirs(out, exist_ok=True)
cf = openpyxl.load_workbook(glob.glob(os.path.join(src, 'Cuotas novedades*'))[0])['CUOTAS']
reim0 = next(c for c in range(3, 30) if cf.cell(1, c).value == 'REIMPRESIONES')
cols = [c for c in range(3, cf.max_column + 1) if cf.cell(3, c).value]
NOV = [(cf.cell(2, c).value.strip().upper(), int(cf.cell(3, c).value), c) for c in cols if c < reim0]
REI = [(cf.cell(2, c).value.strip().upper(), int(cf.cell(3, c).value), c) for c in cols if c >= reim0]
CUO = {cf.cell(r, 2).value.strip(): {c: cf.cell(r, c).value for _, _, c in NOV} for r in range(6, 17) if cf.cell(r, 2).value}
GRILLAS = [('EGISTI', 'Hoja1', [(67, 'EGISTI, CLAUDIO')]),
           ('FUCITO SERVICIO', 'Hoja1', [(55, 'FUCITO  generales'), (56, 'FUCITO  Cúspide'), (57, 'FUCITO  Queleer'), (58, 'FUCITO  Santa Fe')]),
           ('FUCITO Queleer', 'Hoja1', [(20, 'FUCITO  Queleer')]),
           ('FUCITO Santa Fe', 'GRILLA', [(13, 'FUCITO  Santa Fe')]),
           ('HOUNDAKJIAN', 'Hoja1', [(29, 'HOUNDAKJIAN, PABLO')]),
           ('MARTIRENA', 'Hoja1', [(60, 'MARTIRENA, EZEQUIEL')]),
           ('NAVARRO', 'Hoja1', [(52, 'NAVARRO, GONZALO')]),
           ('RODRIGUEZ', 'Hoja1', [(34, 'RODRIGUEZ, GONZALO')]),
           ('SALICE', 'Hoja1', [(29, 'SÁLICE, MARCELO')])]
avisos = []
def put(ws, r, c, v, quien):
    cell = ws.cell(r, c)
    if v is None: return
    if cell.value is not None:
        avisos.append(f'{quien}: {cell.coordinate} ya tenía {cell.value!r}, NO se pisó (iba {v!r})'); return
    cell.value = v
for pref, hoja, filas in GRILLAS:
    f = glob.glob(os.path.join(src, f'{pref}*DICIEMBRE 26.xlsx')); assert len(f) == 1, (pref, f)
    wb = openpyxl.load_workbook(f[0]); ws = wb[hoja]
    assert ws.cell(1, 10).value == 'NOVEDADES' and ws.cell(1, 22).value == 'REIMPRESIONES', pref   # J y V
    assert len(NOV) <= 12 and len(REI) <= 17
    cn = {c: 10 + k for k, (_, _, c) in enumerate(NOV)}                 # J..
    for k, (t, i, c) in enumerate(NOV): put(ws, 2, 10 + k, t, pref); put(ws, 3, 10 + k, i, pref)
    for k, (t, i, c) in enumerate(REI): put(ws, 2, 22 + k, t, pref); put(ws, 3, 22 + k, i, pref)
    for r, clave in filas:
        lab = ws.cell(r, 9).value
        assert lab is None or lab.strip() == clave.strip(), (pref, r, lab)
        assert ws.cell(r, 10).number_format == '0' and (ws.cell(r, 10).border.top.style or ws.cell(r, 10).border.bottom.style), (pref, r)   # la fila de cuotas
        for c, v in CUO[clave].items(): put(ws, r, cn[c], v, f'{pref} cuotas')
    wb.save(os.path.join(out, os.path.basename(f[0])))
    print(f'{pref:<16} títulos/ISBN J-{L(9 + len(NOV))} y V-{L(21 + len(REI))} | cuotas fila(s) {[r for r, _ in filas]}')
f = glob.glob(os.path.join(src, '*CUSPIDE GRILLA*'))[0]; wb = openpyxl.load_workbook(f); ws = wb['El Ateneo - Ampliación locales']
cq = next(c for c in range(1, ws.max_column + 1) if str(ws.cell(4, c).value).strip().lower() == 'cuotas')
assert [ws.cell(4, c).value for c in (1, 3)] == ['ISBN 13', 'Autor y Titulo']
for k, (t, i, c) in enumerate(NOV):
    put(ws, 5 + k, 1, i, 'CUSPIDE'); put(ws, 5 + k, 3, t, 'CUSPIDE'); put(ws, 5 + k, cq, CUO['FUCITO  Cúspide'][c], 'CUSPIDE')
for k, (t, i, c) in enumerate(REI):
    put(ws, 17 + k, 1, i, 'CUSPIDE'); put(ws, 17 + k, 3, t, 'CUSPIDE')
wb.save(os.path.join(out, os.path.basename(f)))
print(f'CUSPIDE          novedades filas 5-{4 + len(NOV)} (cuota en {L(cq)}), reimpresiones filas 17-{16 + len(REI)}')
print('Novedades:', [(t, i) for t, i, _ in NOV]); print('Reimpresiones:', [(t, i) for t, i, _ in REI])
for a in avisos: print('AVISO', a)
