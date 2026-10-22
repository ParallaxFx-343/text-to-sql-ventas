"""Deja en cada sábana consolidada de vendedor solo las 6 reimpresiones de noviembre 2026 (por ISBN exacto),
ordenadas por cliente, sucursal y título. Mismas columnas y formato del original.
Uso: python -P sabanas_reimpresiones.py SALIDA/ SABANA1.xlsx [SABANA2.xlsx ...]"""
import sys, os, re, copy, openpyxl
REIMP = {'9789500217859': 'VENDER ES PODER', '9789500213400': 'TU AMIGO INVISIBLE', '9789500216845': 'CUENTOS CORTOS DE NAVIDAD',
         '9789500215015': 'MEDITACIONES', '9789500216593': 'ANSIEDAD', '9789500213998': 'LA VACA NO SE CALLA (ED. RÚSTICA)'}
out = sys.argv[1]; os.makedirs(out, exist_ok=True)
for f in sys.argv[2:]:
    wb = openpyxl.load_workbook(f); ws = wb.worksheets[0]
    h = [c.value for c in ws[1]]; iI, iC, iS, iT = (h.index(x) for x in ('ISBN', 'CLIENTE', 'SUCURSAL', 'TITULO'))
    n0 = ws.max_row
    keep = [r for r in range(2, n0 + 1) if str(ws.cell(r, iI + 1).value).strip() in REIMP]
    data = [[(c.value, copy.copy(c._style)) for c in ws[r]] for r in keep]
    data.sort(key=lambda row: (str(row[iC][0]), str(row[iS][0]), str(row[iT][0])))
    for k, row in enumerate(data, 2):
        for j, (v, st) in enumerate(row, 1):
            c = ws.cell(k, j); c.value = v; c._style = st
    if n0 > len(data) + 1: ws.delete_rows(len(data) + 2, n0 - len(data) - 1)
    if ws.auto_filter.ref: ws.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(h))}{len(data) + 1}'
    vend = re.search(r'([A-Z]+)_CONSOLIDADA_(\d\d-\d\d-\d{4})', os.path.basename(f))
    name = f'{vend[1]}_CONSOLIDADA_{vend[2]}_REIMPRESIONES_NOV26.xlsx'
    wb.save(os.path.join(out, name))
    tit = {}
    for row in data: tit[str(row[iI][0])] = tit.get(str(row[iI][0]), 0) + 1
    print(f'{name}: {len(data)} filas de {n0 - 1}, {len({r[0][0] for r in data})} cuentas |',
          ', '.join(f'{REIMP[i][:14]} {tit.get(i, 0)}' for i in REIMP))
