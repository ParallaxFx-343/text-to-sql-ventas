"""Completa la columna Paec de 'Recomposición Octubre' (19/10) con el mail de Franco a Yamina del 19/10.
En el mail ATENEO ROSARIO decía '22 88358': es 25-88358 (va en la secuencia)."""
import os, re, openpyxl
AQUI = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(AQUI, 'Cadena_Stock_26.27_ACTUALIZADO.xlsx')
TXT = os.path.join(AQUI, '..', 'recomposicion_19-10', 'PAECS_RECOMPOSICION_19-10_cargados.txt')
ALIAS = {'QUILMES': 'QUILMES CENTRO', 'REMEROS': 'REMEROS (NORDELTA)', 'BENDÚ': 'MAR DEL PLATA M. BENDU',
         'LOS GALLEGOS': 'MAR DEL PLATA LOS GALLEGOS', 'MAR DEL PLATA SHOP': 'MAR DEL PLATA PASEO ALDREY'}
paec = {m[1]: '25-' + m[2] for m in (re.match(r'(.+?):\s*25 (\d{5})$', l.strip()) for l in open(TXT)) if m}
assert len(paec) == 56, len(paec)
wb = openpyxl.load_workbook(XLSX); ws = wb['Reposición x local']
cur = None; usados = set()
for row in ws.iter_rows(min_row=2):
    a = row[0].value
    if a and not str(a).startswith('GRUPO'):
        cur = ALIAS.get(a, a)
    if row[1].value == 'Recomposición Octubre':
        assert row[3].value in (None, paec[cur]), cur
        row[3].value = paec[cur]; usados.add(cur)
assert usados == set(paec), set(paec) ^ usados
wb.save(XLSX); print(len(usados), 'filas')
