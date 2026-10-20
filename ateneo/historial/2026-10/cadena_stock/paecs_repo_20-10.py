"""Completa la columna Paec de 'Repo Ranking (20/10)' con los PAEC que pasó Franco."""
import os, re, openpyxl
AQUI = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(AQUI, 'Cadena_Stock_26.27_ACTUALIZADO.xlsx')
TXT = os.path.join(AQUI, '..', 'reposicion_20-10', 'PAECS_REPOSICION_20-10_cargados.txt')
ALIAS = {'QUILMES': 'QUILMES CENTRO', 'REMEROS': 'REMEROS (NORDELTA)', 'BENDÚ': 'MAR DEL PLATA M. BENDU',
         'LOS GALLEGOS': 'MAR DEL PLATA LOS GALLEGOS', 'MAR DEL PLATA SHOP': 'MAR DEL PLATA PASEO ALDREY'}
paec = {m[1]: '25-' + m[2] for m in (re.match(r'(.+?):\s*25 (\d{5})$', l.strip()) for l in open(TXT)) if m}
assert len(paec) == 9, len(paec)
wb = openpyxl.load_workbook(XLSX); ws = wb['Reposición x local']
cur = None; usados = {}
for row in ws.iter_rows(min_row=2):
    a = row[0].value
    if a and not str(a).startswith('GRUPO'):
        cur = ALIAS.get(a, a)
    if row[1].value == 'Repo Ranking (20/10)':
        assert row[3].value in (None, paec[cur]), cur
        row[3].value = paec[cur]; usados[cur] = row[4].value
assert set(usados) == set(paec), set(paec) ^ set(usados)
wb.save(XLSX)
for k in sorted(usados): print(k, paec[k], usados[k])
print('total', sum(usados.values()))
