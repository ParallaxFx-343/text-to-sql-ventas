"""Completa la columna Paec de las acciones de septiembre en la planilla Cadena Stock.
Fuente: mails de Franco a Yamina (11, 15 y 17/09) y PAECS_PROMO_2X1_11-09.txt."""
import os, re, sys, openpyxl
AQUI = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(AQUI, 'Cadena_Stock_26.27_ACTUALIZADO.xlsx')
TXT = os.path.join(AQUI, '..', '..', '2026-09', 'PAECS_PROMO_2X1_11-09.txt')
ALIAS = {'FLORIDA 632': 'ATENEO FLORIDA 632', 'GRAND SPLENDID': 'ATENEO GRAND SPLENDID',
         'COMODORO': 'COMODORO RIVADAVIA', 'CORRIENTES': 'CORRIENTES SHOPPING',
         'BENDU': 'BENDÚ', 'SANTA FE RIBERA': 'SANTA FE Ribera Shopping',
         'SOLAR': 'SOLAR DE LA ABADIA', 'TORTUGUITAS': 'TORTUGAS OPEN MALL'}
p2x1 = {}
for l in open(TXT, encoding='utf-8'):
    m = re.match(r'\s*(.+?):\s*25\s*(\d{5})\s*$', l)
    if m:
        p2x1[ALIAS.get(m[1], m[1])] = '25-' + m[2]
assert len(p2x1) == 33, len(p2x1)
FIJOS = {'Refuerzo Día de la Madre': '25-88006', 'Agenda + autores de octubre': '25-88034'}
wb = openpyxl.load_workbook(XLSX)
ws = wb['Reposición x local']
cur = None; n = {}; usados = set()
for row in ws.iter_rows(min_row=2):
    a, b, d = row[0].value, row[1].value, row[3]
    if a and not str(a).startswith('GRUPO'):
        cur = a
    if b == 'Promo 2x1 $28.500':
        assert d.value is None and cur in p2x1, cur
        d.value = p2x1[cur]; usados.add(cur)
    elif b in FIJOS:
        assert d.value is None, (cur, b)
        d.value = FIJOS[b]
    else:
        continue
    n[b] = n.get(b, 0) + 1
assert usados == set(p2x1), set(p2x1) - usados
print(n)
wb.save(XLSX)
