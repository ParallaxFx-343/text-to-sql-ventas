"""Carga una acción (un PAEC por sucursal) en la planilla 'Cadena. Stock 26.27' de Yamina (formato 21/10:
hoja 'Reposición x local' con una fila por sucursal: Sucursal | Fecha | Paec | Cant. Unid.).
No pisa filas que ya tienen Paec o Fecha y no suma repetidos: si la sucursal ya tiene datos, avisa y no escribe.
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
filas = {}
for r in range(2, ws.max_row + 1):
    a = ws.cell(r, 1).value
    if a: filas.setdefault(a.strip(), []).append(r)
modelo = 2   # fila con datos cargados (Grand Splendid): de ahí sale el formato de Fecha/Paec/Cant.
for suc, q in u.items():
    nom = ALIAS.get(suc, suc)
    assert nom in filas and len(filas[nom]) == 1, (nom, filas.get(nom))
    r = filas[nom][0]
    b, c_, d = (ws.cell(r, c).value for c in (2, 3, 4))
    if (b, d) == (f, q) and c_ is None and paecs.get(suc):     # misma carga sin PAEC: solo se completa el PAEC
        pass
    elif any(v is not None for v in (b, c_, d)):
        print('YA TIENE DATOS, no se toca:', nom, [b, c_, d]); continue
    for c, v in ((2, f), (3, paecs.get(suc)), (4, q)):
        ws.cell(r, c).value = v
        ws.cell(r, c).number_format = ws.cell(modelo, c).number_format
    print(f'{nom:<25} fila {r}: {fecha} {paecs.get(suc) or "(sin PAEC todavía)"} {q} u')
wb.save(fout)
