"""Lee un .msg de liquidación de Boutique Sur (Ajamil 21653): saca la tabla del cuerpo y controla contra "Cantidad de ejemplares".
Uso: python -P leer_mail.py MAIL.msg SALIDA.tsv  (después: siloma_05-11/siloma.py SALIDA.tsv SII.xls LIQUIDACIONOCTUBRE x.xlsx)"""
import olefile, sys, re
o = olefile.OleFileIO(sys.argv[1])
b = o.openstream('__substg1.0_1000001F').read().decode('utf-16-le')
suc = re.search(r'sucursal:\s*(\w+)', b).group(1)
ej = int(re.search(r'Cantidad de ejemplares:\s*(\d+)', b).group(1))
ped = re.search(r'Pedidos:\s*(.+)', b).group(1).strip()
fechas = re.search(r'entre fechas:\s*(.+)', b).group(1).strip()
rows = []
L = b.splitlines()
for k in range(len(L) - 1):          # a veces el ISBN queda solo en una línea y el resto de la fila en la siguiente
    if re.fullmatch(r'\s*[\dXx-]{10,17}\s*', L[k]) and L[k + 1].startswith('\t'):
        L[k], L[k + 1] = L[k].strip() + L[k + 1], ''
for l in L:
    p = l.split('\t')
    if len(p) >= 6 and re.fullmatch(r'[\dXx-]{10,17}', p[0].strip()):
        rows.append((p[0].strip(), p[1].strip(), p[2].strip(), p[5].strip()))
open(sys.argv[2], 'w', encoding='utf-8').write(''.join('\t'.join(r) + '\n' for r in rows))
print(suc, '| ejemplares', ej, '| pedido', ped, '|', fechas, '| líneas', len(rows), '| suma', sum(int(r[3]) for r in rows))
print('   tablas con ISBN no leídas:', [l for l in b.splitlines() if re.match(r'\s*97[89]', l) and len(l.split('\t')) < 6])
