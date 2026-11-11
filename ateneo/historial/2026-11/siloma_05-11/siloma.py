"""Planilla de carga SILOMA desde una lista pegada (ISBN, Cód., título, cantidad).
Formato SILOMA: hoja Hoja1, filas 1-5 vacías, desde la fila 6: A = ISBN entero (formato 0), B = cantidad, C = etiqueta.
ISBN: el Cód. si viene; si no, el ISBN sin guiones, resuelto con isbn.resolver (ISBN-10 viejo -> 978 + verificador nuevo).
Uso: python -P siloma.py LISTA.tsv SII.xls ETIQUETA SALIDA.xlsx"""
import sys, pandas as pd
from pathlib import Path
from openpyxl import Workbook
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from isbn import resolver
flista, fsii, etq, fout = sys.argv[1:5]
s = pd.read_excel(fsii); s = s[s['ID ARTICULO'].notna()].copy(); s['n'] = pd.to_numeric(s.ISBN, errors='coerce'); by = s.set_index('n')
filas, avisos = [], []
for k, l in enumerate(open(flista, encoding='utf-8'), 1):
    isbn_txt, cod, tit, cant = l.rstrip('\n').split('\t')
    a, ma = resolver(isbn_txt.replace('-', '').strip())
    cod = cod.strip()
    idcod = None
    if cod.isdigit() and len(cod) <= 7:          # a veces la columna Cód. trae el ID de artículo, no el ISBN
        idcod, cod = int(cod), '-'
    elif cod not in ('', '-') and not (cod.isdigit() and len(cod) == 13):   # código interno del cliente (A1862R...)
        avisos.append(f'{tit}: Cód. {cod} es un código del cliente, se usa el ISBN'); cod = '-'
    i = cod if cod not in ('', '-') else a
    if idcod and int(i) in by.index and int(by.loc[int(i), 'ID ARTICULO']) != idcod:
        avisos.append(f'{tit}: Cód. {idcod} pero el ISBN {i} es el ID {int(by.loc[int(i), "ID ARTICULO"])}')
    if cod not in ('', '-') and a != cod:
        avisos.append(f'{tit}: ISBN {isbn_txt} -> {a} ({ma}) y Cód. {cod}: se usa el Cód.')
    elif ma != 'ya era ISBN-13':
        avisos.append(f'{tit}: ISBN {isbn_txt} -> {a} ({ma})')
    i = int(i)
    if i not in by.index:
        avisos.append(f'{tit}: {i} NO ESTÁ EN EL SII (va igual, revisar)')
        filas.append((i, int(cant), tit, '¿?', 0, 0)); continue
    filas.append((i, int(cant), tit, str(by.loc[i, 'TITULO']), int(by.loc[i, 'ID ARTICULO']), int(by.loc[i, 'DISPONIBLE'])))
isb = [f[0] for f in filas]; assert len(isb) == len(set(isb)), 'ISBN repetido'
wb = Workbook(); ws = wb.active; ws.title = 'Hoja1'
for r, (i, q, *_ ) in enumerate(filas, 6):
    ws.cell(r, 1, i).number_format = '0'; ws.cell(r, 2, q); ws.cell(r, 3, etq)
ws.column_dimensions['A'].width = 16; ws.column_dimensions['C'].width = 22
wb.save(fout)
print(len(filas), 'líneas', sum(f[1] for f in filas), 'u')
for a in avisos: print('  -', a)
for f in filas: print(f'{f[0]} x{f[1]}  ID {f[4]}  disp {f[5]:>5}  {f[3][:45]}   <- {f[2][:40]}')
