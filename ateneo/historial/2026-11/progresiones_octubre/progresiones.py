"""Progresiones de octubre 2026 sobre la planilla de Franco (Progresiones_Cadena_*.xlsx).
Reemplaza los datos de las hojas Inf_Progresiones, 2025 y 2026 con los exports del mes completo y extiende los rangos
de las fórmulas de Regiones (quedaban cortos: '2026'!A2:J75 dejaba afuera filas, Sorpresa!B2:B3040 cortaba el detalle).
Los totales auxiliares al pie de las hojas de datos se rehacen por número de sucursal (antes apuntaban a filas fijas).
La hoja Sorpresa se llena con lo que haya (detalle o, si no hay, totales por local del tablero de Sorpresa).
Uso: python -P progresiones.py PLANTILLA.xlsx INF.xls VTAS_2025.xlsx VTAS_2026.xlsx SORPRESA.(xlsx|csv) SALIDA.xlsx"""
import sys, re, pandas as pd
from openpyxl import load_workbook
fpl, finf, f25, f26, fsor, fout = sys.argv[1:7]
wb = load_workbook(fpl)

def limpiar(ws, desde=1):
    for row in ws.iter_rows(min_row=desde, max_row=ws.max_row):
        for c in row: c.value = None

# --- 2025 / 2026: export tal cual (encabezado en la fila 1) + totales auxiliares por número de sucursal
for nombre, f in (('2025', f25), ('2026', f26)):
    ws = wb[nombre]; d = pd.read_excel(f, header=None, engine='calamine')
    estilo = [ws.cell(2, c)._style for c in range(1, 14)]
    limpiar(ws)
    for i, fila in enumerate(d.values.tolist(), 1):
        for j, v in enumerate(fila, 1):
            c = ws.cell(i, j, None if (isinstance(v, float) and pd.isna(v)) else v)
            if i > 1: c._style = estilo[j - 1]
    n = len(d)
    # pie: total, ferias, Editorial (4) y total de locales = total - ferias - Editorial (igual que antes, sin filas fijas)
    ws.cell(n + 4, 9, 'Total');                   ws.cell(n + 4, 10, f'=SUM(J2:J{n})')
    ws.cell(n + 5, 9, 'Ferias (300-305)');        ws.cell(n + 5, 10, f'=SUMIFS(J2:J{n},A2:A{n},">=300",A2:A{n},"<=305")')
    ws.cell(n + 6, 9, 'Editorial El Ateneo (4)'); ws.cell(n + 6, 10, f'=SUMIF(A2:A{n},4,J2:J{n})')
    ws.cell(n + 7, 9, 'Locales + canales');       ws.cell(n + 7, 10, f'=J{n+4}-J{n+5}-J{n+6}')

# --- Inf_Progresiones: el export no trae el número de sucursal; va en B, como en la planilla
ws = wb['Inf_Progresiones']; d = pd.read_excel(finf, header=None, engine='calamine')
st_txt, st_num = ws.cell(3, 1)._style, [ws.cell(3, c)._style for c in range(2, 33)]
limpiar(ws)
for i, fila in enumerate(d.values.tolist(), 1):
    fila = [None if (isinstance(v, float) and pd.isna(v)) else v for v in fila]
    if i <= 2:
        ws.cell(i, 1, fila[0])
        for j, v in enumerate(fila[1:], 3): ws.cell(i, j, v)
        continue
    m = re.search(r'\((\d+)\)\s*$', str(fila[0]))
    ws.cell(i, 1, fila[0])._style = st_txt
    ws.cell(i, 2, int(m.group(1)) if m else None)._style = st_num[0]
    for j, v in enumerate(fila[1:], 3): ws.cell(i, j, v)._style = st_num[min(j - 2, len(st_num) - 1)]
n = len(d)
for col in 'DE':   # pie auxiliar como el de la planilla (total, Editorial 4, ferias), por número de sucursal
    ws[f'{col}{n+4}'] = f'=SUM({col}3:{col}{n})'
    ws[f'{col}{n+5}'] = f'=SUMIF($B$3:$B${n},4,{col}3:{col}{n})'
    ws[f'{col}{n+6}'] = f'={col}{n+4}-{col}{n+5}'
    ws[f'{col}{n+7}'] = f'=SUMIFS({col}3:{col}{n},$B$3:$B${n},">=300",$B$3:$B${n},"<=305")'
    ws[f'{col}{n+8}'] = f'={col}{n+6}-{col}{n+7}'

# --- Sorpresa: detalle (hoja Sorpresa con B = nro y M = cantidad) o totales por local
ws = wb['Sorpresa']; limpiar(ws, 2)
if fsor.endswith('.csv'):
    s = pd.read_csv(fsor, sep=';')
    for i, r in enumerate(s.itertuples(), 2):
        ws.cell(i, 1, f'{r.local} ({r.nro})'); ws.cell(i, 2, int(r.nro)); ws.cell(i, 12, r.titulo); ws.cell(i, 13, int(r.cant))
else:
    s = pd.read_excel(fsor, header=None, engine='calamine').values.tolist()
    for i, fila in enumerate(s[1:], 2):
        for j, v in enumerate(fila, 1):
            ws.cell(i, j, None if (isinstance(v, float) and pd.isna(v)) else v)
        m = re.search(r'\((\d+)\)', str(fila[0]))
        ws.cell(i, 2, int(m.group(1)) if m else None)

# --- Regiones: rangos completos
ws = wb['Regiones']
for row in ws.iter_rows():
    for c in row:
        if isinstance(c.value, str) and c.value.startswith('='):
            v = c.value
            v = v.replace("'2025'!$A$2:$J$68", "'2025'!$A$2:$J$500").replace("'2026'!$A$2:$J$75", "'2026'!$A$2:$J$500")
            v = v.replace('Inf_Progresiones!$B$3:$E$68', 'Inf_Progresiones!$B$3:$E$500').replace('Inf_Progresiones!$B$3:$F$68', 'Inf_Progresiones!$B$3:$F$500')
            v = v.replace('Sorpresa!$B$2:$B$3040', 'Sorpresa!$B$2:$B$20000').replace('Sorpresa!$M$2:$M$3040', 'Sorpresa!$M$2:$M$20000')
            v = re.sub(r"'(2025|2026)'!\$A\$2:\$A\$6\d", r"'\1'!$A$2:$A$500", v)
            v = re.sub(r"'(2025|2026)'!\$J\$2:\$J\$6\d", r"'\1'!$J$2:$J$500", v)
            c.value = v
# sorpresa de los locales comparables: sin las bocas nuevas (antes restaba la de toda la región y un 2129 fijo)
ws['K80'] = '=+(D80-(K18-K17))/C80-1'
ws['K81'] = '=+(D81-(K34-K33))/C81-1'
ws['K82'] = '=+(D82-(K53-K52-K47))/C82-1'
ws['K83'] = '=+(D83-(K69-K67))/C83-1'
ws['J75'] = '=+D75-(K71+K18-K17+K34-K33+K53-K52-K47+K69-K67)'
wb.save(fout)
print('ok', fout)
