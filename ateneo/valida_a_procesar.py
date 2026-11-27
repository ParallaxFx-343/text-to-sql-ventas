"""Revisa un A PROCESAR de Servicio de Novedades (cualquier vendedor): la hoja de carga (la primera; IdCliente en A1, IDs desde la I)
contra la grilla (la hoja con 'NOVEDADES' en la fila 1 e ISBN en la fila 3, clientes desde la fila 6), los IDs por ISBN (LP Argentina /
SII), celdas vacías o con texto, lista de precios, DirEnt contra el maestro (ENTR activa) e importados (0 en los 4 o en ninguno).
Uso: python -P valida_a_procesar.py ARCHIVO.xlsx LP_ARGENTINA.xlsx SII.xls MAESTRO.xls LISTA [IMPORTADOS]"""
import sys, unicodedata, openpyxl, pandas as pd
from openpyxl.utils import get_column_letter as L
fa, flp, fsii, fm, lista = sys.argv[1:6]; lista = int(lista)
IMP = [int(x) for x in sys.argv[6].split(',')] if len(sys.argv) > 6 else []
norm = lambda s: unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().strip().upper()
wb = openpyxl.load_workbook(fa); p = wb.worksheets[0]
g = next(ws for ws in wb.worksheets[1:] if norm(ws['J1'].value) == 'NOVEDADES')
print(f'hoja de carga: {p.title} (activa: {wb.active.title}) | grilla: {g.title}')
a = pd.read_excel(flp, header=2); a = a[a['I.D.'].notna()]; ID = dict(zip(a['I.S.B.N.'].astype('int64'), a['I.D.'].astype(int)))
s = pd.read_excel(fsii); s = s[s['ID ARTICULO'].notna()]
for k, v in zip(pd.to_numeric(s.ISBN, errors='coerce'), s['ID ARTICULO']):
    if pd.notna(k): ID.setdefault(int(k), int(v))
prob, aviso = [], []
gh = {norm(g.cell(5, c).value): c for c in range(1, 10) if g.cell(5, c).value}
ph = [norm(p.cell(1, c).value) for c in range(1, 9)]
cols = [c for c in range(10, g.max_column + 1) if g.cell(3, c).value]
for c in cols:
    i = ID.get(int(g.cell(3, c).value))
    if i != g.cell(5, c).value: prob.append(f'grilla {L(c)}5: ID {g.cell(5, c).value} pero el ISBN {g.cell(3, c).value} es {i}')
pids = [p.cell(1, c).value for c in range(9, p.max_column + 1)]
gids = [ID.get(int(g.cell(3, c).value)) for c in cols]
if pids != gids: prob.append(f'IDs de la carga {pids} != grilla {gids}')
if len(set(pids)) != len(pids): prob.append('ID repetido en la carga')
for c in range(10, g.max_column + 1):
    if c not in cols and any(isinstance(g.cell(r, c).value, (int, float)) and g.cell(r, c).value for r in range(6, g.max_row + 1)):
        aviso.append(f'grilla columna {L(c)} sin ISBN con cantidades')
fg = []
for r in range(6, g.max_row + 1):
    if isinstance(g.cell(r, 1).value, (int, float)): fg.append(r)
    elif fg: break
fp = [r for r in range(2, p.max_row + 1) if p.cell(r, 1).value is not None]
if len(fg) != len(fp): prob.append(f'filas: grilla {len(fg)} vs carga {len(fp)}')
campos = [k for k in ph if k and k not in ('IDCLIENTE',)]
for rg, rp in zip(fg, fp):
    for k in ph:
        if not k or k not in gh: continue
        vg, vp = g.cell(rg, gh[k]).value, p.cell(rp, ph.index(k) + 1).value
        if k.startswith('LISTA'):
            if vp != lista: prob.append(f'carga fila {rp}: lista {vp}')
            continue
        if vg != vp: prob.append(f'carga fila {rp} ({p.cell(rp,1).value}) {k}: grilla {vg!r} / carga {vp!r}')
    for k, c in enumerate(cols):
        vg, vp = g.cell(rg, c).value, p.cell(rp, 9 + k).value
        if vp is None or isinstance(vp, str): prob.append(f'carga {L(9+k)}{rp} ({p.cell(rp,1).value}) vacía/texto {vp!r}')
        elif isinstance(vg, str) and vg.strip(): prob.append(f'grilla {L(c)}{rg} texto {vg!r}')
        elif (vg if isinstance(vg, (int, float)) else 0) != vp: prob.append(f'carga {L(9+k)}{rp}={vp} vs grilla {L(c)}{rg}={vg!r} ({p.cell(rp,1).value})')
    for c in range(1, 8):
        if c != ph.index(next((k for k in ph if k.startswith('OBSERV') or k == 'NRO PEDIDO'), 'X')) + 1 if any(k.startswith('OBSERV') or k == 'NRO PEDIDO' for k in ph) else True:
            v = p.cell(rp, c).value
            if c in (1, 2, 4, 5, 7) and not isinstance(v, int): prob.append(f'carga {L(c)}{rp}: {v!r}')
m = pd.read_excel(fm)
for rp in fp:
    idc, de = p.cell(rp, 1).value, p.cell(rp, 2).value
    mm = m[m.IDCLIENTE == idc]; d = mm[mm.NROSECUENCIA == de]
    if mm.empty: prob.append(f'{idc}: no está en el maestro'); continue
    t = d['TIPO DIRECCION'].tolist(); dire = d.DIRECCION.tolist()
    if t != ['ENTR']: prob.append(f'{idc} {mm.iloc[0].NOMBRE}: DirEnt {de} es {t or "inexistente"} {dire}')
    elif 'NO USAR' in str(dire[0]).upper(): prob.append(f'{idc} {mm.iloc[0].NOMBRE}: DirEnt {de} dice {dire[0]!r}')
    if mm.FECHABAJA.notna().all(): prob.append(f'{idc}: dado de baja')
if IMP:
    ci = [9 + pids.index(i) for i in IMP if i in pids]
    for rp in fp:
        q = [p.cell(rp, c).value or 0 for c in ci]
        if 0 in q and any(q): aviso.append(f'{p.cell(rp,1).value} DirEnt {p.cell(rp,2).value}: importados a medias {q}')
    sin = [f'{p.cell(rp,1).value}/{p.cell(rp,2).value}' for rp in fp if not any(p.cell(rp, c).value for c in ci) and any(p.cell(rp, c).value for c in range(9, 13))]
    aviso.append(f'sin importados (con novedades): {len(sin)} {sin}')
tot = [sum(p.cell(r, 9 + k).value or 0 for r in fp) for k in range(len(cols))]
cuota_row = next((r for r in range(fg[-1] + 1, g.max_row + 1) if any(isinstance(g.cell(r, c).value, (int, float)) for c in cols)), None)
print(f'{len(fp)} filas | {len(pids)} IDs | {sum(tot)} u')
for k, c in enumerate(cols):
    cu = g.cell(cuota_row, c).value if cuota_row else None
    print(f'   {pids[k]} {str(g.cell(2, c).value)[:30]:<30} {tot[k]:>4}' + (f'  cuota {cu}' if cu not in (None, '') else ''))
print(f'\n{len(prob)} PROBLEMAS'); print('\n'.join('  *** ' + x for x in prob))
print(f'\n{len(aviso)} avisos'); print('\n'.join('  - ' + x for x in aviso))
