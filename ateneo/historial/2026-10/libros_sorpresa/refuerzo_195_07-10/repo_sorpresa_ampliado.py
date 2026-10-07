"""Refuerzo de Libros Sorpresa ampliado a 195 de cada titulo (Franco, 07/10/2026).

Reglas:
1. AMBA: como la repo anterior. Por titulo en cada local, si el stock le alcanza
   para menos de 3 dias, se repone lo vendido del 01 al 06/10. Grand Splendid,
   que nunca recibio, va con carga inicial.
2. Carga inicial de 8 de cada titulo para los locales que nunca recibieron:
   GRAND SPLENDID (AMBA), NEUQUEN PORTAL PATAGONIA y USHUAIA (interior).
3. Interior, con lo que sobra de cada titulo (195 - AMBA - cargas iniciales):
   a) primero se repone lo vendido a los que tienen menos de 3 dias; si no
      alcanza, se recorta en proporcion a lo que necesita cada uno
   b) lo que sobra se reparte en proporcion a lo vendido de ese titulo en cada
      local del interior
4. Bio: todo a AMBA (salvo las cargas iniciales del interior), nivelando el
   stock: cada unidad va al local de AMBA que tiene menos Bio; si empatan,
   primero el que mas Sorpresa vende.
Redondeo: metodo del mayor resto, asi cada titulo suma exacto.

Uso: python -P repo_sorpresa_ampliado.py DETALLE.xlsx SII.xlsx config.json amba.txt SALIDA_DIR
"""
import json
import sys
from pathlib import Path

import pandas as pd

f_det, f_sii, f_cfg, f_amba, out = sys.argv[1:6]
out = Path(out)
TOPE, CORTE, INICIAL = 195, 3, 8
SIN_RECIBIR = ['ATENEO GRAND SPLENDID', 'NEUQUEN PORTAL PATAGONIA', 'USHUAIA']
AMBA = open(f_amba, encoding='utf-8').read().split('\n')
padron = json.load(open(f_cfg, encoding='utf-8'))['orden_paec'] + ['DEVOTO']
INTERIOR = [l for l in padron if l not in AMBA]

det = pd.read_excel(f_det)
det['local'] = det.Sucursal.str.replace(r'\s*\(\d+\)\s*$', '', regex=True).str.strip()
det['dia'] = det.Fecha.dt.date
dias = sorted(det.dia.unique())
ult = dias[-2]                                   # el ultimo dia del export es parcial
d = det[det.dia <= ult]
primera = det[det.Cantidad > 0].groupby('local').dia.min()
vend = d.groupby(['local', 'Articulo']).Cantidad.sum()
vend_total_local = d.groupby('local').Cantidad.sum()

sii = pd.read_excel(f_sii)
s = sii[sii.TITULO.astype(str).str.startswith('LIBRO SORPRESA')]
assert len(s) == 6
arts = {int(r['ID ARTICULO']): (str(int(r.ISBN)), r.TITULO.replace('LIBRO SORPRESA ', '').strip())
        for _, r in s.iterrows()}
BIO = next(a for a, (_, t) in arts.items() if t == 'BIO')
stock = {(loc, int(r['ID ARTICULO'])): int(0 if pd.isna(r[loc]) else r[loc])
         for _, r in s.iterrows() for loc in padron}


def necesidad(loc, art):
    """Lo vendido si el stock alcanza para menos de 3 dias; si no, 0."""
    if loc not in primera.index:
        return 0
    n = (ult - primera[loc]).days + 1
    v = int(vend.get((loc, art), 0))
    if v <= 0:
        return 0
    return v if stock[(loc, art)] / (v / n) < CORTE else 0


def mayor_resto(total, pesos):
    """Reparte 'total' unidades en proporcion a 'pesos' (dict), en enteros."""
    suma = sum(pesos.values())
    if total <= 0 or suma <= 0:
        return {k: 0 for k in pesos}
    exacto = {k: total * p / suma for k, p in pesos.items()}
    base = {k: int(v) for k, v in exacto.items()}
    resto = total - sum(base.values())
    for k in sorted(pesos, key=lambda k: (-(exacto[k] - base[k]), -pesos[k], k))[:resto]:
        base[k] += 1
    return base


filas = []                                       # (local, art, cant, zona, motivo)
for art in arts:
    asignado = 0
    # cargas iniciales (incluye Bio)
    for loc in SIN_RECIBIR:
        q = max(0, INICIAL - stock[(loc, art)])
        if q and not (art == BIO and loc in AMBA):   # el Bio de AMBA sale de la nivelacion
            filas.append((loc, art, q, 'AMBA' if loc in AMBA else 'INTERIOR', 'carga inicial'))
            asignado += q
    if art == BIO:
        pool = TOPE - asignado
        nivel = {loc: stock[(loc, art)] for loc in AMBA}
        extra = {loc: 0 for loc in AMBA}
        for _ in range(pool):
            loc = min(AMBA, key=lambda l: (nivel[l], -int(vend_total_local.get(l, 0)), l))
            nivel[loc] += 1
            extra[loc] += 1
        for loc, q in extra.items():
            if q:
                filas.append((loc, art, q, 'AMBA', 'nivelacion Bio'))
        continue
    # AMBA: menos de 3 dias, reponer lo vendido
    for loc in AMBA:
        if loc in SIN_RECIBIR:
            continue
        q = necesidad(loc, art)
        if q:
            filas.append((loc, art, q, 'AMBA', 'menos de 3 dias'))
            asignado += q
    # interior
    libre = TOPE - asignado
    assert libre >= 0, (arts[art], libre)
    nec = {loc: necesidad(loc, art) for loc in INTERIOR if loc not in SIN_RECIBIR}
    nec = {k: v for k, v in nec.items() if v}
    if sum(nec.values()) > libre:
        reparto = mayor_resto(libre, nec)        # no alcanza: recorte proporcional
        motivo_extra = None
    else:
        reparto = dict(nec)
        sobra = libre - sum(nec.values())
        pesos = {loc: int(vend.get((loc, art), 0)) for loc in INTERIOR if loc not in SIN_RECIBIR}
        pesos = {k: v for k, v in pesos.items() if v > 0}
        prop = mayor_resto(sobra, pesos)
        motivo_extra = prop
    for loc, q in reparto.items():
        if q:
            filas.append((loc, art, q, 'INTERIOR', 'menos de 3 dias' if q == nec[loc] else 'menos de 3 dias (recortado)'))
    if motivo_extra:
        for loc, q in motivo_extra.items():
            if q:
                filas.append((loc, art, q, 'INTERIOR', 'reparto por venta'))

rep = pd.DataFrame(filas, columns=['local', 'art', 'cant', 'zona', 'motivo'])
por_local = rep.groupby(['local', 'art', 'zona']).cant.sum().reset_index()
tot = rep.groupby('art').cant.sum()
assert (tot == TOPE).all(), tot.to_dict()
orden = ['MISTERIO', 'EROTICA', 'ROMANTICA', 'CONTEMPORANEA', 'HISTORICA', 'BIO']
arts_ord = sorted(arts, key=lambda a: orden.index(arts[a][1]))

out.mkdir(parents=True, exist_ok=True)
(out / 'PAEC_REFUERZO_SORPRESA_AMBA_E_INTERIOR.csv').write_bytes(
    ('\r\n'.join(f'{arts[a][0]};{int(tot[a])}' for a in arts_ord) + '\r\n').encode())
L = ['REFUERZO LIBROS SORPRESA - AMBA E INTERIOR - DESGLOSE POR TITULO',
     f'{TOPE} de cada titulo. Stock: SII 07/10 07:00. Venta: {dias[0]:%d/%m} al {ult:%d/%m}.',
     'AMBA: titulos con menos de 3 dias en el local, se repone lo vendido.',
     'Grand Splendid, Neuquen Portal Patagonia y Ushuaia (nunca recibieron): carga inicial de 8.',
     'Interior: primero los de menos de 3 dias (reponer lo vendido); lo que sobra, por venta.',
     'Bio: todo en AMBA, nivelando el stock entre locales.',
     f'{len(arts)} titulos | {por_local.local.nunique()} locales '
     f'({por_local[por_local.zona == "AMBA"].local.nunique()} AMBA, '
     f'{por_local[por_local.zona == "INTERIOR"].local.nunique()} interior) | {int(tot.sum())} unidades', '']
for k, a in enumerate(arts_ord, 1):
    b = por_local[por_local.art == a].sort_values('local')
    L += ['=' * 66, f'{k}. LIBRO SORPRESA {arts[a][1]:<34}{int(tot[a]):>6} u',
          f'   EAN {arts[a][0]}   ID {a}   en {len(b)} locales', '=' * 66]
    L += [f'   {int(r.cant):>4}  {r.local}' for r in b.itertuples()] + ['']
L += ['=' * 66, f'{"TOTAL":<52}{int(tot.sum()):>6} u']
(out / 'DESGLOSE_REFUERZO_SORPRESA_AMBA_E_INTERIOR.txt').write_text('\n'.join(L) + '\n', encoding='utf-8-sig')
rep.assign(ean=rep.art.map(lambda a: arts[a][0]), titulo=rep.art.map(lambda a: arts[a][1]),
           stock_07_10=[stock[(l, a)] for l, a in zip(rep.local, rep.art)]) \
   .to_csv(out / 'detalle_refuerzo.csv', sep=';', index=False, encoding='utf-8-sig')

print(f'{int(tot.sum())} u | {por_local.local.nunique()} locales')
print(rep.pivot_table(index='motivo', columns=rep.art.map(lambda a: arts[a][1]), values='cant',
                      aggfunc='sum', fill_value=0).to_string())
print(rep.groupby('zona').cant.sum().to_dict())
