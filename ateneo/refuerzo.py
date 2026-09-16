#!/usr/bin/env python3
"""Genera los archivos de una accion comercial de reposicion a partir del SII.

Todas las acciones que pide la editorial son la misma cuenta con parametros
distintos: "completar a N ejemplares en cada local, salvo estos que van a M".
Cubre el ranking de Splendid, los grupos de reposicion, las prioridades, las
promos y los refuerzos por fecha (Dia de la Madre, etc.).

    falta = max(0, objetivo_del_local - stock_actual_del_local)

Entradas
    SII       export de stock por sucursal (.xls / .xlsx)
    regla     JSON con los titulos y los objetivos (ver ejemplo_regla.json)

Salidas (en el directorio que se indique)
    csv/refuerzo_<SUCURSAL>.csv   un archivo por local, para cargar
    PAEC_<nombre>.csv             un solo archivo con el total por titulo
    desglose_x_titulo.txt         para separar en el deposito
    desglose_x_sucursal.txt       para la pantalla de distribucion
    resumen_<nombre>.xlsx         control por titulo

Uso:
    python refuerzo.py --sii SII.xls --regla regla.json --salida ./out
"""
import argparse
import json
import os
import re
import shutil
from pathlib import Path

import pandas as pd

AQUI = Path(__file__).parent


def slug(nombre):
    return re.sub(r'_+', '_', re.sub(r'[^A-Z0-9]', '_', nombre.upper())).strip('_')


def limpiar(texto):
    """Los listados traen espacios duros pegados; rompen las busquedas."""
    return re.sub(r'\s+', ' ', str(texto).replace('\xa0', ' ')).strip()


def columnas_sucursal(df, excluir):
    """Las columnas de sucursal del SII son todas las que siguen a CLASIFIC."""
    cols = list(df.columns)
    if 'CLASIFIC' in cols:
        inicio = cols.index('CLASIFIC') + 1
    else:
        inicio = 21
    todas = cols[inicio:]
    return todas, [c for c in todas if c not in excluir]


def cargar_sii(ruta, excluir):
    df = pd.read_excel(ruta)
    df['ISBN'] = pd.to_numeric(df['ISBN'], errors='coerce')
    todas, activas = columnas_sucursal(df, excluir)
    return df.set_index('ISBN'), todas, activas


def calcular(sii, activas, regla):
    """Devuelve (falta, titulos). falta = {(sucursal, isbn): cantidad}."""
    falta, titulos = {}, []
    for grupo in regla['grupos']:
        base = grupo['completar_a']
        exc = grupo.get('excepciones', {})
        for isbn in grupo['isbns']:
            isbn = int(isbn)
            if isbn not in sii.index:
                raise SystemExit(f'ISBN {isbn} no esta en el SII')
            fila = sii.loc[isbn]
            titulos.append({
                'isbn': isbn,
                'id': int(fila['ID ARTICULO']),
                'titulo': limpiar(fila['TITULO']),
                'grupo': grupo.get('nombre', ''),
                'deposito': int(fila.get('DEPOSITO', 0) or 0),
                'disponible': int(fila.get('DISPONIBLE', 0) or 0),
            })
            for suc in activas:
                objetivo = exc.get(suc, base)
                stock = fila[suc]
                stock = 0 if pd.isna(stock) else int(stock)
                f = max(0, objetivo - stock)
                if f:
                    falta[(suc, isbn)] = f
    return falta, titulos


def redondear(falta, titulos, activas, regla):
    """Redondea el pedido por titulo y manda el excedente a los locales grandes.

    Nunca pide mas de lo que hay disponible en deposito: si el redondeo se
    pasa, el titulo queda en su cantidad exacta.
    """
    paso = regla.get('redondeo')
    destinos = [d for d in regla.get('excedente_a', []) if d in activas]
    for t in titulos:
        exacto = sum(v for (s, i), v in falta.items() if i == t['isbn'])
        t['exacto'] = exacto
        t['pedido'] = exacto
        t['excedente'] = 0
        if not paso or exacto == 0 or not destinos:
            continue
        r = -(-exacto // paso) * paso
        if r > t['disponible']:
            continue
        sobra = r - exacto
        for k in range(sobra):
            d = destinos[k % len(destinos)]
            falta[(d, t['isbn'])] = falta.get((d, t['isbn']), 0) + 1
        t['pedido'] = r
        t['excedente'] = sobra
    return falta, titulos


def escribir(falta, titulos, activas, regla, salida):
    salida = Path(salida)
    salida.mkdir(parents=True, exist_ok=True)
    nombre = regla.get('nombre', 'ACCION')
    prefijo = regla.get('prefijo_archivo', 'refuerzo')
    orden_suc = sorted(activas)
    activos = [t for t in titulos if t['pedido'] > 0]

    # --- un CSV por sucursal ---
    dcsv = salida / 'csv'
    shutil.rmtree(dcsv, ignore_errors=True)
    dcsv.mkdir()
    for suc in orden_suc:
        lineas = [f"{t['isbn']};{falta[(suc, t['isbn'])]}" for t in titulos
                  if (suc, t['isbn']) in falta]
        if lineas:
            (dcsv / f'{prefijo}_{slug(suc)}.csv').write_bytes(
                ('\r\n'.join(lineas) + '\r\n').encode())

    # --- un solo CSV con el total por titulo (el PAEC) ---
    (salida / f'PAEC_{nombre}.csv').write_bytes(
        ('\r\n'.join(f"{t['isbn']};{t['pedido']}" for t in activos) + '\r\n').encode())

    cab = [f'{nombre} - SII procesado con refuerzo.py']
    if regla.get('rotulo'):
        cab.append(f'ROTULAR LA MERCADERIA COMO "{regla["rotulo"]}"')
    for g in regla['grupos']:
        exc = ', '.join(f'{k} {v}' for k, v in g.get('excepciones', {}).items())
        cab.append(f'Grupo {g.get("nombre","")} ({len(g["isbns"])} tit.): '
                   f'completar a {g["completar_a"]}' + (f' / excepciones: {exc}' if exc else ''))
    if regla.get('redondeo'):
        cab.append(f'Pedido redondeado a multiplos de {regla["redondeo"]}; '
                   f'el excedente va a {", ".join(regla.get("excedente_a", []))}.')

    # --- desglose por titulo (como separa el deposito) ---
    L = list(cab) + ['']
    for n, t in enumerate(activos, 1):
        lin = [(s, falta[(s, t['isbn'])]) for s in orden_suc if (s, t['isbn']) in falta]
        L += ['=' * 72,
              f"{n:>2}. {t['titulo'][:52]:<54}{sum(v for _, v in lin):>5} u",
              f"    ISBN {t['isbn']}   ID {t['id']}   grupo {t['grupo']}   en {len(lin)} sucursales",
              '=' * 72]
        L += [f'    {v:>4}  {s}' for s, v in lin] + ['']
    L += ['=' * 72, f'{"TOTAL":<58}{sum(t["pedido"] for t in activos):>5} u  en {len(activos)} titulos']
    (salida / 'desglose_x_titulo.txt').write_text('\n'.join(L) + '\n', encoding='utf-8-sig')

    # --- desglose por sucursal ---
    L = list(cab) + ['']
    for suc in orden_suc:
        lin = [(t, falta[(suc, t['isbn'])]) for t in titulos if (suc, t['isbn']) in falta]
        if not lin:
            continue
        L += ['=' * 78, f'{suc:<58}TOTAL: {sum(v for _, v in lin):>6}', '=' * 78]
        L += [f"  {t['isbn']:<15}{t['id']:<9}{t['grupo']}  {t['titulo'][:42]:<44}{v:>4}" for t, v in lin]
        L += ['']
    (salida / 'desglose_x_sucursal.txt').write_text('\n'.join(L) + '\n', encoding='utf-8-sig')

    # --- resumen de control ---
    resumen = pd.DataFrame([{
        'ISBN': t['isbn'], 'ID': t['id'], 'GRUPO': t['grupo'], 'TITULO': t['titulo'],
        'A PEDIR': t['pedido'], 'NECESIDAD EXACTA': t['exacto'], 'EXCEDENTE': t['excedente'],
        'DEPOSITO': t['deposito'], 'DISPONIBLE': t['disponible'],
        'ESTADO': 'OK' if t['disponible'] >= t['pedido'] else f"FALTAN {t['pedido'] - t['disponible']}",
    } for t in titulos])
    with pd.ExcelWriter(salida / f'resumen_{nombre}.xlsx') as xw:
        resumen.to_excel(xw, sheet_name='Titulos', index=False)
        pd.DataFrame([{'SUCURSAL': s, 'ISBN': t['isbn'], 'TITULO': t['titulo'],
                       'CANTIDAD': falta[(s, t['isbn'])]}
                      for s in orden_suc for t in titulos if (s, t['isbn']) in falta]
                     ).to_excel(xw, sheet_name='Distribucion', index=False)
    return resumen


def controlar(falta, titulos, activas):
    """Cuadre por dos caminos independientes, y cobertura de deposito."""
    por_local = sum(falta.values())
    por_titulo = sum(t['pedido'] for t in titulos)
    print(f'\n--- controles ---')
    print(f'suma por local  : {por_local}')
    print(f'suma por titulo : {por_titulo}')
    print(f'cuadre          : {"OK" if por_local == por_titulo else "*** NO CUADRA ***"}')
    print(f'sucursales con pedido: {len({s for s, _ in falta})} de {len(activas)}')
    cortos = [t for t in titulos if t['pedido'] > t['disponible']]
    if cortos:
        print(f'\ntitulos que el deposito NO cubre ({len(cortos)}):')
        for t in cortos:
            print(f"  {t['isbn']}  pide {t['pedido']:>4}  disponible {t['disponible']:>4}  "
                  f"faltan {t['pedido'] - t['disponible']:>4}  {t['titulo'][:40]}")
    vacios = [t for t in titulos if t['pedido'] == 0]
    if vacios:
        print(f'\ntitulos que no entran (todos los locales ya completos): '
              f'{", ".join(t["titulo"][:30] for t in vacios)}')
    return por_local == por_titulo


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--sii', required=True)
    ap.add_argument('--regla', required=True)
    ap.add_argument('--salida', default='./salida')
    ap.add_argument('--config', default=str(AQUI / 'config.json'))
    args = ap.parse_args()

    cfg = json.load(open(args.config, encoding='utf-8'))
    regla = json.load(open(args.regla, encoding='utf-8'))

    excluir = set(cfg['excluir_siempre']) | set(regla.get('excluir_sucursales', []))
    sii, todas, activas = cargar_sii(args.sii, excluir)
    print(f'SII: {len(sii)} articulos | {len(todas)} columnas de sucursal '
          f'| {len(activas)} activas tras excluir {len(todas) - len(activas)}')
    fuera = [c for c in regla.get('excluir_sucursales', []) if c in todas]
    if fuera:
        print(f'excluidas por la regla: {", ".join(fuera)}')

    falta, titulos = calcular(sii, activas, regla)
    falta, titulos = redondear(falta, titulos, activas, regla)
    controlar(falta, titulos, activas)
    escribir(falta, titulos, activas, regla, args.salida)
    print(f'\nescrito en {args.salida}/')


if __name__ == '__main__':
    main()
