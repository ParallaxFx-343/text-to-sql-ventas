#!/usr/bin/env python3
"""Validacion y correccion de ISBN de El Ateneo.

Los listados que llegan de los sistemas traen los codigos en tres formas
distintas, y no se distinguen a simple vista porque todas empiezan con 950:

  1. ISBN-13 sin el prefijo 978  -> alcanza con anteponer "978"
  2. ISBN-10 viejo               -> hay que descartar el digito verificador
                                    y recalcularlo sobre el ISBN-13
  3. ISBN-13 con un digito mal   -> el verificador no cierra

Prefijar 978 a ciegas produce codigos que no existen. Este script decide
cual es cada caso por el digito verificador y, si se le pasa un SII o una
lista de precios, confirma contra esas fuentes.

Uso:
    python isbn.py codigos.txt
    python isbn.py codigos.txt --sii SII.xls --lp lista.xlsx
    cat codigos.txt | python isbn.py -
"""
import argparse
import re
import sys


def dv13(doce):
    """Digito verificador de un ISBN-13 a partir de sus primeros 12 digitos."""
    s = sum(int(c) * (1 if i % 2 == 0 else 3) for i, c in enumerate(doce))
    return (10 - s % 10) % 10


def valido13(codigo):
    return len(codigo) == 13 and codigo[12].isdigit() and dv13(codigo[:12]) == int(codigo[12])


def valido10(codigo):
    if len(codigo) != 10:
        return False
    total = 0
    for i, c in enumerate(codigo):
        if c in 'Xx':
            v = 10
        elif c.isdigit():
            v = int(c)
        else:
            return False
        total += v * (10 - i)
    return total % 11 == 0


def resolver(codigo):
    """Devuelve (isbn13, metodo). isbn13 es None si no se pudo resolver."""
    c = re.sub(r'[^0-9Xx]', '', str(codigo))

    if len(c) == 13:
        if valido13(c):
            return c, 'ya era ISBN-13'
        corr = c[:12] + str(dv13(c[:12]))
        return corr, 'ISBN-13 con dv corregido'

    if len(c) == 10:
        # El caso mas comun: es el ISBN-13 al que le cortaron el 978.
        cand = '978' + c
        if valido13(cand):
            return cand, '978 + codigo'
        # Si no, es un ISBN-10 de verdad: se descarta su dv y se recalcula.
        if valido10(c):
            base = '978' + c[:9]
            return base + str(dv13(base)), 'ISBN-10 convertido'
        # Ni una cosa ni la otra: se asume 13 truncado y se recalcula igual,
        # pero se marca para revision manual.
        base = '978' + c[:9]
        return base + str(dv13(base)), 'DUDOSO - revisar'

    return None, f'longitud {len(c)} inesperada'


def cargar_fuente(ruta, col_isbn=None, col_titulo=None, col_id=None):
    """Devuelve {isbn: (titulo, id)} desde un SII o una lista de precios."""
    import pandas as pd
    df = pd.read_excel(ruta)

    def buscar(candidatos, defecto=None):
        for c in df.columns:
            n = str(c).upper()
            if any(k in n for k in candidatos):
                return c
        return defecto

    ci = col_isbn or buscar(['ISBN', 'CODIGO'])
    ct = col_titulo or buscar(['TITULO', 'TÍTULO', 'DESCRIP'])
    cd = col_id or buscar(['ID ARTICULO', 'NRO. ART', 'ID ART'])
    if ci is None:
        raise SystemExit(f'{ruta}: no encuentro la columna de ISBN')

    df[ci] = pd.to_numeric(df[ci], errors='coerce')
    out = {}
    for _, r in df.dropna(subset=[ci]).iterrows():
        out[int(r[ci])] = (r[ct] if ct else None, r[cd] if cd else None)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('archivo', help='archivo con un codigo por linea, o "-" para stdin')
    ap.add_argument('--sii', help='SII para confirmar que el ISBN existe')
    ap.add_argument('--lp', help='lista de precios para confirmar')
    ap.add_argument('--csv', help='escribir el resultado a un CSV')
    args = ap.parse_args()

    texto = sys.stdin.read() if args.archivo == '-' else open(args.archivo, encoding='utf-8-sig').read()
    codigos = [l.strip() for l in texto.splitlines() if l.strip() and any(ch.isdigit() for ch in l)]

    sii = cargar_fuente(args.sii) if args.sii else {}
    lp = cargar_fuente(args.lp) if args.lp else {}

    filas = []
    print(f"{'CODIGO':<17}{'ISBN':<15}{'SII':<5}{'LP':<5}{'METODO':<26}TITULO")
    for c in codigos:
        isbn, metodo = resolver(c)
        if isbn is None:
            print(f'{c:<15}{"-":<15}{"-":<5}{"-":<5}{metodo}')
            filas.append((c, None, metodo, None, None))
            continue
        n = int(isbn)
        en_sii, en_lp = n in sii, n in lp
        titulo = (sii.get(n) or lp.get(n) or (None, None))[0]
        ida = (sii.get(n) or lp.get(n) or (None, None))[1]
        marca = '   ' if metodo in ('978 + codigo', 'ya era ISBN-13') else '>> '
        print(f'{marca}{c:<14}{isbn:<15}'
              f'{("SI" if en_sii else "--") if sii else "":<5}'
              f'{("SI" if en_lp else "--") if lp else "":<5}'
              f'{metodo:<26}{titulo if titulo else ""}')
        filas.append((c, isbn, metodo, ida, titulo))

    raros = [f for f in filas if f[1] is None or 'corregido' in f[2] or 'convertido' in f[2] or 'DUDOSO' in f[2]]
    print(f'\n{len(filas)} codigos | {len(filas) - len(raros)} salen con 978 | {len(raros)} necesitan correccion')
    for f in raros:
        print(f'  {f[0]} -> {f[1]}  ({f[2]})')
    if (sii or lp):
        sin = [f for f in filas if f[1] and int(f[1]) not in sii and int(f[1]) not in lp]
        print(f'sin match en las fuentes: {[f[0] for f in sin] if sin else "ninguno"}')

    if args.csv:
        import csv as _csv
        with open(args.csv, 'w', newline='', encoding='utf-8-sig') as fh:
            w = _csv.writer(fh, delimiter=';')
            w.writerow(['CODIGO', 'ISBN', 'METODO', 'ID', 'TITULO'])
            w.writerows(filas)
        print(f'escrito {args.csv}')


if __name__ == '__main__':
    main()
