"""Normaliza los exports crudos de ventas para tablero_sorpresa.py.

Los exports del sistema traen un styles.xml que openpyxl no acepta (IndexError en
cellStyleXfs), asi que se leen con calamine (pip install python-calamine).
- Detalle Venta Articulo: deja solo los LIBRO SORPRESA y las columnas
  Sucursal, Fecha, Articulo, Titulo, Cantidad, Bruto, Total, ISBN (lo mismo que
  Franco hacia a mano). Las NC quedan, con cantidad negativa.
- Ventas Proveedor Agrupadas: se reescribe igual, solo para que se pueda abrir.

Uso: python -P normalizar_ventas.py DETALLE_CRUDO.xlsx AGRUPADAS_CRUDO.xlsx SALIDA_DIR
"""
import sys
from pathlib import Path

import pandas as pd

f_det, f_agr, out = sys.argv[1:4]
out = Path(out)
out.mkdir(parents=True, exist_ok=True)
d = pd.read_excel(f_det, engine='calamine', dtype={'ISBN': str})
s = d[d.Titulo.astype(str).str.startswith('LIBRO SORPRESA')]
s = s[['Sucursal', 'Fecha', 'Articulo', 'Titulo', 'Cantidad', 'Bruto', 'Total', 'ISBN']]
assert s.Articulo.nunique() <= 6, sorted(s.Articulo.unique())
s.to_excel(out / 'detalle_sorpresa.xlsx', index=False)
a = pd.read_excel(f_agr, engine='calamine', dtype={'Código / ISBN': str})
a.to_excel(out / 'agrupadas.xlsx', index=False)
print(f'detalle: {len(d)} filas -> {len(s)} de Sorpresa ({int(s.Cantidad.sum())} u, '
      f'{s.Fecha.min():%d/%m} al {s.Fecha.max():%d/%m}) | agrupadas: {len(a)} filas')
