# Herramientas operativas — El Ateneo

Scripts para las tareas que se repiten todos los meses. Las reglas de negocio
que hay detrás están en el `CLAUDE.md` de la raíz del repo.

Requisitos: `pandas`, `openpyxl`, `xlrd` (para los `.xls` del SII).

```
pip install pandas openpyxl xlrd
```

---

## `config.json` — padrón de sucursales

El dato más caro de perder. Contiene:

- `excluir_siempre` — las 14 columnas del SII que no son locales de venta
  (depósitos, ferias itinerantes, canales alternativos, pruebas)
- `orden_paec` — el orden en que van los locales en los TXT de PAEC
- `grupos` — los 4 grupos de reposición de 15 locales cada uno
- `notas_sucursales` — altas recientes y casos especiales

Al 16/09/2026 el SII trae **76 columnas de sucursal**; menos las 14 exclusiones
quedan **62 activas**. De esas, 60 están repartidas en los 4 grupos;
ATENEO GRAND SPLENDID tiene circuito propio y DEVOTO todavía no integra ninguno.

---

## `refuerzo.py` — acciones de reposición

Todas las acciones que pide la editorial son la misma cuenta:

```
falta = max(0, objetivo_del_local − stock_actual_del_local)
```

Lo que cambia son los objetivos. El script toma el SII y un JSON con la regla,
y escribe todo lo que hace falta para cargar y para que el depósito prepare.

```bash
python ateneo/refuerzo.py --sii SII.xls --regla mi_regla.json --salida ./out
```

Salidas en `./out`:

| Archivo | Para qué |
|---|---|
| `csv/refuerzo_<SUCURSAL>.csv` | cargar local por local |
| `PAEC_<nombre>.csv` | un solo PAEC, total por título |
| `desglose_x_titulo.txt` | para separar en el depósito |
| `desglose_x_sucursal.txt` | para la pantalla de distribución |
| `resumen_<nombre>.xlsx` | control por título y distribución completa |

### La regla

`ejemplo_regla.json` es el Día de la Madre 2026, que sirve de molde:

```json
{
  "nombre": "DIA_DE_LA_MADRE_2026",
  "rotulo": "DIA DE LA MADRE",
  "excluir_sucursales": ["DEVOTO"],
  "redondeo": 10,
  "excedente_a": ["ATENEO GRAND SPLENDID", "ATENEO JURAMENTO", "ATENEO FLORIDA 340"],
  "grupos": [
    {
      "nombre": "A",
      "completar_a": 10,
      "excepciones": {"ATENEO GRAND SPLENDID": 30, "EZEIZA": 3, "TRELEW": 3},
      "isbns": [9789500217446, 9789500211598]
    }
  ]
}
```

- **`grupos`** — cada uno con su `completar_a` y sus `excepciones` por local.
  Una acción simple lleva un solo grupo.
- **`excluir_sucursales` dentro de un grupo** (opcional) — deja un local afuera
  de ese grupo nada más. Sirve cuando la misma acción trata distinto a un local
  según el título: en la acción de agenda + autores de octubre 2026, Splendid
  quedó fuera de la agenda pero entró a 30 en los títulos de autor.
- **`redondeo`** (opcional) — redondea el pedido de cada título hacia arriba
  para que el depósito reciba números enteros. **Nunca pide más de lo que hay
  disponible**: si el redondeo se pasa del stock, el título queda en su
  cantidad exacta.
- **`excedente_a`** (opcional) — los locales grandes que absorben la diferencia
  del redondeo, repartida en round-robin.
- **`excluir_sucursales`** — se suma a las exclusiones fijas de `config.json`.

### Controles

Antes de escribir nada imprime:

- el cuadre por **dos caminos independientes** (suma por local vs. suma por
  título) — si no coinciden, hay un error de cálculo
- los títulos que el **depósito no llega a cubrir**, con cuánto falta
- los títulos que **no entran** porque todos los locales ya están completos

### Verificado contra

El Día de la Madre 2026 (39 títulos, 61 sucursales). Reproduce byte por byte
el `PAEC_DIA_DE_LA_MADRE.csv` que se entregó, y las 1.423 líneas del desglose
por título son idénticas.

---

## `pedido_cliente.py` — unir los pedidos de un cliente

Junta el `PEDIDO_*` y el `REPO_*` de **cada sucursal** del cliente (no mezcla
sucursales) en un archivo `{VENDEDOR}_-_{CLIENTE}_SUC._{SUCURSAL}_-_{CUENTA}.xlsx`
con Hoja1 (ISBN + cantidad desde la fila 7) y Sheet1 (ISBN + descripción +
cantidad).

```bash
python ateneo/pedido_cliente.py PEDIDO_*.xlsx REPO_*.xlsx --vendedor X --sii SII.xls --salida ./out
```

Corrige el dígito verificador, avisa si un título viene con dos ISBN, frena si
se mezclan cuentas, cuadra entrada contra salida y, con `--sii`, muestra ID y
disponible. `--reemplazar VIEJO=NUEVO` aplica una decisión de edición.

---

## `isbn.py` — validar y corregir ISBN

Los listados llegan con los códigos en tres formas que no se distinguen a
simple vista porque todas empiezan con `950`:

1. **ISBN-13 sin el `978`** → alcanza con anteponerlo
2. **ISBN-10 viejo** → hay que descartar su dígito verificador y recalcularlo
3. **ISBN-13 con un dígito mal** → el verificador no cierra

Prefijar `978` a ciegas produce códigos que no existen en el sistema.

```bash
python ateneo/isbn.py codigos.txt --sii SII.xls --lp lista_precios.xlsx
python ateneo/isbn.py codigos.txt --csv salida.csv
cat codigos.txt | python ateneo/isbn.py -
```

Marca con `>>` los que necesitaron corrección y, si se le pasa el SII o la
lista de precios, confirma que el ISBN resuelto exista y trae el título.

Casos reales que detectó: `9500253399` → `9789500253390` (Diccionario
Español-Inglés), `9500286769` → `9789500286763` (El Príncipe), y dos ISBN-13
con el verificador en 3 en vez de 0.
