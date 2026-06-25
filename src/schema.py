"""Extrae y formatea el schema de la base SQLite para pasarle al LLM."""

import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent.parent / "data.db"

DESCRIPCIONES = {
    "ventas": (
        "Registro de ventas por sucursal y período. Cada fila es una venta "
        "de un título en una sucursal en un mes determinado. "
        "La columna 'categoria' tiene formato 'LIBROS > Subcategoría' "
        "(ej: 'LIBROS > Juvenil', 'LIBROS > No Ficción'). "
        "Para filtrar por categoría usar LIKE '%Juvenil%' o similar."
    ),
    "stock_editorial": (
        "Stock consolidado por artículo desde la perspectiva editorial. "
        "Incluye stock en depósito, cadena, terceros, valuación y datos del producto."
    ),
    "stock_sucursales": (
        "Stock por artículo y sucursal (formato normalizado). "
        "Cada fila indica cuántas unidades de un artículo tiene una sucursal. "
        "Solo incluye combinaciones con stock > 0."
    ),
}


def obtener_schema(db_path: str | Path = DB_PATH) -> str:
    """Devuelve el schema formateado como texto para incluir en el prompt."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tablas = [row[0] for row in cursor.fetchall()]

    partes = []
    for tabla in tablas:
        desc = DESCRIPCIONES.get(tabla, "")
        cursor.execute(f"PRAGMA table_info('{tabla}')")
        columnas = cursor.fetchall()

        lineas_col = []
        for col in columnas:
            nombre = col[1]
            tipo = col[2] or "TEXT"
            lineas_col.append(f"    {nombre} ({tipo})")

        cursor.execute(f"SELECT COUNT(*) FROM {tabla}")
        n_filas = cursor.fetchone()[0]

        bloque = f"Tabla: {tabla}\n"
        if desc:
            bloque += f"Descripción: {desc}\n"
        bloque += f"Filas: {n_filas:,}\n"
        bloque += "Columnas:\n" + "\n".join(lineas_col)
        partes.append(bloque)

    conn.close()
    return "\n\n".join(partes)


if __name__ == "__main__":
    print(obtener_schema())
