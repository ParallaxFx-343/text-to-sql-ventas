"""Carga los CSVs del dataset a una base SQLite normalizada."""

import sqlite3
import pandas as pd
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent.parent / "data.db"
DATA_DIR = Path(__file__).resolve().parent.parent / "data"

COLUMNAS_FIJAS_STOCK = [
    "ID ARTICULO", "ISBN", "IDEDITOR", "DESCEDITOR", "TITULO",
    "CONSIGNADO", "DEPOSITO", "FALLADOS", "DISPONIBLE", "CADENA",
    "TOTAL", "COSTO", "PVP", "SECCION", "GRUPO", "FAMILIA",
    "SUBFAMILIA", "PESO",
]


def cargar_ventas(conn: sqlite3.Connection) -> int:
    df = pd.read_csv(DATA_DIR / "ventas.csv")
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    df.to_sql("ventas", conn, if_exists="replace", index=False)
    return len(df)


def cargar_stock_editorial(conn: sqlite3.Connection) -> int:
    df = pd.read_csv(DATA_DIR / "stock_editorial.csv")
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    df.to_sql("stock_editorial", conn, if_exists="replace", index=False)
    return len(df)


def cargar_stock_sucursales(conn: sqlite3.Connection) -> int:
    """Normaliza la tabla wide (una columna por sucursal) a formato long."""
    df = pd.read_csv(DATA_DIR / "stock_sucursales.csv")

    cols_fijas = [c for c in COLUMNAS_FIJAS_STOCK if c in df.columns]
    cols_sucursales = [c for c in df.columns if c not in cols_fijas]

    df_long = df.melt(
        id_vars=cols_fijas,
        value_vars=cols_sucursales,
        var_name="SUCURSAL",
        value_name="STOCK_UNIDADES",
    )
    df_long = df_long[df_long["STOCK_UNIDADES"] > 0]
    df_long.columns = [c.strip().lower().replace(" ", "_") for c in df_long.columns]
    df_long.to_sql("stock_sucursales", conn, if_exists="replace", index=False)
    return len(df_long)


def main():
    if DB_PATH.exists():
        DB_PATH.unlink()

    conn = sqlite3.connect(DB_PATH)
    try:
        n_ventas = cargar_ventas(conn)
        n_stock_ed = cargar_stock_editorial(conn)
        n_stock_suc = cargar_stock_sucursales(conn)
        conn.commit()

        print(f"Base creada en {DB_PATH}")
        print(f"  ventas:           {n_ventas:,} filas")
        print(f"  stock_editorial:  {n_stock_ed:,} filas")
        print(f"  stock_sucursales: {n_stock_suc:,} filas (normalizada de wide a long)")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
