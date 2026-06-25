"""Interfaz de línea de comandos interactiva para el agente text-to-SQL."""

import sys
import io

from tabulate import tabulate

from .agent import AgenteSQL

if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8", errors="replace")


def main():
    print("=" * 60)
    print("  Agente Text-to-SQL — Datos de Distribución Editorial")
    print("  Escribí tu pregunta en español. 'salir' para terminar.")
    print("=" * 60)
    print()

    agente = AgenteSQL()

    while True:
        try:
            pregunta = input("> Tu pregunta: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nHasta luego!")
            break

        if not pregunta:
            continue
        if pregunta.lower() in ("salir", "exit", "quit"):
            print("Hasta luego!")
            break

        try:
            resultado = agente.consultar(pregunta)

            print(f"\nSQL generado:\n{resultado['sql']}\n")

            df = resultado["resultado"]
            if df.empty:
                print("(Sin resultados)\n")
            else:
                print(tabulate(df.head(20), headers="keys", tablefmt="rounded_grid", showindex=False))
                if len(df) > 20:
                    print(f"  ... ({len(df)} filas en total, mostrando las primeras 20)")
                print()

            if resultado["explicacion"]:
                print(f"Interpretacion: {resultado['explicacion']}\n")

        except ValueError as e:
            print(f"\n[!] {e}\n")
        except Exception as e:
            print(f"\n[ERROR] {e}\n")

        print("-" * 60)


if __name__ == "__main__":
    main()
