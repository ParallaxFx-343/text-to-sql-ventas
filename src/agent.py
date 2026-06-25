"""Agente text-to-SQL: pregunta en español → SQL → resultado."""

import re
import sqlite3
from pathlib import Path

import anthropic
import pandas as pd
from dotenv import load_dotenv

from .schema import obtener_schema

load_dotenv()

DB_PATH = Path(__file__).resolve().parent.parent / "data.db"
MODELO = "claude-haiku-4-5-20251001"
MAX_REINTENTOS = 1

PALABRAS_PROHIBIDAS = re.compile(
    r"\b(INSERT|UPDATE|DELETE|DROP|ALTER|CREATE|TRUNCATE|REPLACE|MERGE)\b",
    re.IGNORECASE,
)

PROMPT_SQL = """Sos un asistente que genera consultas SQL para SQLite.
Tenés acceso al siguiente esquema de base de datos de una distribuidora editorial:

{schema}

El usuario va a hacer preguntas en español sobre estos datos.
Tu tarea es devolver ÚNICAMENTE la consulta SQL que responde la pregunta.

Reglas:
- Solo SELECT. Nunca INSERT, UPDATE, DELETE, DROP ni ninguna escritura.
- SQL válido para SQLite.
- Si la pregunta es ambigua, hacé tu mejor interpretación y consultá.
- No agregues explicaciones, markdown ni comentarios. Solo la query SQL pura.
"""

PROMPT_EXPLICACION = """Sos un analista de datos que explica resultados de forma clara y concisa.

El usuario preguntó: "{pregunta}"
La consulta SQL ejecutada fue:
{sql}

El resultado fue:
{resultado}

Explicá el resultado en 1-2 oraciones, en español, con tono de negocio.
Mencioná los datos más relevantes. No repitas la pregunta ni el SQL.
"""


class AgenteSQL:
    def __init__(self, db_path: str | Path = DB_PATH):
        self.db_path = Path(db_path)
        self.client = anthropic.Anthropic()
        self.schema = obtener_schema(self.db_path)

    def generar_sql(self, pregunta: str) -> str:
        """Pide al LLM que genere SQL a partir de la pregunta."""
        respuesta = self.client.messages.create(
            model=MODELO,
            max_tokens=1024,
            system=PROMPT_SQL.format(schema=self.schema),
            messages=[{"role": "user", "content": pregunta}],
        )
        sql = respuesta.content[0].text.strip()
        sql = sql.removeprefix("```sql").removeprefix("```").removesuffix("```").strip()
        return sql

    def validar_sql(self, sql: str) -> None:
        """Verifica que la query sea de solo lectura. Lanza ValueError si no."""
        if PALABRAS_PROHIBIDAS.search(sql):
            raise ValueError(
                f"Query rechazada: solo se permiten consultas SELECT.\nSQL recibido: {sql}"
            )

    def ejecutar_sql(self, sql: str) -> pd.DataFrame:
        """Ejecuta la query sobre SQLite y devuelve un DataFrame."""
        conn = sqlite3.connect(self.db_path)
        try:
            df = pd.read_sql_query(sql, conn)
        finally:
            conn.close()
        return df

    def explicar_resultado(self, pregunta: str, sql: str, df: pd.DataFrame) -> str:
        """Pide al LLM una explicación de negocio del resultado."""
        resultado_txt = df.head(20).to_string(index=False)
        respuesta = self.client.messages.create(
            model=MODELO,
            max_tokens=512,
            messages=[
                {
                    "role": "user",
                    "content": PROMPT_EXPLICACION.format(
                        pregunta=pregunta, sql=sql, resultado=resultado_txt
                    ),
                }
            ],
        )
        return respuesta.content[0].text.strip()

    def consultar(self, pregunta: str, explicar: bool = True) -> dict:
        """Flujo completo: pregunta → SQL → ejecución → resultado + explicación."""
        intentos = 0
        ultimo_error = None

        while intentos <= MAX_REINTENTOS:
            if intentos == 0:
                sql = self.generar_sql(pregunta)
            else:
                sql = self.generar_sql(
                    f"{pregunta}\n\n(El intento anterior falló con este error: {ultimo_error}. "
                    f"Corregí la query.)"
                )

            self.validar_sql(sql)

            try:
                df = self.ejecutar_sql(sql)
                resultado = {
                    "pregunta": pregunta,
                    "sql": sql,
                    "resultado": df,
                    "explicacion": None,
                }
                if explicar and not df.empty:
                    resultado["explicacion"] = self.explicar_resultado(pregunta, sql, df)
                return resultado

            except Exception as e:
                ultimo_error = str(e)
                intentos += 1

        raise RuntimeError(
            f"No se pudo ejecutar la consulta después de {MAX_REINTENTOS + 1} intentos. "
            f"Último error: {ultimo_error}"
        )
