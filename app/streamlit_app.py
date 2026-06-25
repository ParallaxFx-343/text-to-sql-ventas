"""Demo web del agente text-to-SQL con Streamlit."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st
from src.agent import AgenteSQL

st.set_page_config(page_title="Text-to-SQL · Distribución Editorial", page_icon="📊")
st.title("📊 Agente Text-to-SQL")
st.caption("Consultá datos de ventas y stock en lenguaje natural, sin escribir SQL.")


@st.cache_resource
def cargar_agente():
    return AgenteSQL()


agente = cargar_agente()

pregunta = st.text_input(
    "Tu pregunta",
    placeholder="Ej: ¿Cuáles fueron las 5 sucursales con más ventas?",
)

if pregunta:
    with st.spinner("Generando consulta..."):
        try:
            resultado = agente.consultar(pregunta)

            st.subheader("🔍 SQL generado")
            st.code(resultado["sql"], language="sql")

            st.subheader("📋 Resultado")
            df = resultado["resultado"]
            if df.empty:
                st.info("Sin resultados para esta consulta.")
            else:
                st.dataframe(df.head(50), use_container_width=True)

            if resultado["explicacion"]:
                st.subheader("💡 Interpretación")
                st.write(resultado["explicacion"])

        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Error: {e}")
