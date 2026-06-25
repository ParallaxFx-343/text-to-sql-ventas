# 📊 Agente Text-to-SQL — Distribución Editorial

**Consultá datos de ventas en lenguaje natural, sin escribir SQL.**

Un agente que traduce preguntas en español a consultas SQL, las ejecuta sobre un dataset real de ventas y stock de una distribuidora editorial, y devuelve el resultado con una explicación de negocio.

![Demo del agente](assets/demo.gif)

---

## El problema de negocio

Un gerente comercial de una distribuidora necesita respuestas sobre sus datos de ventas y stock, pero no escribe SQL. Hoy depende de un analista para cada consulta, lo que genera cuellos de botella y demoras.

Este agente permite hacer preguntas como *"¿Cuáles fueron las 5 sucursales con más ventas?"* y obtener la respuesta al instante, con la query SQL visible para validación.

---

## Cómo funciona

```
Pregunta en español
       │
       ▼
┌─────────────────┐
│  LLM (Claude)   │◄── Schema de las tablas
│  genera SQL     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Validación     │  ← Solo SELECT permitido
│  de seguridad   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Ejecución      │  ← SQLite
│  de la query    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Resultado      │  ← Tabla + explicación
│  + explicación  │    de negocio
└─────────────────┘
```

1. El usuario hace una pregunta en español.
2. El LLM recibe la pregunta junto con el schema completo de las tablas y genera una query SQL.
3. Se valida que la query sea de solo lectura (SELECT).
4. Se ejecuta sobre la base SQLite.
5. Se muestra el resultado como tabla y, opcionalmente, el LLM lo resume en una frase de negocio.

---

## Decisiones técnicas

### SQLite como base de datos
Elegí SQLite por portabilidad: cualquiera puede clonar el repo y correrlo sin levantar un servidor. La base se genera con un solo comando. Para producción, la arquitectura es la misma con PostgreSQL (queda documentado como alternativa).

### Schema como contexto del LLM
El LLM recibe el schema completo (tablas, columnas, tipos, descripciones y convenciones de los datos) en el system prompt. Esto le da contexto suficiente para generar SQL válido sin necesidad de RAG ni embeddings.

### Validación de solo lectura
Antes de ejecutar cualquier query, se verifica con regex que no contenga palabras clave de escritura (INSERT, UPDATE, DELETE, DROP, etc.). Es una capa de seguridad deliberada: este agente solo lee datos, nunca los modifica.

### Manejo de errores y reintentos
Si la query generada falla al ejecutarse, el agente le pasa el error al LLM y le pide que corrija. Máximo 1 reintento — suficiente para corregir errores menores de sintaxis sin entrar en loops.

### SQL visible al usuario
El SQL generado se muestra siempre. No se oculta. Un analista quiere ver la query para validarla, y un reclutador quiere ver que sabés SQL. La transparencia es una decisión de diseño, no un descuido.

---

## Ejemplos reales

### ¿Cuáles fueron las 5 sucursales con más ventas?

```sql
SELECT sucursal, SUM(cantidad) as total_ventas
FROM ventas
GROUP BY sucursal
ORDER BY total_ventas DESC
LIMIT 5;
```

| sucursal             | total_ventas |
|----------------------|--------------|
| Sucursal AMBA 24     | 17.756       |
| Sucursal Feria 01    | 7.579        |
| Sucursal AMBA 01     | 2.664        |
| Sucursal Sur 02      | 750          |
| Sucursal Sur 06      | 600          |

> Sucursal AMBA 24 lidera con 17.756 unidades, más del doble que la segunda (Feria 01 con 7.579).

### ¿Cuánto representa la categoría juvenil sobre el total?

```sql
SELECT SUM(CASE WHEN categoria LIKE '%Juvenil%' THEN cantidad ELSE 0 END)
       * 100.0 / SUM(cantidad) as porcentaje_juvenil
FROM ventas;
```

| porcentaje_juvenil |
|--------------------|
| 53.05%             |

> La categoría Juvenil representa más de la mitad de las unidades vendidas (53%), posicionándose como el segmento más relevante del portafolio.

### Top 10 títulos más vendidos

```sql
SELECT titulo, SUM(cantidad) as total_vendido
FROM ventas
GROUP BY titulo
ORDER BY total_vendido DESC
LIMIT 10;
```

| titulo                    | total_vendido |
|---------------------------|---------------|
| La torre en la niebla     | 1.311         |
| La ventana boreal         | 918           |
| El jardín de cristal      | 859           |
| La ruta luminoso          | 855           |
| La ventana de las sombras | 841           |
| La isla de jade           | 832           |
| El puente luminoso        | 708           |
| El cielo lejano           | 630           |
| La senda nocturno         | 545           |
| La voz de cristal         | 544           |

### ¿Qué títulos tienen ventas mayores al stock disponible?

```sql
SELECT DISTINCT v.titulo,
       SUM(v.cantidad) as total_vendido,
       SUM(ss.stock_unidades) as stock_total
FROM ventas v
LEFT JOIN stock_sucursales ss ON v.articulo = ss.id_articulo
GROUP BY v.titulo
HAVING SUM(v.cantidad) > COALESCE(SUM(ss.stock_unidades), 0)
ORDER BY total_vendido DESC;
```

> Se identificaron 20 títulos con ventas que superan su stock, indicando riesgo de quiebre. El caso más crítico: "La ventana de las sombras" (48.971 vendidas vs 24.922 en stock).

---

## El dataset

Datos anonimizados de una distribuidora editorial argentina.

| Tabla | Filas | Descripción |
|-------|-------|-------------|
| `ventas` | 9.085 | Ventas por sucursal, período, título y categoría |
| `stock_editorial` | 2.100 | Stock consolidado por artículo (depósito, cadena, terceros) |
| `stock_sucursales` | 27.798 | Stock por artículo y sucursal (normalizado de formato wide) |

**Esquema completo:**

- **ventas**: `sucursal`, `fecha`, `categoria`, `articulo`, `isbn`, `titulo`, `cantidad`
- **stock_editorial**: `id_articulo`, `isbn`, `altaprod`, `titulo`, `comercializacion`, `categoria`, `stock_depo`, `val_depo`, `stock_cad`, `val_cad`, `stock_3`, `val_3`, `total_uni`, `total_val`
- **stock_sucursales**: `id_articulo`, `isbn`, `ideditor`, `desceditor`, `titulo`, `consignado`, `deposito`, `fallados`, `disponible`, `cadena`, `total`, `costo`, `pvp`, `seccion`, `grupo`, `familia`, `subfamilia`, `peso`, `sucursal`, `stock_unidades`

---

## Limitaciones y próximos pasos

**Limitaciones:**
- Funciona bien con preguntas sobre las tablas conocidas; preguntas muy ambiguas o que requieren conocimiento externo pueden generar SQL incorrecto.
- Las queries con JOINs complejos entre las 3 tablas a veces necesitan el reintento automático.
- Sin memoria conversacional — cada pregunta es independiente.

**Próximos pasos:**
- RAG sobre metadata para soportar esquemas más grandes sin saturar el contexto.
- Demo en vivo hosteada (Streamlit Cloud o similar).
- Caché de queries frecuentes para reducir llamadas a la API.
- Soporte multi-turn: que el agente recuerde el contexto de preguntas anteriores.

---

## Cómo correrlo

```bash
# 1. Clonar el repo
git clone https://github.com/tu-usuario/text-to-sql-ventas.git
cd text-to-sql-ventas

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Configurar la API key
cp .env.example .env
# Editar .env y poner tu ANTHROPIC_API_KEY

# 4. Crear la base de datos
python src/setup_db.py

# 5. Usar el agente por CLI
python -m src.cli

# 6. (Opcional) Levantar la demo web
pip install streamlit
streamlit run app/streamlit_app.py
```

---

## Stack

`Python` · `Claude API` · `SQLite` · `pandas` · `Streamlit`
