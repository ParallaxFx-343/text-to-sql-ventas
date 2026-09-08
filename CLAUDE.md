# Notas operativas — Editorial El Ateneo / Grupo ILHSA

Memoria de trabajo para el análisis operativo (reposiciones, progresiones,
liquidaciones, servicio de novedades). Se carga automáticamente en cada sesión.

## Sucursales

### Altas pendientes de confirmar

| Sucursal | Nro. | Estado | Pendiente |
|---|---|---|---|
| **DEVOTO** | **56** (a confirmar) | Nueva — informada 08/09/2026 | Franco envía el Excel con los datos |

El número 56 lo dio Franco de memoria ("creo"). **Confirmar contra el maestro
antes de usarlo en cualquier archivo de carga**, sobre todo en los CSV de
reposición y en los TXT de PAEC, donde un número equivocado manda mercadería
a otra boca.

### Regla: toda sucursal nueva necesita fecha de apertura

Cuando se da de alta una sucursal hay que registrar su **fecha de apertura**
y tenerla a mano para los comparativos interanuales. Motivo concreto, ya
verificado en la auditoría de agosto 2026:

> El export "ventas por sucursal y proveedor" resuelve los nombres de
> sucursal con el maestro **vigente al momento de la extracción**, no con el
> maestro del período consultado. Una sucursal nueva que reutiliza un código
> de una boca vieja hereda las ventas históricas de esa boca y aparece con
> datos del año anterior que nunca existieron.

En la planilla auditada esto generó **6.460 unidades fantasma** (321 de
Editorial) repartidas en 8 sucursales, y dio vuelta el signo del titular:
publicaba +0,42% cuando el número real era −3,44%.

**Antes de publicar cualquier interanual:** para cada sucursal con fecha de
apertura posterior al inicio del período comparado, el año anterior tiene que
ser 0. Si trae valor, es dato fantasma y hay que excluirla del "mismas bocas".

### Cuadre válido vs. cuadre circular

Un cuadre solo sirve si compara **dos caminos independientes**. La planilla
que falló comparaba "Calculado" contra "Esperado" usando el mismo valor de
origen, así que daba OK siempre y dejó pasar los datos fantasma.

Cuadre correcto para interanuales: `mismas bocas + bocas nuevas = total`.

### Padrón

- El SII trae 75 columnas de sucursal; 14 son exclusiones (depósitos, ferias
  itinerantes, canales alternativos, pruebas) → **61 activas**.
- Aparte: ATENEO GRAND SPLENDID, que tiene circuito propio.
- Los grupos de reposición son 4 × 15 = 60 locales.
- Con DEVOTO el padrón pasaría a 62 activas. **Verificar al recibir el Excel.**

## Convenciones de archivos

### CSV de reposición
Separador `;`, sin encabezado, fin de línea CRLF (`\r\n`).
Nombre: `reposicion_{SLUG}.csv` / `refuerzo_{SLUG}.csv` / `quiebre_{SLUG}.csv`.

Orden de los locales: BENDU va entre MAR DEL PLATA LOS GALLEGOS y PASEO ALDREY,
para que quede alineado con el orden de los TXT.

### Archivos de pedido por cliente
Nombre: `{VENDEDOR}_-_{CLIENTE}_-_{CUENTA}.xlsx`
- **Hoja1**: ISBN + Cantidad, sin encabezado, datos desde la fila 7.
- **Sheet1**: ISBN + Descripción + Cantidad, encabezado en fila 1, datos desde la fila 2.

### Carga SILOMA
Hoja `Hoja1`, datos desde la **fila 6** (filas 1-5 vacías).
Columna 1 = ISBN como entero con formato de número `0`; columna 2 = cantidad;
columna 3 = etiqueta de liquidación (ej. `LIQUIDACIONJULIO`).
**El sistema no lee las líneas negativas.**

## Catálogo

- **"TR" = tapa Rústica (blanda), NO tapa dura.** Verificado por precio: TR y
  RUSTICA comparten tramo ($24.500 en agosto, $25.500 en septiembre), mientras
  que los títulos sin marca van a $29.900 / $30.900.
- Los códigos de 10 dígitos que arrancan en `950` suelen ser el ISBN-13 sin el
  prefijo `978`, pero **no siempre**: algunos son ISBN-10 viejos de verdad y
  ahí hay que descartar el dígito verificador y recalcularlo. Validar siempre
  el dígito verificador antes de dar por bueno un ISBN.

## Regla general de trabajo

Antes de entregar cualquier número: revisar exhaustivamente, no dar nada por
sentado y apoyarse en los archivos que pasa Franco. Si un control no cierra,
decirlo; si una hipótesis no aguanta los datos, decirlo también.
