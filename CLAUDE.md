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

El SII trae 75 columnas de sucursal. Se excluyen siempre estas **14**:

```
CENTRAL                          FERIA ITINERANTE EDITORIAL
CANALES ALTERNATIVOS             FERIA ITINERANTE N 1 - REGION 3
CABILDO                          FERIA ITINERANTE N 2 - REGION 3
PRUEBA 734                       FERIA ITINERANTE REGION 1
EXPERIENCIA G.SPLENDID (SUC)     FERIA ITINERANTE REGION 2
EXPERIENCIA GRAND SPLENDID       FERIA ITINERANTE REGION 4
Futuro local Cordoba             FERIA YENNY ITINERANTES
```

75 − 14 = **61 activas**. De esas, 60 se reparten en 4 grupos de 15 y
**ATENEO GRAND SPLENDID queda fuera de los grupos** porque tiene circuito propio.

El padrón completo, los 4 grupos y el orden de los TXT de PAEC están en
`scratchpad/stock/config.json` (claves `orden_paec`, `grupos`, `excluir_siempre`).
**Ese archivo vive en el scratchpad y se borra.** Si no está, regenerarlo desde
el SII aplicando las exclusiones de arriba.

Con DEVOTO el padrón pasaría a 62 activas, y hay que definir a qué grupo entra.
**Verificar al recibir el Excel.**

### Fechas de apertura

Aperturas posteriores al 01/07/2025, según el maestro de sucursales
(columna "Inicio Actividad"):

| Sucursal | Inicio | Marca |
|---|---|---|
| FERIA ITINERANTE REGION 1 | 10/09/2025 | YENNY |
| REMEROS (NORDELTA) | 15/09/2025 | YENNY |
| FERIA ITINERANTE REGION 2 | 17/10/2025 | YENNY |
| USHUAIA | 30/10/2025 | YENNY |
| FERIA ITINERANTE N 2 - REGION 3 | 31/10/2025 | YENNY |
| VILLA CRESPO | 20/11/2025 | YENNY |
| EXPERIENCIA GRAND SPLENDID | 01/12/2025 | EL ATENEO |
| MAR DEL PLATA M. BENDU | 02/01/2026 | YENNY |
| VILLA URQUIZA | 19/02/2026 | YENNY |
| FERIA ITINERANTE REGION 4 | 23/03/2026 | YENNY |
| EXPERIENCIA G.SPLENDID (SUC) | 01/06/2026 | EL ATENEO |
| SAN MIGUEL | 23/07/2026 | YENNY |
| FERIA ITINERANTE EDITORIAL | 27/07/2026 | EL ATENEO |
| FERIA ITINERANTE N 1 - REGION 3 | 31/07/2026 | YENNY |

Estas son las 8 que dieron **dato fantasma** en el interanual de julio 26/25
(AT25 = unidades Editorial que el export les atribuía en julio 2025, cuando
todavía no existían):

| Sucursal | Apertura | AT25 | Total 25 |
|---|---|---|---|
| MAR DEL PLATA M. BENDU | 02/01/2026 | 84 | 1.429 |
| VILLA URQUIZA | 19/02/2026 | 57 | 1.443 |
| USHUAIA | 30/10/2025 | 37 | 861 |
| REMEROS (NORDELTA) | 15/09/2025 | 30 | 554 |
| VILLA CRESPO | 20/11/2025 | 46 | 731 |
| SAN MIGUEL | 23/07/2026 | 67 | 1.187 |
| EXPERIENCIA G.SPLENDID (SUC) | 01/06/2026 | 0 | 167 |
| EXPERIENCIA GRAND SPLENDID | 01/12/2025 | 0 | 88 |

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

### Servicio de Novedades — planillas por vendedor

Una planilla por vendedor. La grilla de arriba (títulos) y la tabla de clientes
(abajo a la izquierda) **conviven en las mismas filas**, por eso hay que tener
cuidado al escribir.

| Fila | Columnas A-I | Columnas J en adelante |
|---|---|---|
| 1 | nombre del vendedor | `NOVEDADES` (celda combinada) y después `REIMPRESIONES` |
| 2 | — | títulos |
| 3 | — | **ISBN** |
| 4 | — | vacía |
| 5 | **header de la tabla de clientes** | **IDs de artículo** |
| 6 en adelante | clientes | cantidades pedidas |
| último cliente **+ 2** | — | **cuotas por vendedor** |

- El merge de la fila 1 se ajusta al ancho real: con 7 novedades es `J1:P1`,
  no `J1:S1`. Los merges de la izquierda son `A1:I4`.
- **La fila 5 se limpia todos los meses.** Los IDs no se reutilizan del mes
  anterior. Lo mismo las cuotas: nunca copiar las del mes pasado.
- La fila 5 en A-H tiene que quedar **idéntica al original** — se escribe de
  la J en adelante y no se pisa nada.
- La fila de cuotas se calcula, no se asume: es siempre última fila con cliente
  más 2 (si el último cliente está en la 65, la cuota va en la 67).

### Servicio de Novedades — planilla de Cuotas

Archivo aparte, un vendedor por fila:

| Fila | Contenido |
|---|---|
| 3 | ISBN |
| 5 | IDs (rotulada `Id`; se limpia cada mes) |
| 6 a 16 | vendedores |
| 17 | TOTAL PEDIDO (fórmula SUM) |
| 18 | Resto |
| 20 | TIRADA — relleno **verde** si tiene fecha, **amarillo** si ya está en depósito |
| 21 | Cronograma |
| 23 | ML 45% |

### Cúspide — archivo "A PROCESAR"

Hoja `Hoja3`. **Una fila por PEDIDO, no por sucursal**: una misma sucursal puede
aparecer en más de una fila si tiene más de un pedido (por ejemplo servicio de
novedades y ampliación), y **no se unifican** porque cada una lleva su propio
número de pedido.

Fila 1 = encabezado; de la columna I en adelante, **un ID de artículo por columna**.

| Col | Campo | Valor |
|---|---|---|
| A | IdCliente | `10842` |
| B | DirEnt | nro. de sucursal |
| C | tipocomer | `Q` |
| D | DirTrans | `1` |
| E | Trans Retira | `1` |
| F | NRO PEDIDO | uno distinto por fila |
| G | Lista Precio | 4862 en agosto, **4882 en septiembre** |
| H | — | vacía |
| I en adelante | cantidades | por ID de artículo |

**Locales redirigidos.** Gral. Rodríguez, Castelar y Luján entregan sin
excepción en el depósito central de **Ascasubi 3220**. Para esos:

1. `DirEnt` va con el número de Ascasubi (**149**), no el del local.
2. El `NRO PEDIDO` pasa a **texto** con el destino agregado, para que en
   depósito sepan a quién es: `1451836 LUJAN`, `1451846 GENERAL RODRIGUEZ`.

Referencia: `cuspide/ej_agosto.xlsx` (54 filas, 51 DirEnt distintos) y
`CUSPIDE_SEPTIEMBRE_26_A_PROCESAR.xlsx` (91 filas, novedades + ampliación).
En septiembre Castelar no pidió.

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
