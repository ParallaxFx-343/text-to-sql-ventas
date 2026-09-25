# Notas operativas — Editorial El Ateneo / Grupo ILHSA

Memoria de trabajo para el análisis operativo (reposiciones, progresiones,
liquidaciones, servicio de novedades). Se carga automáticamente en cada sesión.

## Sucursales

### Altas pendientes de confirmar

| Sucursal | Nro. | Estado | Pendiente |
|---|---|---|---|
| **DEVOTO** | **56** (a confirmar) | Nueva — informada 08/09/2026 | Franco envía el Excel con los datos (nunca llegó) |

El número 56 lo dio Franco de memoria ("creo"). **Confirmar contra el maestro
antes de usarlo en cualquier archivo de carga**, sobre todo en los CSV de
reposición y en los TXT de PAEC, donde un número equivocado manda mercadería
a otra boca.

Estado al 23/09/2026: DEVOTO aparece como columna en el SII desde el
11/09 (76 columnas de sucursal), pero con **0 unidades en los 1.230
artículos** en todos los SII de septiembre. Franco decidió **"devoto no
va"**: quedó afuera de todas las acciones de septiembre. Mientras siga en
cero, excluirla y avisar. El SII no trae números de sucursal, así que no
sirve para confirmar el 56.

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

El SII traía 75 columnas de sucursal; desde el 11/09/2026 trae **76** (se sumó
DEVOTO, ver arriba). Se excluyen siempre estas **14**:

```
CENTRAL                          FERIA ITINERANTE EDITORIAL
CANALES ALTERNATIVOS             FERIA ITINERANTE N 1 - REGION 3
CABILDO                          FERIA ITINERANTE N 2 - REGION 3
PRUEBA 734                       FERIA ITINERANTE REGION 1
EXPERIENCIA G.SPLENDID (SUC)     FERIA ITINERANTE REGION 2
EXPERIENCIA GRAND SPLENDID       FERIA ITINERANTE REGION 4
Futuro local Cordoba             FERIA YENNY ITINERANTES
```

76 − 14 = **62 activas** contando DEVOTO; **61 operativas** sin ella. De las 61,
60 se reparten en 4 grupos de 15 y **ATENEO GRAND SPLENDID queda fuera de los
grupos** porque tiene circuito propio. DEVOTO todavía no tiene grupo.

El padrón completo, los 4 grupos y el orden de los TXT de PAEC están en
**`ateneo/config.json`** (commiteado; claves `orden_paec`, `grupos`,
`excluir_siempre`, `notas_sucursales`).

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

### "Un solo PAEC" (formato pedido desde septiembre 2026)

Cuando una acción va en **un solo PAEC**, el depósito prepara **por título**
(agarra los 240 de un título y los separa por local), y Franco distribuye en
pantalla local por local. Se entrega:

1. **Un solo CSV** con todos los títulos: `ISBN;cantidad total`, sin
   encabezado, CRLF. No un CSV por sucursal (lo pidió explícitamente dos veces).
2. **TXT de desglose por título**: los títulos en el mismo orden que el CSV y,
   dentro de cada título, las sucursales en **orden alfabético** con su
   cantidad. Encabezado de cada bloque: total, ISBN, ID y cantidad de locales.
3. Si hace falta, un Excel con la lista de títulos. Si la acción tiene rótulo
   ("ROTULAR COMO DIA DE LA MADRE"), va arriba del txt.

**Si el PAEC ya se cargó, sus cantidades son fijas.** Al redistribuir con un
SII más nuevo hay que respetar el total por título que ya se cargó y avisar
qué líneas cambiarían, sin aplicarlas.

### Criterios de reposición que Franco fijó

- **Redondeo** (Día de la Madre): el pedido de cada título va al múltiplo de
  10 de arriba, y el excedente se reparte entre **ATENEO GRAND SPLENDID,
  ATENEO JURAMENTO y ATENEO FLORIDA 340** ("pueden recibir un poquito más").
  Nunca redondear por encima del disponible de depósito.
- **"Reponer lo que vendieron"** (reposición del 21 al 25): se repone la
  columna Ventas, salvo que el stock del local ya cubra **más de 3 veces** lo
  vendido (≈ mes y medio de venta). Corte estricto: con 3,0 se repone.
- **"Dejame solo lo que se pueda servir"**: afuera los títulos con disponible
  en 0 y también los que el depósito no puede servir completos. Se aplicó en
  la reposición del 21 al 25 y en el refuerzo de Canales del 23/09. En los
  pedidos con cantidades cerradas (Canales), primero preguntar.
- Cuando la cantidad pedida supera el disponible y Franco dice "dejalos así",
  se cargan igual y se marcan (pasó con HEDY y FIN DEL AUTOODIO).
- Pendiente de decidir: medir la cobertura **después** de reponer (que el
  local no quede con más de 3 veces la venta), en vez de antes.

### Nombres de locales en los pedidos

- **Locales chicos** (van a 3 o a 5 cuando el resto va a 8 o a 10): EZEIZA,
  TRELEW, PATIO BULLRICH, PALERMO PORTAL, BAHIA BLANCA SHOPPING.
- "Ateneo Juramento, **Floridas**, Rosario, Tucumán, La Plata y Córdoba" = los
  **7 ATENEO** del SII sin Grand Splendid (Florida 340 y 632).
- "La Plata" a secas = **LA PLATA**, no ATENEO LA PLATA; "Rosario Portal" y
  "Tucumán Portal" son locales distintos de los ATENEO.
- Cuando un archivo trae nombres cortos, verificar el mapeo cruzando su
  columna Stock contra el SII: la columna correcta coincide en 82-100% de
  los títulos y la segunda no pasa del 36%.

### Canales Alternativos

CSV `refuerzo_CANALES_ALTERNATIVOS.csv` (o `reposicion_…` si viene de
ranking con columna Pedido). Lo que no se sirve se informa con un cuadro HTML
para pegar en el cuerpo del Outlook, **con estilos inline** (Outlook ignora
`<style>`), con las columnas ISBN, título, pedido, disponible y motivo.

### Sábanas de vendedores (`{VENDEDOR}_CONSOLIDADA_dd-mm-aaaa.xlsx`)

Son siete: EGISTI, FUCITO, MARTIRENA, NAVARRO, PABLO, RODRIGUEZ y SALICE. Tienen
una hoja "Resultado 1" con 22 columnas (CUENTA, CLIENTE, SUCURSAL,
FECHAULTLIQ, CODIGO = ID de artículo, ISBN como texto, TITULO, ENTREGADA,
VENTA, DEVOL, SALDO, PVP… NOMBRE_VENDEDOR). **No** son las `_SABANAS_`, que
vienen con una hoja por cliente: esas se mandaron por error. Para filtrar por
títulos se usa el **ISBN exacto**, porque las Vacas tienen otras ediciones con
el mismo título y otro ISBN (9789500210041, 9789500213073, 9789878661735,
9789500213080, 9789878816647).

### Archivos de pedido por cliente
Nombre: `{VENDEDOR}_-_{CLIENTE}_-_{CUENTA}.xlsx`
- **Hoja1**: ISBN + Cantidad, sin encabezado, datos desde la fila 7.
- **Sheet1**: ISBN + Descripción + Cantidad, encabezado en fila 1, datos desde la fila 2.

Los clientes mandan un Excel por sucursal y tipo (`PEDIDO_21653_AJAMIL_SUC._CALAFATE.xlsx`,
`REPO_…`), con el encabezado en la fila 3, 4 o 5 según el archivo. "Mergear" =
un solo archivo por cuenta, sumando por ISBN: `ateneo/pedido_cliente.py`.
Traen ISBN con el verificador viejo (978 + ISBN-10 entero: El Príncipe
`…286769` → `9789500286763`) y a veces el mismo título en dos ediciones
(Resetea tus intestinos `…213653` y 3ª ed. `…217514`): corregir y avisar.

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

## SII — trampas del export

- **Trae una fila de totales al pie**, sin ID ni título. Hay que excluirla
  antes de sumar cualquier columna: sumando todo, el disponible del depósito
  al 22/09/2026 daba 1.268.404 en vez de 634.202. Usarla como cuadre:
  DEPOSITO y DISPONIBLE de la suma de artículos coinciden exacto con esa fila.
  La columna CADENA **no** cuadra contra ella (179.837 vs 194.382), así que
  para stock en locales se suman las columnas de sucursal activas.
- Hay filas con ISBN-10 terminado en X (texto, no número). Al convertir el
  ISBN a número quedan en NaN; hoy todas tienen stock cero.
- `DISPONIBLE` no es `DEPOSITO − FALLADOS − EMPENIADOS`: esa cuenta cierra
  solo en un tercio de los artículos. El depósito puede tener stock físico
  que no figura como disponible (Stickers Dinosaurios: 28 físicos, 2
  disponibles). Para reponer, mirar siempre `DISPONIBLE`.

## Catálogo

- **"TR" = tapa Rústica (blanda), NO tapa dura.** Verificado por precio: TR y
  RUSTICA comparten tramo ($24.500 en agosto, $25.500 en septiembre), mientras
  que los títulos sin marca van a $29.900 / $30.900.
- Los códigos de 10 dígitos que arrancan en `950` suelen ser el ISBN-13 sin el
  prefijo `978`, pero **no siempre**: algunos son ISBN-10 viejos de verdad y
  ahí hay que descartar el dígito verificador y recalcularlo. Validar siempre
  el dígito verificador antes de dar por bueno un ISBN.

## Entorno (contenedor en la nube)

- **Para ver un Excel renderizado**: `apt-get install -y libreoffice-calc` (viene
  solo el núcleo y sin él no abre ningún xlsx), después
  `soffice -env:UserInstallation=file://$PWD/_lohome --headless --convert-to pdf archivo.xlsx`
  y `pypdfium2` para pasar la página a PNG. En LibreOffice falta Calibri, así que
  la letra sale distinta, y los miles aparecen con coma.
- Si pdfplumber falla con `_cffi_backend`: `pip install --force-reinstall --no-cache-dir cffi`.
- **No se pueden transcribir audios** (no hay whisper ni ffmpeg): pedirle a
  Franco el texto.
- `/design` (Claude Design) solo lo puede lanzar Franco. No edita Excel: arma
  láminas que se exportan a PDF o PNG.
- El scratchpad se borra. Lo que tenga que sobrevivir va al repo.

## Estado al 23/09/2026

Acciones de septiembre (archivos cargados en `ateneo/historial/2026-09/`):

| Fecha | Acción | Títulos | Locales | Unidades |
|---|---|---|---|---|
| 11/09 | Promo 2x1 $28.500 | 11 | 33 | 255 |
| 15/09 | Día de la Madre (PAEC único, redondeo a 10) | 38 | 61 | 5.101 |
| 17/09 | Agenda + autores de octubre | 12 | 61 | 2.819 |
| 17/09 | Canales · ranking | 56 | — | 243 |
| 21/09 | Reposición 21 al 25 (6.201 vendidas → 4.004) | 170 | 20 | 4.004 |
| 23/09 | Canales · refuerzo (126 pedidas → 101) | 43 | — | 101 |

Total 12.523 u. y 247 títulos distintos. El faltante de depósito sumó 2.016
u.: el 76% está en PRINCIPITO, GATITO ¿POR QUE ESTAS ENOJADO?, SAGA DE LOS
ROMANOV y los dos STICKERS MAGICOS. Los Stickers no están agotados, sino en
la cadena (711 y 377) y en consignación (1.212 y 1.008).

Tablero del mes: `ateneo/historial/2026-09/Reposicion_Septiembre_2026_tablero.xlsx`
(armado con `tablero.py` a partir de los `dash_*.json`). La versión web es
el artifact "Reposición Septiembre 2026".

**Reimpresiones de octubre** (16 títulos). La lista con ISBN e ID está en
`ateneo/historial/2026-09/reimpresiones_octubre.csv`.

Pendientes:
- DEVOTO: número de sucursal y grupo (el Excel nunca llegó).
- TAO TE CHING en Ateneo Córdoba: Franco ve 33 y el SII dice 20. Hay que ver
  qué pantalla muestra 33; si incluye mercadería en tránsito, el SII no la ve
  y se estaría reponiendo encima de envíos en viaje.
- Situación x Vendedor 08/2026: la fila "Ejercicio 2026-27" está cargada a
  mano y desactualizada; a GESUALDI (24135) le falta la fila en la hoja EM; el
  BUSCARV de AD y AE llega hasta la fila 623 y deja afuera 8 clientes.
- Servicio de Novedades de septiembre: faltan las cuotas, la fila de precios,
  el cronograma y las reimpresiones.

## Regla general de trabajo

Antes de entregar cualquier número: revisar exhaustivamente, no dar nada por
sentado y apoyarse en los archivos que pasa Franco. Si un control no cierra,
decirlo; si una hipótesis no aguanta los datos, decirlo también.
