# Reporte de cobranza

Lee la hoja **BASE FINAL** de un Excel y genera un Excel con:

- **`RES MES`**: suma de *Real Pago Dolar* por *Fecha Pago* (filas) y *Coordinadora Cobranza* (columnas).
- **Una hoja por asesora** (Katherin, Yamile, ...) con tres tablas listas para copiar y pegar:
  1. **Tabla 1**: por día, del 01 del mes actual al 09 del mes siguiente. Dos filas por día: `Monto` (suma) y `Numero de cuotas` (conteo). Sábados y domingos en rojo.
  2. **Tabla 2**: por PERIODO, Monto.
  3. **Tabla 3**: por PERIODO, Numero de cuotas.
- **`RESUMEN ADELANTOS`** (segunda pestaña): una fila por asesora y la fila Total, en US$, con cuatro columnas: *Adelanto mes natural* (vence desde el día 1 del mes siguiente), *Adelanto considerado meta en reportes anteriores* (Adelanto Cierre), *Pagos vencidos del 1 al 5* y *Pagos vencidos a partir del 6*. Las dos últimas se clasifican **solo por fecha de vencimiento** (sin importar cuándo se pagó) y suman exactamente el Adelanto mes natural. Los encabezados se ajustan solos al mes y a la fecha de cierre (ej. "del 1 al 5 de dic", "a partir del 6 de Diciembre").
- **`VENCIDOS 1-5 DIC`** y **`VENCIDOS 6 DIC EN ADELANTE`** (se llaman según el mes y el cierre): los pagos con vencimiento del día 1 al cierre, y desde el día siguiente al cierre en adelante, sin importar cuándo se pagaron. Cada hoja trae arriba la tabla *Fecha Pago x asesora* (estilo `RES MES`, con una fila de *Numero de cuotas*) y debajo el **detalle de los pagos uno por uno** (código, nro de cuota, vencimiento, PERIODO, monto, si cuenta en Adelanto Cierre y la fila en tu BASE). Si en el rango no hay pagos, la hoja lo dice.
- **`DETALLE`**: todas las filas consideradas en los totales, como Tabla de Excel con filtros. Incluye la asesora, las fechas, el nro de cuota, el monto, una **Categoría** (*Adelanto Cierre*, *Adelanto (sin cierre)* u *Otros*) con el **Motivo** que explica por qué, la **columna de vencimiento** en que cae cada pago, y su clasificación en SI/NO: **Adelanto**, **Adelanto Cierre**, **Vencido del 1 al 5** y **Vencido desde el 6**, el código de matrícula y el **número de fila en tu hoja BASE**. Para auditar un total, filtra aquí (ej. Hoja = Katherin, Columna de vencimiento = oct-26) y compara la suma; o inserta tu propia tabla dinámica sobre esta tabla (Insertar -> Tabla dinámica).

  Las columnas de las tres tablas se agrupan por *Fecha Vencimiento Actual*:
  `Anteriores | 2022 | 2023 | 2024 | 2025 | ene-26 ... mes actual | Adelanto | Adelanto Cierre | Total general`

## Dos formas de usarlo

### A) App web con Streamlit (recomendada)

```bash
pip install -r requirements.txt
streamlit run app.py
```

En la pantalla: subes el Excel, eliges el **mes actual**, la **fecha de cierre** y el **orden / nombres de las asesoras**, pulsas *Generar reporte* y descargas el Excel. Nada de esto requiere tocar Git.

- Los ajustes hechos en pantalla valen para esa sesión. Cada vez que abras la app, parte de los valores de `config.yaml`.
- Si quieres que unos ajustes queden como los iniciales, usa el botón *Descargar config.yaml* y reemplaza el `config.yaml` del repositorio (el archivo descargado no trae los comentarios).

### B) Línea de comandos

Crea una carpeta `data/` (no viene en el repositorio) y guarda ahí tu Excel; esa carpeta y `output/` nunca se suben a GitHub.

```bash
python -m src.main --input data/base.xlsx
python -m src.main --input data/base.xlsx --solo Katherin     # una sola asesora (pruebas)
```

Toma los parámetros de `config.yaml` y deja el Excel en `output/Reporte_cobranza_AAAA-MM.xlsx` (la carpeta `output/` se crea sola).

## Publicar la app (Streamlit Community Cloud)

1. Sube este repositorio a GitHub (la base **no** se sube; ver "Datos personales").
2. Entra a <https://share.streamlit.io>, inicia sesión con GitHub y elige **Create app**.
3. Selecciona el repositorio, la rama `main` y como archivo principal `app.py`. Pulsa **Deploy**.
4. Cada cambio que subas al repositorio actualiza la app.

## Parámetros (`config.yaml` y pantalla de la app)

| Campo | Qué hace |
|---|---|
| `mes_actual` | `"2026-10"`. En noviembre: `"2026-11"`. Define hasta qué mes hay columnas y desde cuándo empieza Adelanto (día 1 del mes siguiente). |
| `fecha_cierre` | Fecha de cierre de ATC en el mes siguiente, p. ej. `"2026-11-05"`. Cámbiala si el cierre es el 6, 7, etc. Ver regla de Adelanto Cierre. |
| `tabla1_hasta_dia` | La Tabla 1 llega hasta este día del mes siguiente (9 -> 01/10 a 09/11). Solo formato. |
| `asesoras` | **Orden** = orden de columnas en `RES MES` y de pestañas. `nombre` = exactamente como viene en la base; `hoja` = nombre corto de la pestaña. |
| `fusionar` | Suma la cartera de una asesora a la de otra: `"nombre_origen": "nombre_destino"` (nombres idénticos a la base; el destino debe estar en la lista de `asesoras`). En la app es el cuadro *Unir asesoras*. En `DETALLE` se conserva la coordinadora original. |
| `anio_inicio_columnas` | Vencimientos anteriores a este año -> columna *Anteriores* (2022 -> 2021 hacia atrás). |
| `periodo_anteriores` | `PERIODO` <= este valor -> fila *Anteriores* (2019 agrupa 2018, 2019, 2020, 2021...). |

## Reglas

- Solo se cuentan filas con `Real Pago Dolar` > 0.
- **Adelanto** = vencimiento desde el día 1 del mes siguiente al mes actual.
- **Adelanto Cierre** (parte de Adelanto) incluye:
  1. los que vencen **después** de la fecha de cierre (cierre 05/11 -> desde el 06/11), sin importar cuándo se pagaron; y
  2. los que vencen **del día 1 hasta la fecha de cierre** (01/11 al 05/11) **solo si se pagaron en un mes anterior** al del vencimiento (ej. pago en octubre).

  **No entran** los que vencen del 01/11 al 05/11 y se pagaron en noviembre (mismo mes). Siguen contando en Adelanto.
  Como es un subconjunto de Adelanto, **no se suma al Total general** (así coincide con `RES MES`).
- Los días sin pagos quedan vacíos. Las filas de PERIODO son las mismas en todas las hojas.
- `PERIODO` puede venir como número o como el texto `Anteriores`; ambos van a la fila **Anteriores**.
- Si en la base aparece una coordinadora que no está en tu lista, se avisa y se agrega al final (con un nombre de hoja que no se repite).
- Si una asesora de tu lista no tiene pagos en la base, se avisa y no se genera su hoja ni su columna en `RES MES`.
- Si el **mes actual** no corresponde a las fechas de pago de la base, el primer aviso indica qué mes y qué fecha de cierre poner (el ciclo va del 6 de un mes al 5 del siguiente).

## Datos personales

La base trae DNI, correos y teléfonos. Este programa **solo lee 7 columnas**: `PERIODO`, `Fecha Vencimiento Actual`, `Fecha Pago`, `Real Pago Dolar`, `Coordinadora Cobranza` y, si existen, `Código Matricula` y `Nro Cuota` (solo para rastrear cada pago en `DETALLE` y en las hojas de vencidos). El resto nunca se carga, así que puedes dejar la base completa o borrar esas columnas antes de subirla a la app.

`.gitignore` excluye `*.xlsx`, `*.csv`, `data/` y `output/`: **no subas la base a GitHub**.

## Pruebas

```bash
pip install -r requirements-dev.txt
python -m pytest -q
```

## Estructura

```
app.py               <- app web (Streamlit)
config.yaml          <- valores iniciales
src/pipeline.py      <- flujo completo (lo usan la app y la consola)
src/main.py          <- línea de comandos
src/loader.py        <- lee y valida la base (solo 5 columnas)
src/buckets.py       <- reglas de Anteriores / meses / Adelanto / Adelanto Cierre
src/tables.py        <- RES MES y tablas por asesora
src/excel_writer.py  <- formato del Excel
tests/               <- pruebas
```
