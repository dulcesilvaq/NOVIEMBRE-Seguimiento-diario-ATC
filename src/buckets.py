"""Lógica de agrupación: columnas por Fecha Vencimiento y filas por PERIODO.

Regla clave (confirmada con el usuario):
  * ADELANTO        = vencimiento >= 1er día del mes siguiente al mes actual.
  * ADELANTO CIERRE = Adelanto con vencimiento DESPUÉS de la fecha de cierre (ej. desde el 06/11),
    MÁS los que vencen entre el día 1 y la fecha de cierre (01/11 al 05/11) pero se pagaron
    en un mes ANTERIOR al del vencimiento (ej. pago en octubre).
    NO entran los que vencen del 01/11 al 05/11 y se pagaron en noviembre.
    Es un SUBCONJUNTO de Adelanto, por eso NO se suma en el Total general.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta

MESES = ["ene", "feb", "mar", "abr", "may", "jun",
         "jul", "ago", "sept", "oct", "nov", "dic"]

MESES_COMPLETO = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                  "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

ANT, ADEL, CIERRE, TOTAL = "ANT", "ADEL", "CIERRE", "TOTAL"
SIN_PERIODO = "SIN"


@dataclass(frozen=True)
class Params:
    anio: int
    mes: int
    anio_inicio: int
    periodo_anteriores: int
    fecha_cierre: date     # último día en que ATC recibe pagos del ciclo (ej. 05/11)
    adelanto_desde: date   # primer día de Adelanto
    cierre_desde: date     # día siguiente al cierre (ej. 06/11)
    periodos_filas: tuple = (2023, 2024, 2025, 2026)   # filas FIJAS de las tablas 2 y 3 (además de Anteriores)


@dataclass(frozen=True)
class Columna:
    key: str
    label: str
    estilo: str  # anio | mes | adelanto | cierre | total


def _leer_fecha(valor, campo: str) -> date:
    """Acepta fecha de Excel/YAML (2026-11-05), texto '2026-11-05' o '05/11/2026'."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(valor).strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"config.yaml: '{campo}' debe ser una fecha (ej. 2026-11-05 o 05/11/2026)")


def params_desde_config(cfg: dict) -> Params:
    try:
        anio, mes = (int(x) for x in str(cfg["mes_actual"]).split("-"))
    except Exception as e:
        raise ValueError("config.yaml: 'mes_actual' debe tener formato AAAA-MM (ej. 2026-10)") from e
    if not 1 <= mes <= 12:
        raise ValueError("config.yaml: el mes de 'mes_actual' debe estar entre 01 y 12")

    if cfg.get("fecha_cierre") in (None, ""):
        raise ValueError("config.yaml: falta 'fecha_cierre' (ej. \"2026-11-05\")")
    fecha_cierre = _leer_fecha(cfg["fecha_cierre"], "fecha_cierre")
    a_sig, m_sig = (anio + 1, 1) if mes == 12 else (anio, mes + 1)
    adelanto_desde = date(a_sig, m_sig, 1)
    # Filas de PERIODO fijas (para poder pegar siempre en el mismo formato): Anteriores + estos años.
    # Todo PERIODO anterior al primer año de la lista se agrupa en "Anteriores".
    try:
        filas = tuple(sorted(int(x) for x in cfg.get("periodos_filas", (2023, 2024, 2025, 2026))))
    except (TypeError, ValueError) as e:
        raise ValueError("config.yaml: 'periodos_filas' debe ser una lista de años (ej. [2023, 2024, 2025, 2026])") from e
    if not filas:
        raise ValueError("config.yaml: 'periodos_filas' no puede estar vacía")
    periodo_anteriores = max(int(cfg["periodo_anteriores"]), filas[0] - 1)
    if fecha_cierre < adelanto_desde:
        raise ValueError(
            f"config.yaml: 'fecha_cierre' ({fecha_cierre:%d/%m/%Y}) debe ser posterior al fin de "
            f"'mes_actual' (>= {adelanto_desde:%d/%m/%Y})"
        )

    return Params(
        anio=anio, mes=mes, fecha_cierre=fecha_cierre,
        anio_inicio=int(cfg["anio_inicio_columnas"]),
        periodo_anteriores=periodo_anteriores,
        periodos_filas=filas,
        adelanto_desde=adelanto_desde,
        cierre_desde=fecha_cierre + timedelta(days=1),
    )


def columnas(p: Params) -> list[Columna]:
    """Columnas de las tablas, de izquierda a derecha."""
    cols = [Columna(ANT, "Anteriores", "anio")]
    for y in range(p.anio_inicio, p.anio):
        cols.append(Columna(str(y), str(y), "anio"))
    for m in range(1, p.mes + 1):
        cols.append(Columna(f"{p.anio}-{m:02d}", f"{MESES[m-1]}-{str(p.anio)[2:]}", "mes"))
    cols.append(Columna(ADEL, "Adelanto", "adelanto"))
    cols.append(Columna(CIERRE, "Adelanto Cierre", "cierre"))
    cols.append(Columna(TOTAL, "Total general", "total"))
    return cols


def clave_vencimiento(d: date, p: Params) -> str:
    """Columna (única) a la que pertenece una Fecha Vencimiento."""
    if d >= p.adelanto_desde:
        return ADEL
    if d.year < p.anio_inicio:
        return ANT
    if d.year < p.anio:
        return str(d.year)
    return f"{d.year}-{d.month:02d}"


def es_cierre(venc: date, pago: date, p: Params) -> bool:
    """Adelanto Cierre.

    * Vence después de la fecha de cierre (desde el 06/11): entra siempre.
    * Vence del 01/11 a la fecha de cierre (01/11-05/11): entra solo si se pagó en un mes
      ANTERIOR al del vencimiento (ej. pago en octubre). Pagadas en noviembre: no entran.
    """
    if venc < p.adelanto_desde:
        return False
    if venc >= p.cierre_desde:
        return True
    return (pago.year, pago.month) < (venc.year, venc.month)


def clave_periodo(periodo: int, p: Params):
    if periodo < 0:
        return SIN_PERIODO
    return ANT if periodo <= p.periodo_anteriores else int(periodo)


def etiqueta_periodo(k) -> str:
    return "Anteriores" if k == ANT else ("Sin PERIODO" if k == SIN_PERIODO else str(k))


def clasificar(df, p: Params):
    """Agrega columnas auxiliares: col_venc, es_cierre, per_key."""
    out = df.copy()
    fechas = out["Fecha Vencimiento Actual"].dt.date
    pagos = out["Fecha Pago"].dt.date
    out["col_venc"] = fechas.map(lambda d: clave_vencimiento(d, p))
    out["es_cierre"] = [es_cierre(v, g, p) for v, g in zip(fechas, pagos)]
    # Adelanto dividido SOLO por fecha de vencimiento: del día 1 a la fecha de cierre / desde el día siguiente
    out["es_adelanto"] = out["col_venc"] == ADEL
    out["venc_1_al_cierre"] = out["es_adelanto"] & fechas.map(lambda d: d <= p.fecha_cierre)
    out["venc_desde_cierre"] = out["es_adelanto"] & fechas.map(lambda d: d >= p.cierre_desde)
    out["per_key"] = out["PERIODO"].map(lambda x: clave_periodo(int(x), p))
    return out


def descripcion_ciclo(p: Params) -> str:
    return (f"Adelanto = vencimiento desde {p.adelanto_desde:%d/%m/%Y}. "
            f"Adelanto Cierre = Adelanto con vencimiento desde {p.cierre_desde:%d/%m/%Y} (cierre ATC: "
            f"{p.fecha_cierre:%d/%m/%Y}), más los que vencen del {p.adelanto_desde:%d/%m} al "
            f"{p.fecha_cierre:%d/%m} pagados en un mes anterior al del vencimiento. "
            f"Los que vencen del {p.adelanto_desde:%d/%m} al {p.fecha_cierre:%d/%m} y se pagaron "
            f"en el mismo mes NO entran. Adelanto Cierre NO se suma al Total general.")


def etiquetas_adelantos(p: Params) -> dict:
    """Encabezados del resumen de adelantos (se ajustan solos al mes y a la fecha de cierre)."""
    m = p.adelanto_desde.month
    return {
        "adelanto": "Adelanto mes natural",
        "cierre": "Adelanto considerado meta en reportes anteriores",
        "v15": f"Pagos vencidos del 1 al {p.fecha_cierre.day} de {MESES[m - 1]}",
        "v6": f"Pagos vencidos a partir del {p.cierre_desde.day} de {MESES_COMPLETO[p.cierre_desde.month - 1].capitalize()}",
    }


def encabezados_detalle_adelantos(p: Params) -> tuple[str, str]:
    """Encabezados de las 2 últimas columnas de DETALLE ADELANTOS (se ajustan solos al mes y al cierre)."""
    return (f"Vencidos 1-{p.fecha_cierre.day} {MESES[p.adelanto_desde.month - 1]}",
            f"Vencidos {p.cierre_desde.day} {MESES[p.cierre_desde.month - 1]} en adelante")
