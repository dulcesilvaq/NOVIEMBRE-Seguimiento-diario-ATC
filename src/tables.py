"""Construcción de las tablas (sin formato): RES MES y tablas por asesora."""
from __future__ import annotations

import re

import pandas as pd

from .buckets import ADEL, ANT, CIERRE, TOTAL, Columna, Params, columnas, etiquetas_adelantos

VAL = "Real Pago Dolar"


def dias_del_rango(df: pd.DataFrame) -> pd.DatetimeIndex:
    """Todos los días entre la primera y la última Fecha Pago de la base."""
    return pd.date_range(df["Fecha Pago"].min(), df["Fecha Pago"].max(), freq="D")


def dias_tabla1(p: Params, hasta_dia: int) -> pd.DatetimeIndex:
    """Tabla 1 de cada asesora: del 01 del mes actual al día `hasta_dia` del mes siguiente.

    Ej.: mes actual 2026-10, hasta_dia 9  ->  01/10/2026 a 09/11/2026.
    (Solo por formato; las tablas 2 y 3 usan todos los pagos de la base.)
    """
    inicio = pd.Timestamp(p.anio, p.mes, 1)
    fin = pd.Timestamp(p.adelanto_desde.year, p.adelanto_desde.month, hasta_dia)
    return pd.date_range(inicio, fin, freq="D")


def _hoja_unica(base: str, usadas: set[str]) -> str:
    """Nombre de pestaña válido y único (Excel no distingue mayúsculas, máx. 31 caracteres)."""
    limpio = re.sub(r"[\[\]:*?/\\]", "", base).strip()[:31] or "Hoja"
    cand, k = limpio, 2
    while cand.lower() in usadas:
        suf = f" {k}"
        cand, k = limpio[:31 - len(suf)] + suf, k + 1
    usadas.add(cand.lower())
    return cand


def orden_asesoras(df: pd.DataFrame, cfg: dict, advertencias: list[str]) -> list[dict]:
    """Asesoras en el orden de la config.

    * Las de la config que NO tienen pagos en esta base se omiten (se avisa).
    * Las coordinadoras de la base que NO están en la config se agregan al final (se avisa),
      con un nombre de hoja que no choque con ninguno existente.
    """
    todas = [dict(a) for a in cfg["asesoras"]]
    hojas = [a["hoja"].lower() for a in todas]
    if len(set(hojas)) != len(hojas):
        raise ValueError("Hay nombres de hoja repetidos en la lista de asesoras.")
    reservadas = {"res mes", "detalle", "resumen adelantos"}
    if reservadas & set(hojas):
        raise ValueError("Los nombres de hoja 'RES MES', 'RESUMEN ADELANTOS' y 'DETALLE' están reservados; usa otro para la asesora.")
    usadas = set(hojas) | reservadas

    presentes = set(df["Coordinadora Cobranza"])
    sin_datos = [a["nombre"] for a in todas if a["nombre"] not in presentes]
    lista = [a for a in todas if a["nombre"] in presentes]
    if sin_datos:
        advertencias.append("Asesoras de tu lista SIN pagos en esta base (no se generó su hoja ni su columna en "
                            "RES MES): " + "; ".join(sin_datos) + ".")

    conocidas = {a["nombre"] for a in todas}
    for n in sorted(presentes - conocidas):
        apellido, _, nombres = n.partition(",")
        corto = (nombres or apellido).strip().split(" ")[0]
        base = corto if corto.lower() not in usadas else f"{corto} {apellido.strip()[:1]}"
        hoja = _hoja_unica(base, usadas)
        advertencias.append(f"Coordinadora NO está en tu lista: '{n}'. Se agregó al final (hoja '{hoja}'). "
                            f"Agrégala a la lista de asesoras para controlar su posición.")
        lista.append({"nombre": n, "hoja": hoja})
    return lista


def res_mes(df: pd.DataFrame, asesoras: list[dict], dias=None) -> pd.DataFrame:
    """Filas = Fecha Pago (todos los días), columnas = coordinadoras, valor = suma."""
    dias = dias_del_rango(df) if dias is None else dias
    nombres = [a["nombre"] for a in asesoras]
    if df.empty:
        t = pd.DataFrame(index=dias, columns=nombres, dtype=float)
    else:
        t = (df.groupby(["Fecha Pago", "Coordinadora Cobranza"])[VAL].sum()
               .unstack("Coordinadora Cobranza")
               .reindex(index=dias, columns=nombres))
    t["Total general"] = t.sum(axis=1, min_count=1)
    t.loc["Total general"] = t.sum(axis=0, min_count=1)
    return t


def lista_pagos(df: pd.DataFrame, asesoras: list[dict], mascara) -> pd.DataFrame:
    """Pagos (uno por fila) que cumplen `mascara`, ordenados por Fecha Pago y asesora."""
    orden = {a["nombre"]: i for i, a in enumerate(asesoras)}
    hoja_de = {a["nombre"]: a["hoja"] for a in asesoras}
    d = df[mascara].copy()
    d["hoja"] = d["Coordinadora Cobranza"].map(hoja_de)
    d["_o"] = d["Coordinadora Cobranza"].map(orden)
    return d.sort_values(["Fecha Pago", "_o", "Fecha Vencimiento Actual"], kind="stable").drop(columns="_o")


def matriz(df: pd.DataFrame, fila_col: str, filas: list, cols: list[Columna], agg: str) -> pd.DataFrame:
    """Tabla con columnas = Columna.key (incl. CIERRE y TOTAL) y filas = `filas`.

    TOTAL = suma de columnas de vencimiento (SIN Adelanto Cierre, que es subconjunto de Adelanto).
    """
    keys = [c.key for c in cols if c.key not in (CIERRE, TOTAL)]
    desconocidas = set(df["col_venc"]) - set(keys)
    if desconocidas:
        raise ValueError(f"Fechas de vencimiento fuera de las columnas definidas: {desconocidas}")

    base = (df.groupby([fila_col, "col_venc"])[VAL].agg(agg)
              .unstack("col_venc")
              .reindex(index=filas, columns=keys)
              .fillna(0))
    cierre = (df[df["es_cierre"]].groupby(fila_col)[VAL].agg(agg)
              .reindex(filas).fillna(0))
    out = base.copy()
    out[CIERRE] = cierre
    out[TOTAL] = base.sum(axis=1)
    return out[[c.key for c in cols]]


def fila_total(m: pd.DataFrame) -> pd.Series:
    return m.sum(axis=0)


def filas_periodo(p: Params) -> list:
    """Filas FIJAS de las tablas 2 y 3: Anteriores + los años configurados (aunque no haya pagos en alguno)."""
    return [ANT, *p.periodos_filas]


def resumen_adelantos(df: pd.DataFrame, asesoras: list[dict], p: Params) -> pd.DataFrame:
    """Una fila por asesora (nombre corto de hoja) + fila Total. Montos en US$ (Real Pago Dolar).

    * Adelanto mes natural : vencimiento desde el día 1 del mes siguiente
    * Adelanto considerado meta en reportes anteriores : Adelanto Cierre (subconjunto de Adelanto)
    * Pagos vencidos del 1 al <cierre>      : vence del día 1 a la fecha de cierre (sin importar cuándo se pagó)
    * Pagos vencidos a partir del <cierre+1> : vence desde el día siguiente al cierre (sin importar cuándo se pagó)
    """
    v = df["Real Pago Dolar"]
    et = etiquetas_adelantos(p)
    d = pd.DataFrame({
        "Coordinadora Cobranza": df["Coordinadora Cobranza"],
        et["adelanto"]: v.where(df["es_adelanto"], 0.0),
        et["cierre"]: v.where(df["es_cierre"], 0.0),
        et["v15"]: v.where(df["venc_1_al_cierre"], 0.0),
        et["v6"]: v.where(df["venc_desde_cierre"], 0.0),
    })
    g = d.groupby("Coordinadora Cobranza").sum()
    out = g.reindex([a["nombre"] for a in asesoras]).fillna(0.0)
    out.index = [a["hoja"] for a in asesoras]
    out.loc["Total"] = out.sum()
    return out
