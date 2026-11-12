"""Pipeline completo: base -> Excel. Lo usan la línea de comandos (main.py) y la app (app.py)."""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd
from openpyxl import Workbook

from . import buckets, loader, tables
from .excel_writer import (escribir_detalle, escribir_hoja_asesora, escribir_res_mes,
                           escribir_resumen_adelantos, escribir_vencidos)


def _ciclo_sugerido(df: pd.DataFrame) -> pd.Period:
    """Mes del ciclo al que parecen pertenecer los pagos (el ciclo va de ~6 a ~5 del mes siguiente).
    Se usan los pagos con día >= 8, que pertenecen sin duda a ese mes."""
    d = df["Fecha Pago"]
    base = d[d.dt.day >= 8]
    return (base if len(base) else d).dt.to_period("M").mode().iloc[0]


def _fusionar(df: pd.DataFrame, fusionar: dict, adv: list[str]) -> pd.DataFrame:
    """Suma la cartera de una asesora a la de otra: {nombre_origen: nombre_destino}.
    Se conserva el nombre original en la columna 'Coordinadora original' (visible en DETALLE)."""
    df = df.copy()
    df["Coordinadora original"] = df["Coordinadora Cobranza"]
    for origen, destino in (fusionar or {}).items():
        origen, destino = str(origen).strip(), str(destino).strip()
        if not origen or not destino or origen == destino:
            continue
        m = df["Coordinadora Cobranza"] == origen
        if not m.any():
            adv.append(f"Unir asesoras: '{origen}' no aparece en la base (el nombre debe ser idéntico); no se unió.")
            continue
        df.loc[m, "Coordinadora Cobranza"] = destino
        adv.append(f"Unidas a la cartera de '{destino}': '{origen}' ({int(m.sum())} pagos, "
                   f"US$ {df.loc[m, 'Real Pago Dolar'].sum():,.2f}).")
    return df


@dataclass
class Resultado:
    libro: Workbook
    advertencias: list[str]
    params: buckets.Params
    res_mes: pd.DataFrame
    dias_base: pd.DatetimeIndex
    dias_tabla1: pd.DatetimeIndex
    n_filas: int
    # hoja -> {"monto": DataFrame, "cuotas": DataFrame} (tablas por PERIODO, columnas con su etiqueta)
    vistas: dict = field(default_factory=dict)
    resumen: pd.DataFrame | None = None     # resumen de adelantos por asesora (con fila Total)
    vencidos: dict = field(default_factory=dict)   # {"1-5": pivot, "6+": pivot} (Fecha Pago x asesora)


def generar(entrada, cfg: dict, solo: list[str] | None = None) -> Resultado:
    """`entrada`: ruta o archivo subido. `cfg`: dict con la misma forma que config.yaml."""
    p = buckets.params_desde_config(cfg)
    adv: list[str] = []

    df = loader.cargar_base(entrada, cfg["hojas_base"], adv)
    df = _fusionar(df, cfg.get("fusionar"), adv)
    df = buckets.clasificar(df, p)
    asesoras = tables.orden_asesoras(df, cfg, adv)
    cols = buckets.columnas(p)
    etiquetas = {c.key: c.label for c in cols}

    dias_base = tables.dias_del_rango(df)                        # RES MES
    dias = tables.dias_tabla1(p, int(cfg["tabla1_hasta_dia"]))   # Tabla 1 de cada asesora
    fuera = df[(df["Fecha Pago"] < dias[0]) | (df["Fecha Pago"] > dias[-1])]
    sug = _ciclo_sugerido(df)
    if (sug.year, sug.month) != (p.anio, p.mes):
        sig = sug + 1
        adv.insert(0, (
            f"REVISA EL CICLO: los pagos de la base van del {df['Fecha Pago'].min():%d/%m/%Y} al "
            f"{df['Fecha Pago'].max():%d/%m/%Y} y parecen del ciclo de {sug.year}-{sug.month:02d}, pero el mes "
            f"actual está en {p.anio}-{p.mes:02d}. Pon mes actual = {sug.year}-{sug.month:02d} y la fecha de "
            f"cierre en el mes siguiente (ej. 05/{sig.month:02d}/{sig.year})."))
    elif len(fuera):
        adv.append(f"{len(fuera)} pago(s) con Fecha Pago fuera de {dias[0]:%d/%m/%Y}-{dias[-1]:%d/%m/%Y}: "
                   f"no salen en la Tabla 1 (sí en RES MES y tablas 2 y 3). Revisa el mes actual.")
    if (p.fecha_cierre.year, p.fecha_cierre.month) != (p.adelanto_desde.year, p.adelanto_desde.month):
        adv.append(f"La fecha de cierre ({p.fecha_cierre:%d/%m/%Y}) no cae en el mes siguiente al mes actual; revisa.")
    periodos = tables.filas_periodo(p)
    fuera_filas = df[(df["per_key"] != buckets.SIN_PERIODO) & ~df["per_key"].isin(periodos)]
    if len(fuera_filas):
        anios = sorted({int(x) for x in fuera_filas["per_key"]})
        adv.append(f"{len(fuera_filas)} pago(s) con PERIODO {anios}, que no está en las filas fijas "
                   f"{[buckets.etiqueta_periodo(x) for x in periodos]}: no salen en las tablas 2 y 3. "
                   f"Agrega ese año a 'periodos_filas' si lo necesitas.")

    wb = Workbook()
    res = tables.res_mes(df, asesoras)
    escribir_res_mes(wb, res)
    resumen = tables.resumen_adelantos(df, asesoras, p)
    escribir_resumen_adelantos(wb, resumen)

    # hojas de pagos vencidos: del 1 al cierre y desde el día siguiente al cierre
    n15, n6 = buckets.nombres_hojas_vencidos(p)
    if {n15.lower(), n6.lower()} & {a["hoja"].lower() for a in asesoras}:
        raise ValueError("El nombre de hoja de una asesora coincide con el de las hojas de vencidos; usa otro.")
    t15, t6 = buckets.titulos_hojas_vencidos(p)
    vencidos = {}
    for clave, nombre, mascara, titulo, indice in (("1-5", n15, df["venc_1_al_cierre"], t15, 2),
                                                   ("6+", n6, df["venc_desde_cierre"], t6, 3)):
        piv = tables.res_mes(df[mascara], asesoras, dias_base)
        escribir_vencidos(wb, indice, nombre, titulo, piv, tables.lista_pagos(df, asesoras, mascara), p)
        vencidos[clave] = piv

    vistas = {}
    for a in asesoras:
        if solo is not None and a["hoja"] not in solo:
            continue
        d = df[df["Coordinadora Cobranza"] == a["nombre"]]
        per_m = tables.matriz(d, "per_key", periodos, cols, "sum")
        per_c = tables.matriz(d, "per_key", periodos, cols, "count")
        escribir_hoja_asesora(
            wb, a["hoja"], a["hoja"], p, cols, dias,
            tables.matriz(d, "Fecha Pago", list(dias), cols, "sum"),
            tables.matriz(d, "Fecha Pago", list(dias), cols, "count"),
            per_m, per_c, periodos, cfg["etiqueta_tabla"],
        )
        vistas[a["hoja"]] = {
            k: v.rename(columns=etiquetas).rename(index=buckets.etiqueta_periodo)
            for k, v in (("monto", per_m), ("cuotas", per_c))
        }

    escribir_detalle(wb, df, asesoras, cols, p)

    return Resultado(wb, adv, p, res, dias_base, dias, len(df), vistas, resumen, vencidos)
