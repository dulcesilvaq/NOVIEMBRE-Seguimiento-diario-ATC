"""Escritura del Excel con formato listo para copiar y pegar."""
from __future__ import annotations

from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .buckets import (ADEL, CIERRE, MESES, TOTAL, Columna, Params, descripcion_ciclo,
                      etiqueta_periodo)

F = "Arial"
AZUL_OSC, AZUL_CLARO = "1F3864", "DDEBF7"
borde_c = Side(style="thin", color="8EA9DB")
BORDE = Border(left=borde_c, right=borde_c, top=borde_c, bottom=borde_c)
ROJO = "FF0000"


def _fill(hexa):  # noqa: D401
    return PatternFill("solid", start_color=hexa, end_color=hexa)


def _num(v):
    """Ceros/NaN -> celda vacía."""
    if v is None or pd.isna(v) or v == 0:
        return None
    return float(v)


# ---------------------------------------------------------------- RES MES (y pivots iguales)
def _escribir_pivot(ws, tabla: pd.DataFrame, fila0: int = 1) -> int:
    """Escribe la tabla Fecha Pago x asesora (estilo tabla dinámica). Devuelve la última fila escrita."""
    ws.cell(fila0, 1, "Suma de Real Pago Dolar")
    ws.cell(fila0, 2, "Etiquetas de columna")
    for k in (1, 2):
        ws.cell(fila0, k).font = Font(name=F, bold=True, size=10)
        ws.cell(fila0, k).fill = _fill(AZUL_CLARO)

    ws.cell(fila0 + 1, 1, "Etiquetas de fila")
    for j, nombre in enumerate(tabla.columns, start=2):
        ws.cell(fila0 + 1, j, nombre)
    for j in range(1, len(tabla.columns) + 2):
        c = ws.cell(fila0 + 1, j)
        c.font = Font(name=F, bold=True, size=10)
        c.fill = _fill(AZUL_CLARO)
        c.alignment = Alignment(wrap_text=True, vertical="bottom",
                                horizontal="left" if j == 1 else "right")
        c.border = BORDE
    ws.row_dimensions[fila0 + 1].height = 58

    for i, (idx, fila) in enumerate(tabla.iterrows(), start=fila0 + 2):
        es_total = isinstance(idx, str)
        c0 = ws.cell(i, 1, idx.to_pydatetime() if not es_total else idx)
        if not es_total:
            c0.number_format = "d/m/yyyy"
        c0.alignment = Alignment(horizontal="left")
        for j, v in enumerate(fila, start=2):
            c = ws.cell(i, j, _num(v))
            c.number_format = "0"
        for j in range(1, len(tabla.columns) + 2):
            c = ws.cell(i, j)
            c.font = Font(name=F, size=10, bold=es_total or j == len(tabla.columns) + 1)
            c.border = BORDE
            if es_total:
                c.fill = _fill(AZUL_CLARO)
    return fila0 + 1 + len(tabla)


def escribir_res_mes(wb: Workbook, tabla: pd.DataFrame) -> None:
    ws = wb.active
    ws.title = "RES MES"
    _escribir_pivot(ws, tabla, 1)
    ws.column_dimensions["A"].width = 24
    for j in range(2, len(tabla.columns) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 15
    ws.freeze_panes = "B3"


# ---------------------------------------------------------------- DETALLE ADELANTOS
def escribir_detalle_adelantos(wb: Workbook, indice: int, tabla: pd.DataFrame, p: Params) -> None:
    """Lista de TODOS los pagos (una fila por pago) con SI/NO en Adelanto, Adelanto Cierre,
    Vencidos 1-5 y Vencidos 6 en adelante. Al final: nº de SI y monto de los SI por columna."""
    ws = wb.create_sheet("DETALLE ADELANTOS", indice)
    cols = list(tabla.columns)
    flags = cols[-4:]                                   # Adelanto, Adelanto Cierre, Vencidos 1-5, Vencidos 6+
    for j, nombre in enumerate(cols, start=1):
        c = ws.cell(1, j, nombre)
        c.font = Font(name=F, bold=True, size=10, color="FFFFFF")
        c.fill = _fill(AZUL_OSC)
        c.border = BORDE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 34

    verde = _fill("E2EFDA")
    for i, (_, r) in enumerate(tabla.iterrows(), start=2):
        for j, nombre in enumerate(cols, start=1):
            v = r[nombre]
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif pd.isna(v):
                v = None
            c = ws.cell(i, j, v)
            c.font = Font(name=F, size=10, bold=(v == "SI"))
            c.border = BORDE
            if nombre in ("Fecha Pago", "Fecha Vencimiento"):
                c.number_format = "dd/mm/yyyy"
                c.alignment = Alignment(horizontal="left")
            elif nombre == "Monto US$":
                c.number_format = "0.00"
            elif nombre in flags:
                c.alignment = Alignment(horizontal="center")
                if v == "SI":
                    c.fill = verde
    ult = len(tabla) + 1
    if len(tabla):
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ult}"

    # totales: nº de SI y monto de los SI por columna
    j_monto = cols.index("Monto US$") + 1
    f1, f2 = ult + 1, ult + 2
    ws.cell(f1, 1, "Total pagos / Nº de SI")
    ws.cell(f2, 1, "Monto US$ de los SI")
    ws.cell(f1, j_monto, round(float(tabla["Monto US$"].sum()), 2))
    for nombre in flags:
        j = cols.index(nombre) + 1
        mask = tabla[nombre] == "SI"
        ws.cell(f1, j, int(mask.sum()))
        ws.cell(f2, j, round(float(tabla.loc[mask, "Monto US$"].sum()), 2))
    for f in (f1, f2):
        for j in range(1, len(cols) + 1):
            c = ws.cell(f, j)
            c.font = Font(name=F, size=10, bold=True)
            c.fill = _fill(AZUL_CLARO)
            c.border = BORDE
            if j > 1:
                c.alignment = Alignment(horizontal="center" if cols[j - 1] in flags else "right")
        ws.cell(f, j_monto).number_format = "0.00"
    for nombre in flags:
        ws.cell(f2, cols.index(nombre) + 1).number_format = "0.00"
    ws.cell(f2 + 2, 1, descripcion_ciclo(p)).font = Font(name=F, italic=True, size=9, color="7F7F7F")

    anchos = {"Asesora": 14, "Código Matricula": 18, "Nro Cuota": 10, "Fecha Pago": 12, "Fecha Vencimiento": 14,
              "Monto US$": 12}
    for j, nombre in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(j)].width = anchos.get(nombre, 20)
    ws.column_dimensions["A"].width = 24
    ws.freeze_panes = "B2"


# ---------------------------------------------------------------- hoja por asesora
def _encabezado(ws, fila: int, cols: list[Columna], etiqueta: str) -> None:
    ws.row_dimensions[fila].height = 42
    for j in (1, 2):
        ws.cell(fila, j).fill = _fill(AZUL_CLARO)
        ws.cell(fila, j).border = BORDE
    c = ws.cell(fila, 2, etiqueta)
    c.font = Font(name=F, bold=True, size=10, color="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for k, col in enumerate(cols, start=3):
        c = ws.cell(fila, k, col.label)
        oscuro = col.estilo == "anio"
        c.fill = _fill(AZUL_OSC if oscuro else AZUL_CLARO)
        c.font = Font(name=F, bold=True, size=10, color="FFFFFF" if oscuro else "1F3864")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE


def _fila_valores(ws, fila: int, serie: pd.Series, cols: list[Columna], negrita=False, relleno=None) -> None:
    for k, col in enumerate(cols, start=3):
        c = ws.cell(fila, k, _num(serie[col.key]))
        c.number_format = "0"
        c.font = Font(name=F, size=10, bold=negrita or col.key == TOTAL)
        c.border = BORDE
        c.alignment = Alignment(horizontal="right")
        if relleno:
            c.fill = _fill(relleno)


def _titulo(ws, fila: int, texto: str) -> None:
    c = ws.cell(fila, 2, texto)
    c.font = Font(name=F, italic=True, size=9, color="7F7F7F")


def escribir_hoja_asesora(wb: Workbook, nombre_hoja: str, titulo: str, p: Params,
                          cols: list[Columna], dias: pd.DatetimeIndex,
                          t_monto: pd.DataFrame, t_cuotas: pd.DataFrame,
                          per_monto: pd.DataFrame, per_cuotas: pd.DataFrame,
                          periodos: list, etiqueta: str) -> None:
    ws = wb.create_sheet(nombre_hoja)
    ws["B1"] = titulo
    ws["B1"].font = Font(name=F, bold=True, size=11, color="0000FF")

    # ---- Tabla 1: por día (Monto / Numero de cuotas)
    fila = 3
    _titulo(ws, fila, "Tabla 1 · por Fecha Pago (Monto = suma, Nº de cuotas = conteo de Real Pago Dolar)")
    fila += 1
    _encabezado(ws, fila, cols, etiqueta)
    fila += 1
    for d in dias:
        rojo = d.weekday() >= 5  # sábado/domingo
        for etiq, tabla in (("Monto", t_monto), ("Numero de cuotas", t_cuotas)):
            c = ws.cell(fila, 1, d.to_pydatetime())
            c.number_format = "dd/mm"
            c.alignment = Alignment(horizontal="center")
            c.font = Font(name=F, size=10, color=ROJO if rojo else "000000")
            c.border = BORDE
            b = ws.cell(fila, 2, etiq)
            b.font = Font(name=F, size=10)
            b.border = BORDE
            _fila_valores(ws, fila, tabla.loc[d], cols)
            fila += 1

    # ---- Tablas 2 y 3: por PERIODO
    for titulo_t, tm in (("Tabla 2 · por PERIODO · Monto (suma de Real Pago Dolar)", per_monto),
                         ("Tabla 3 · por PERIODO · Numero de cuotas (conteo de Real Pago Dolar)", per_cuotas)):
        fila += 1
        _titulo(ws, fila, titulo_t)
        fila += 1
        _encabezado(ws, fila, cols, etiqueta)
        fila += 1
        ws.cell(fila, 1).border = BORDE
        b = ws.cell(fila, 2, "TOTAL")
        b.font = Font(name=F, bold=True, size=10)
        b.border = BORDE
        _fila_valores(ws, fila, tm.sum(axis=0), cols, negrita=True, relleno=AZUL_CLARO)
        fila += 1
        for per in periodos:
            ws.cell(fila, 1).border = BORDE
            b = ws.cell(fila, 2, etiqueta_periodo(per))
            b.font = Font(name=F, size=10, color="7F7F7F")
            b.alignment = Alignment(horizontal="left", indent=1)
            b.border = BORDE
            _fila_valores(ws, fila, tm.loc[per], cols)
            fila += 1

    # ---- Nota (fuera de las tablas)
    fila += 1
    ws.cell(fila, 2, descripcion_ciclo(p)).font = Font(name=F, italic=True, size=9, color="7F7F7F")
    ws.cell(fila + 1, 2, f"Generado: {datetime.now():%d/%m/%Y %H:%M}").font = Font(
        name=F, italic=True, size=9, color="7F7F7F")

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 26
    for k, col in enumerate(cols, start=3):
        ws.column_dimensions[get_column_letter(k)].width = 12 if col.key in (ADEL, CIERRE, TOTAL) else 9.5


# ---------------------------------------------------------------- DETALLE
def escribir_detalle(wb: Workbook, df: pd.DataFrame, asesoras: list[dict], cols: list[Columna], p: Params) -> None:
    """Hoja con TODAS las filas consideradas en los totales, como Tabla de Excel (filtros incluidos).

    Sirve para auditar cualquier número (filtrando) o para insertar una tabla dinámica propia encima.
    """
    etiquetas = {c.key: c.label for c in cols}
    si_no = {True: "SI", False: "NO"}
    hoja_de = {a["nombre"]: a["hoja"] for a in asesoras}
    orden = {a["nombre"]: i for i, a in enumerate(asesoras)}
    d = (df.assign(_o=df["Coordinadora Cobranza"].map(orden))
           .sort_values(["_o", "Fecha Pago", "Fecha Vencimiento Actual"], kind="stable"))

    # categoría y motivo de cada pago (para auditar por qué es Adelanto, Adelanto Cierre u otro)
    categorias, motivos = [], []
    for adel, cierre, v, g, k in zip(d["es_adelanto"], d["es_cierre"], d["Fecha Vencimiento Actual"],
                                     d["Fecha Pago"], d["col_venc"]):
        if cierre and v.date() >= p.cierre_desde:
            categorias.append("Adelanto Cierre")
            motivos.append(f"Vence desde el {p.cierre_desde:%d/%m} (después del cierre del {p.fecha_cierre:%d/%m})")
        elif cierre:
            categorias.append("Adelanto Cierre")
            motivos.append(f"Vence del {p.adelanto_desde:%d/%m} al {p.fecha_cierre:%d/%m} y se pagó en {MESES[g.month - 1]} "
                           f"(mes anterior al del vencimiento)")
        elif adel:
            categorias.append("Adelanto (sin cierre)")
            motivos.append(f"Vence del {p.adelanto_desde:%d/%m} al {p.fecha_cierre:%d/%m} y se pagó en {MESES[g.month - 1]} "
                           f"(mismo mes del vencimiento)")
        else:
            categorias.append("Otros (no es adelanto)")
            motivos.append(f"Vence antes del {p.adelanto_desde:%d/%m} ({etiquetas[k]})")

    campos = [
        ("Hoja", d["Coordinadora Cobranza"].map(hoja_de)),
        ("Coordinadora Cobranza", d["Coordinadora Cobranza"]),
        *([("Coordinadora original", d["Coordinadora original"])] if "Coordinadora original" in d.columns else []),
        ("Fecha Pago", d["Fecha Pago"]),
        ("Fecha Vencimiento Actual", d["Fecha Vencimiento Actual"]),
        *([("Nro Cuota", d["Nro Cuota"])] if "Nro Cuota" in d.columns else []),
        ("PERIODO", d["per_key"].map(etiqueta_periodo)),
        ("Real Pago Dolar", d["Real Pago Dolar"]),
        ("Categoría", pd.Series(categorias, index=d.index)),
        ("Motivo", pd.Series(motivos, index=d.index)),
        ("Columna de vencimiento", d["col_venc"].map(etiquetas)),
        ("Adelanto", d["es_adelanto"].map(si_no)),
        ("Adelanto Cierre", d["es_cierre"].map(si_no)),
        (f"Vencido del 1 al {p.fecha_cierre.day} de {MESES[p.adelanto_desde.month - 1]}", d["venc_1_al_cierre"].map(si_no)),
        (f"Vencido desde el {p.cierre_desde.day} de {MESES[p.cierre_desde.month - 1]}", d["venc_desde_cierre"].map(si_no)),
    ]
    if "Código Matricula" in d.columns:
        campos.append(("Código Matricula", d["Código Matricula"]))
    campos.append(("Fila en BASE", d["Fila en BASE"]))

    ws = wb.create_sheet("DETALLE")
    ws.append([n for n, _ in campos])
    columnas_vals = [list(s) for _, s in campos]
    for fila in zip(*columnas_vals):
        out = []
        for v in fila:
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif pd.isna(v):
                v = None
            out.append(v)
        ws.append(out)

    n = len(d)
    ultima = get_column_letter(len(campos))
    tabla = Table(displayName="Detalle", ref=f"A1:{ultima}{n + 1}")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabla)
    ws.freeze_panes = "A2"

    nombres = [nm for nm, _ in campos]
    for j, nm in enumerate(nombres, start=1):
        letra = get_column_letter(j)
        ws.column_dimensions[letra].width = max(12, min(34, len(nm) + 4))
        fmt = {"Fecha Pago": "dd/mm/yyyy", "Fecha Vencimiento Actual": "dd/mm/yyyy",
               "Real Pago Dolar": "0.00"}.get(nm)
        for c in ws[letra][1:]:
            c.font = Font(name=F, size=10)
            if fmt:
                c.number_format = fmt
        ws.cell(1, j).font = Font(name=F, size=10, bold=True, color="FFFFFF")
    for nm, ancho in (("Coordinadora Cobranza", 38), ("Coordinadora original", 38), ("Categoría", 24), ("Motivo", 92)):
        if nm in nombres:
            ws.column_dimensions[get_column_letter(nombres.index(nm) + 1)].width = ancho


# ---------------------------------------------------------------- RESUMEN ADELANTOS
def escribir_resumen_adelantos(wb: Workbook, resumen: pd.DataFrame) -> None:
    """Hoja con una fila por asesora y los 4 montos de adelantos (segunda pestaña, tras RES MES)."""
    ws = wb.create_sheet("RESUMEN ADELANTOS", 1)
    azul = _fill(AZUL_CLARO)
    ws.cell(1, 1, "Asesora")
    for j, h in enumerate(resumen.columns, start=2):
        ws.cell(1, j, h)
    for j in range(1, len(resumen.columns) + 2):
        c = ws.cell(1, j)
        c.font = Font(name=F, bold=True, size=10, color="1F3864")
        c.fill, c.border = azul, BORDE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 62
    for i, (nombre, fila) in enumerate(resumen.iterrows(), start=2):
        es_tot = nombre == "Total"
        celdas = [nombre] + [round(float(x), 2) for x in fila]
        for j, v in enumerate(celdas, start=1):
            c = ws.cell(i, j, v)
            c.font = Font(name=F, size=10, bold=es_tot)
            c.border = BORDE
            if j > 1:
                c.number_format = "0"
                c.alignment = Alignment(horizontal="right")
            if es_tot:
                c.fill = azul
    ws.column_dimensions["A"].width = 16
    for j in range(2, len(resumen.columns) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 22
