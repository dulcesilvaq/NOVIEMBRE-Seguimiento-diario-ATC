"""Lectura y validación de la base."""
from __future__ import annotations

import re

import pandas as pd

REQUERIDAS = [
    "PERIODO", "Fecha Vencimiento Actual", "Fecha Pago",
    "Real Pago Dolar", "Coordinadora Cobranza",
]


# Opcionales (si existen): sirven para rastrear cada pago en la hoja DETALLE.
# No son datos personales (no se leen DNI, correos, teléfonos ni nombres de alumnos).
OPCIONALES = ["Código Matricula", "Código Matricula.1", "Nro Cuota"]


def _hoja_base(xl: pd.ExcelFile, hojas_posibles: list[str]) -> str:
    hoja = next((h for h in hojas_posibles if h in xl.sheet_names), None)
    if hoja is None:
        raise ValueError(
            f"No encontré ninguna de las hojas {hojas_posibles} en el archivo. "
            f"Hojas disponibles: {xl.sheet_names}"
        )
    return hoja


def listar_coordinadoras(ruta, hojas_posibles: list[str]) -> list[str]:
    """Nombres de coordinadoras presentes en la base (lee solo esa columna)."""
    if hasattr(ruta, "seek"):
        ruta.seek(0)
    xl = pd.ExcelFile(ruta)
    d = pd.read_excel(xl, sheet_name=_hoja_base(xl, hojas_posibles),
                      usecols=lambda c: str(c).strip() == "Coordinadora Cobranza")
    if hasattr(ruta, "seek"):
        ruta.seek(0)
    return sorted(d.iloc[:, 0].dropna().astype(str).str.strip().unique()) if d.shape[1] else []


def cargar_base(ruta, hojas_posibles: list[str], advertencias: list[str]) -> pd.DataFrame:
    if hasattr(ruta, "seek"):
        ruta.seek(0)
    xl = pd.ExcelFile(ruta)
    hoja = _hoja_base(xl, hojas_posibles)
    # Solo se leen las columnas necesarias: DNI, correos, teléfonos, etc. nunca se cargan.
    df = pd.read_excel(xl, sheet_name=hoja, usecols=lambda c: str(c).strip() in REQUERIDAS + OPCIONALES)
    df.columns = [str(c).strip() for c in df.columns]
    df["Fila en BASE"] = df.index + 2
    # Código de matrícula: si la primera columna no trae un código válido (ej. pegaron una fecha), se usa la segunda
    if "Código Matricula.1" in df.columns:
        patron = r"^\d+A\d+$"
        c1 = df.get("Código Matricula", pd.Series(index=df.index, dtype=object)).astype("string").str.strip()
        c2 = df["Código Matricula.1"].astype("string").str.strip()
        usar_c2 = ~c1.str.match(patron, na=False) & c2.str.match(patron, na=False)
        if usar_c2.any():
            advertencias.append(f"{int(usar_c2.sum())} fila(s) con 'Código Matricula' inválido (no es un código): "
                                f"se usó la segunda columna de código (filas de la hoja: "
                                f"{[int(x) for x in (df.loc[usar_c2, 'Fila en BASE']).head(8)]}).")
        df["Código Matricula"] = c1.where(~usar_c2, c2).astype(object)
        df = df.drop(columns=["Código Matricula.1"])          # número de fila en la hoja original (la fila 1 son los encabezados)

    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas obligatorias en la hoja '{hoja}': {faltan}")

    n0 = len(df)
    df["Real Pago Dolar"] = pd.to_numeric(df["Real Pago Dolar"], errors="coerce")
    for c in ("Fecha Pago", "Fecha Vencimiento Actual"):
        df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True).dt.normalize()
    # PERIODO: el texto "Anteriores" (como viene en algunas bases) equivale al grupo Anteriores (<= 2019)
    per_raw = df["PERIODO"].copy()
    es_ant = df["PERIODO"].astype("string").str.strip().str.lower().str.startswith("anterior", na=False)
    df["PERIODO"] = pd.to_numeric(df["PERIODO"], errors="coerce")
    no_anio = df["PERIODO"].notna() & ~(df["PERIODO"].between(1990, 2100) & (df["PERIODO"] % 1 == 0))
    df.loc[no_anio, "PERIODO"] = float("nan")          # ej. 282.02 (un monto pegado en esa columna): no es un año
    df.loc[es_ant, "PERIODO"] = 0
    if "Nro Cuota" in df.columns:
        df["Nro Cuota"] = df["Nro Cuota"].astype("string").str.strip().astype(object)
    df["Coordinadora Cobranza"] = df["Coordinadora Cobranza"].astype("string").str.strip()

    # Solo pagos reales (> 0)
    sin_pago = ~(df["Real Pago Dolar"] > 0)
    # Filas con datos clave vacíos (sin ellos no se pueden ubicar en las tablas): se excluyen y se avisa
    cols = ["Fecha Pago", "Fecha Vencimiento Actual", "Coordinadora Cobranza"]
    incompletas = (~sin_pago) & df[cols].isna().any(axis=1)
    if incompletas.any():
        det = {c: int(df.loc[incompletas, c].isna().sum()) for c in cols if df.loc[incompletas, c].isna().any()}
        advertencias.append(f"{int(incompletas.sum())} fila(s) con pago > 0 pero con datos vacíos {det}: "
                            f"se excluyeron (revisa la base).")
    # PERIODO vacío o no reconocido: el pago SÍ cuenta en RES MES, Tabla 1 y adelantos (no dependen del PERIODO),
    # pero no tiene fila en las tablas 2 y 3 (esas filas son fijas); se avisa
    sin_per = (~sin_pago) & (~incompletas) & df["PERIODO"].isna()
    if sin_per.any():
        raros = per_raw[sin_per].dropna().astype(str).unique()[:5]
        extra = f" Valores no reconocidos: {list(raros)}." if len(raros) else ""
        advertencias.append(f"{int(sin_per.sum())} pago(s) (US$ {df.loc[sin_per, 'Real Pago Dolar'].sum():,.2f}) con PERIODO "
                            f"vacío o no reconocido: cuentan en RES MES, Tabla 1 y adelantos, pero NO en las tablas 2 y 3 "
                            f"(sus filas de PERIODO son fijas). Corrige el PERIODO en la base si los necesitas ahí.{extra}")
        df.loc[sin_per, "PERIODO"] = -1
    df = df[~sin_pago & ~incompletas].copy()
    excl = n0 - len(df) - int(incompletas.sum())
    if excl:
        advertencias.append(f"{excl} fila(s) con Real Pago Dolar vacío o <= 0: no se cuentan.")

    if df.empty:
        raise ValueError("La base no tiene pagos con 'Real Pago Dolar' mayor a 0 (revisa la hoja y las columnas).")
    df["PERIODO"] = df["PERIODO"].astype(int)
    df["Coordinadora Cobranza"] = df["Coordinadora Cobranza"].astype(str)
    return df.reset_index(drop=True)
